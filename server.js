'use strict';

const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { Pool } = require('pg');
const dotenv = require('dotenv');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

dotenv.config({ path: path.join(__dirname, '.env'), override: true });

const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const MAX_BYTES = Number(process.env.UPLOAD_MAX_BYTES) || 52 * 1024 * 1024;

function buildHeaderNames(headerRow) {
  const count = {};
  const names = [];
  const len = headerRow ? headerRow.length : 0;
  for (let j = 0; j < len; j++) {
    let base =
      headerRow[j] == null || headerRow[j] === ''
        ? `Column_${j + 1}`
        : String(headerRow[j]).trim();
    if (!base) base = `Column_${j + 1}`;
    const n = (count[base] = (count[base] || 0) + 1);
    names.push(n === 1 ? base : `${base}_${n}`);
  }
  return names;
}

function rowToNamedObject(row, headerNames) {
  const o = {};
  for (let j = 0; j < headerNames.length; j++) {
    o[headerNames[j]] =
      row[j] !== undefined && row[j] !== null ? row[j] : '';
  }
  return o;
}

function parseUploadBuffer(originalName, buffer, mimeType) {
  const lower = (originalName || '').toLowerCase();
  const isCsv =
    lower.endsWith('.csv') ||
    mimeType === 'text/csv' ||
    mimeType === 'application/csv';

  let wb;
  if (isCsv) {
    const text = buffer.toString('utf8');
    wb = XLSX.read(text, { type: 'string', raw: false });
  } else {
    wb = XLSX.read(buffer, { type: 'buffer' });
  }

  const sheetNames = wb.SheetNames || [];
  if (sheetNames.length === 0) {
    throw new Error('Workbook has no sheets');
  }

  const firstName = sheetNames[0];
  const sheet = wb.Sheets[firstName];
  const allData = XLSX.utils.sheet_to_row_object_array(sheet, { header: 1 });

  if (!allData.length) {
    return {
      sheetNames,
      firstSheetName: firstName,
      rows: [],
      headerNames: [],
    };
  }

  const headerRow = allData[0];
  const headerNames = buildHeaderNames(headerRow);
  const rows = [];
  for (let i = 1; i < allData.length; i++) {
    rows.push(rowToNamedObject(allData[i], headerNames));
  }

  return {
    sheetNames,
    firstSheetName: firstName,
    rows,
    headerNames,
  };
}

/**
 * TLS for managed database hosts (e.g. Aiven). Uses pattern checks instead of `new URL()`
 * so passwords with @, #, etc. do not break parsing and leave ssl unset (strict verify → crash).
 */
function pgSslOption(connectionString) {
  if (!connectionString) return undefined;
  const lower = connectionString.toLowerCase();
  const managed =
    lower.includes('aivencloud.com') ||
    lower.includes('amazonaws.com');
  const sslMode =
    /[?&]sslmode=(require|verify-ca|verify-full)(?:&|$)/i.test(connectionString);
  if (managed || sslMode) {
    return { rejectUnauthorized: false };
  }
  return undefined;
}

/** Remove sslmode from URI so pg-connection-string does not force verify-full over our ssl object. */
function stripSslModeFromConnectionString(conn) {
  const i = conn.indexOf('?');
  if (i < 0) return conn;
  const base = conn.slice(0, i);
  const qs = conn.slice(i + 1);
  const params = qs.split('&').filter((p) => p && !/^sslmode=/i.test(p));
  return params.length ? `${base}?${params.join('&')}` : base;
}

let pool = null;

function getPool() {
  const conn = (process.env.DATABASE_URL || '').trim();
  if (!conn) return null;
  if (!pool) {
    const ssl = pgSslOption(conn);
    pool = new Pool({
      connectionString: ssl ? stripSslModeFromConnectionString(conn) : conn,
      ssl: ssl,
      max: 10,
    });
  }
  return pool;
}

async function ensureSchema(client) {
  const schemaPath = path.join(ROOT, 'db', 'schema.sql');
  const sql = fs.readFileSync(schemaPath, 'utf8');
  await client.query(sql);
}

function headerIndex(headers, candidates) {
  if (!headers || !headers.length) return -1;
  const lower = candidates.map((c) => c.toLowerCase());
  for (let j = 0; j < headers.length; j++) {
    const h = headers[j] == null ? '' : String(headers[j]).trim().toLowerCase();
    if (!h) continue;
    for (const c of lower) {
      if (h === c || h.includes(c)) return j;
    }
  }
  return -1;
}

/**
 * Jira / export "Issue key" column — must NOT use substring `key` (matches "Parent Key", etc.).
 * Mirrors client findIssueKeyColumnIndex: exact names first, then safe multi-word substrings only.
 */
function headerIndexIssueKey(headers) {
  if (!headers || !headers.length) return -1;
  const exactCandidates = [
    'issue key',
    'key',
    'issue id',
    'jira key',
    'issuekey',
    'work item key',
  ];
  const lowerHeaders = [];
  for (let j = 0; j < headers.length; j++) {
    lowerHeaders.push(
      String(headers[j] == null ? '' : headers[j]).trim().toLowerCase()
    );
  }
  for (const cand of exactCandidates) {
    const idx = lowerHeaders.indexOf(cand);
    if (idx >= 0) return idx;
  }
  const substringNeedles = ['issue key', 'jira key', 'issue id', 'work item key'];
  for (let j = 0; j < lowerHeaders.length; j++) {
    const h = lowerHeaders[j];
    if (!h) continue;
    for (const needle of substringNeedles) {
      if (h.includes(needle)) return j;
    }
  }
  return -1;
}

function normalizeIssueKeyDedupe(ik) {
  return String(ik || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[\u2010-\u2015\u2212]/g, '-')
    .toUpperCase();
}

/** Typical Jira-style issue keys (PROJ-123); avoids matching sprint numbers like 14.1 */
function looksLikeJiraIssueKey(val) {
  const t = String(val || '').trim();
  return /^[A-Za-z][A-Za-z0-9]{1,15}-\d+$/i.test(t);
}

/**
 * When the primary issue-key column is empty or mis-detected, infer key from likely columns or cell values.
 */
function fallbackIssueKeyFromCells(cells, headers) {
  if (!cells || typeof cells !== 'object') return '';
  if (headers && headers.length) {
    for (let j = 0; j < headers.length; j++) {
      const hn = String(headers[j] == null ? '' : headers[j]).trim().toLowerCase();
      if (
        (hn.includes('issue') && hn.includes('key')) ||
        hn.includes('jira key') ||
        hn.includes('issue id') ||
        hn.includes('work item key')
      ) {
        const v = pickCell(cells, headers, j);
        if (looksLikeJiraIssueKey(v)) return normalizeIssueKeyDedupe(v);
      }
    }
  }
  for (const k of Object.keys(cells)) {
    const v = cells[k];
    if (looksLikeJiraIssueKey(v)) return normalizeIssueKeyDedupe(v);
  }
  return '';
}

function pickCell(cells, headers, idx) {
  if (idx < 0 || !headers[idx]) return '';
  const key = headers[idx];
  const v = cells[key];
  if (v === undefined || v === null) return '';
  return String(v).trim();
}

/** Headers like "Assigned To" (ADO) do not match substring `assignee`; scan assignee-like columns. */
function headerIndexAssignee(headers) {
  return headerIndex(headers, [
    'assignee',
    'assignee name',
    'assignee_name',
    'assigned to',
  ]);
}

function pickAssigneeFromRow(cells, headers, assigneeIdx) {
  if (cells && typeof cells === 'object') {
    if (assigneeIdx >= 0) {
      const v = pickCell(cells, headers, assigneeIdx);
      if (v) return v;
    }
    if (headers && headers.length) {
      for (let j = 0; j < headers.length; j++) {
        const hn = String(headers[j] == null ? '' : headers[j]).trim().toLowerCase();
        if (
          hn === 'assignee' ||
          hn === 'assigned to' ||
          hn.includes('assignee') ||
          hn.includes('assigned to')
        ) {
          const v = pickCell(cells, headers, j);
          if (v) return v;
        }
      }
    }
  }
  return '';
}

const SPRINT_THRESHOLD_EPS = 1e-6;

function sprintColumnIndicesFromHeaders(headers) {
  if (!headers || !headers.length) return [];
  const idx = [];
  for (let j = 5; j < headers.length; j++) {
    const h = headers[j] == null ? '' : String(headers[j]).trim();
    if (h === 'Sprint' || /^Sprint_[0-9]+$/.test(h)) {
      idx.push(j);
    }
  }
  return idx;
}

function sprintCellMatchesThreshold(cellValue, thresholdStr) {
  const thresholdNum = parseFloat(thresholdStr);
  if (Number.isNaN(thresholdNum)) return false;
  if (cellValue === null || cellValue === undefined || cellValue === '') return false;
  const numValue = parseFloat(cellValue);
  if (Number.isNaN(numValue)) return false;
  return Math.abs(numValue - thresholdNum) < SPRINT_THRESHOLD_EPS;
}

function rowHasSpilloverRaw(cells, headers, sprintIndices, thresholdStr) {
  if (!thresholdStr || String(thresholdStr).trim() === '') return false;
  if (Number.isNaN(parseFloat(thresholdStr))) return false;
  for (const j of sprintIndices) {
    const key = headers[j];
    const v = cells[key];
    if (sprintCellMatchesThreshold(v, thresholdStr)) return true;
  }
  return false;
}

function isBugRowCells(cells, headers, issueTypeIdx) {
  if (issueTypeIdx < 0) return false;
  const v = pickCell(cells, headers, issueTypeIdx);
  if (!v) return false;
  return String(v).trim().toLowerCase() === 'bug';
}

async function computePeopleInsights(client, uploadId, thresholdStr) {
  const meta = await client.query(
    `SELECT column_headers, original_name FROM file_uploads WHERE id = $1`,
    [uploadId]
  );
  if (!meta.rows.length) return null;
  let headers = meta.rows[0].column_headers || [];
  const originalName = meta.rows[0].original_name;

  const dataRows = await client.query(
    `SELECT cells FROM file_upload_rows WHERE upload_id = $1 ORDER BY row_index ASC`,
    [uploadId]
  );

  if (!headers.length && dataRows.rows.length) {
    headers = Object.keys(dataRows.rows[0].cells || {});
  }

  const keyIdx = headerIndexIssueKey(headers);
  const assigneeIdx = headerIndexAssignee(headers);
  const issueTypeIdx = headerIndex(headers, ['issue type']);
  const sprintIdxs = sprintColumnIndicesFromHeaders(headers);

  const spillMap = new Map();
  const bugMap = new Map();

  const spillIssueKeys = new Set();
  const bugIssueKeys = new Set();

  for (const row of dataRows.rows) {
    const cells = row.cells || {};
    const ikRaw = keyIdx >= 0 ? pickCell(cells, headers, keyIdx) : '';
    const ik = normalizeIssueKeyDedupe(ikRaw);
    if (!ik) continue;

    let assignee = pickAssigneeFromRow(cells, headers, assigneeIdx);
    if (!assignee) assignee = '(Unassigned)';

    if (rowHasSpilloverRaw(cells, headers, sprintIdxs, thresholdStr)) {
      spillIssueKeys.add(ik);
      if (!spillMap.has(assignee)) spillMap.set(assignee, new Set());
      spillMap.get(assignee).add(ik);
    }

    if (isBugRowCells(cells, headers, issueTypeIdx)) {
      bugIssueKeys.add(ik);
      if (!bugMap.has(assignee)) bugMap.set(assignee, new Set());
      bugMap.get(assignee).add(ik);
    }
  }

  function mapToList(m) {
    return Array.from(m.entries())
      .map(([name, keys]) => ({
        assignee: name,
        issueKeys: Array.from(keys).sort(),
        count: keys.size,
      }))
      .sort((a, b) => a.assignee.localeCompare(b.assignee));
  }

  return {
    originalName,
    sprintColumnCount: sprintIdxs.length,
    thresholdUsed: thresholdStr ? String(thresholdStr).trim() : '',
    spilloverAssignees: mapToList(spillMap),
    bugAssignees: mapToList(bugMap),
    totals: {
      spilloverIssues: spillIssueKeys.size,
      bugIssues: bugIssueKeys.size,
    },
  };
}

/** First-seen assignee display name per normalized issue key (for contributor activity UI). */
async function assigneeByIssueKeyMap(client, uploadId) {
  const meta = await client.query(
    `SELECT column_headers FROM file_uploads WHERE id = $1`,
    [uploadId]
  );
  const out = new Map();
  if (!meta.rows.length) return out;
  let headers = meta.rows[0].column_headers || [];
  const dataRows = await client.query(
    `SELECT cells FROM file_upload_rows WHERE upload_id = $1 ORDER BY row_index ASC`,
    [uploadId]
  );
  if (!headers.length && dataRows.rows.length) {
    headers = Object.keys(dataRows.rows[0].cells || {});
  }
  let keyIdx = headerIndexIssueKey(headers);
  if (keyIdx < 0) {
    keyIdx = headerIndex(headers, ['issue key', 'issue id', 'work item key']);
  }
  const assigneeIdx = headerIndexAssignee(headers);

  for (const row of dataRows.rows) {
    const cells = row.cells || {};
    let ikRaw = keyIdx >= 0 ? pickCell(cells, headers, keyIdx) : '';
    let ik = normalizeIssueKeyDedupe(ikRaw);
    if (!ik) ik = fallbackIssueKeyFromCells(cells, headers);
    if (!ik || out.has(ik)) continue;
    let assignee = pickAssigneeFromRow(cells, headers, assigneeIdx);
    if (!assignee) assignee = '(Unassigned)';
    out.set(ik, assignee);
  }
  return out;
}

function contributorEditKindsFromRow(row) {
  const sr = String(row.spillover_reason || '').trim();
  const sc = String(row.spillover_category || '').trim();
  const pr = String(row.prod || '').trim();
  const rc = String(row.rca || '').trim();
  const kinds = [];
  if (sr !== '' || sc !== '') kinds.push('spillover');
  if (pr !== '' || rc !== '') kinds.push('bug');
  return kinds;
}

async function insertRowsChunked(client, uploadId, rowObjects) {
  const chunkSize = 150;
  for (let offset = 0; offset < rowObjects.length; offset += chunkSize) {
    const chunk = rowObjects.slice(offset, offset + chunkSize);
    const nums = [];
    const params = [uploadId];
    let p = 2;
    for (let i = 0; i < chunk.length; i++) {
      const rowIndex = offset + i;
      nums.push(`($1, $${p}, $${p + 1}::jsonb)`);
      params.push(rowIndex, JSON.stringify(chunk[i]));
      p += 2;
    }
    await client.query(
      `INSERT INTO file_upload_rows (upload_id, row_index, cells) VALUES ${nums.join(', ')}`,
      params
    );
  }
}

async function fetchInviteOr404(client, token) {
  const r = await client.query(
    `SELECT i.*, f.original_name
     FROM invites i
     JOIN file_uploads f ON f.id = i.upload_id
     WHERE i.token = $1::uuid`,
    [token]
  );
  if (!r.rows.length) return null;
  const inv = r.rows[0];
  if (inv.expires_at && new Date(inv.expires_at) < new Date()) {
    return { expired: true };
  }
  return { invite: inv };
}

function normalizeAssigneeForMatch(s) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/** When filter_by_assignee is true: exact match on assignee_scope if set; else substring match of assignee cell against invitee_email. */
function inviteAssigneeAllowsRow(invite, assigneeRaw) {
  if (!invite.filter_by_assignee) return true;
  const scope =
    invite.assignee_scope != null ? String(invite.assignee_scope).trim() : '';
  if (scope) {
    return (
      normalizeAssigneeForMatch(assigneeRaw) === normalizeAssigneeForMatch(scope)
    );
  }
  const emailPart =
    invite.invitee_email != null ? String(invite.invitee_email).trim() : '';
  if (emailPart) {
    return assigneeRaw.toLowerCase().indexOf(emailPart.toLowerCase()) !== -1;
  }
  return true;
}

async function buildIssuesForInvite(client, invite) {
  const uploadId = invite.upload_id;
  const meta = await client.query(
    `SELECT column_headers FROM file_uploads WHERE id = $1`,
    [uploadId]
  );
  if (!meta.rows.length) return { issues: [], headers: [] };
  let headers = meta.rows[0].column_headers || [];
  const dataRows = await client.query(
    `SELECT row_index, cells FROM file_upload_rows WHERE upload_id = $1 ORDER BY row_index ASC`,
    [uploadId]
  );

  if (!headers.length && dataRows.rows.length) {
    const first = dataRows.rows[0].cells || {};
    headers = Object.keys(first);
  }

  const keyIdx = headerIndexIssueKey(headers);
  const assigneeIdx = headerIndexAssignee(headers);
  const summaryIdx = headerIndex(headers, ['summary', 'issue summary', 'title']);
  const issueTypeIdx = headerIndex(headers, ['issue type']);
  const sprintIdxs = sprintColumnIndicesFromHeaders(headers);

  const thresholdRaw =
    invite.sprint_threshold != null ? String(invite.sprint_threshold).trim() : '';
  const useWorkloadFilter =
    thresholdRaw !== '' && !Number.isNaN(parseFloat(thresholdRaw));

  /** Issue keys that qualify for this invite: spillover at threshold and/or Bug row (matches Send invite lists). */
  const allowedDedupeKeys = new Set();
  const keyHasSpillover = new Map();
  const keyHasBug = new Map();
  if (useWorkloadFilter) {
    for (const row of dataRows.rows) {
      const cells = row.cells || {};
      const assignee = pickAssigneeFromRow(cells, headers, assigneeIdx);
      if (!inviteAssigneeAllowsRow(invite, assignee)) continue;
      const ikRaw = keyIdx >= 0 ? pickCell(cells, headers, keyIdx) : '';
      const ikDedupe = normalizeIssueKeyDedupe(ikRaw);
      if (!ikDedupe) continue;
      const spill = rowHasSpilloverRaw(
        cells,
        headers,
        sprintIdxs,
        thresholdRaw
      );
      const bug = isBugRowCells(cells, headers, issueTypeIdx);
      if (spill) {
        allowedDedupeKeys.add(ikDedupe);
        keyHasSpillover.set(ikDedupe, true);
      }
      if (bug) {
        allowedDedupeKeys.add(ikDedupe);
        keyHasBug.set(ikDedupe, true);
      }
    }
  }

  const issues = [];
  const seen = new Set();
  for (const row of dataRows.rows) {
    const cells = row.cells || {};
    const ikRaw = keyIdx >= 0 ? pickCell(cells, headers, keyIdx) : '';
    const ikDedupe = normalizeIssueKeyDedupe(ikRaw);
    if (!ikRaw || !ikDedupe || seen.has(ikDedupe)) continue;
    const assignee = pickAssigneeFromRow(cells, headers, assigneeIdx);
    if (!inviteAssigneeAllowsRow(invite, assignee)) continue;
    if (useWorkloadFilter && !allowedDedupeKeys.has(ikDedupe)) continue;
    seen.add(ikDedupe);
    const summary =
      summaryIdx >= 0 ? pickCell(cells, headers, summaryIdx) : '';
    const representativeBug = isBugRowCells(cells, headers, issueTypeIdx);
    let needsSpilloverFields;
    let needsBugFields;
    if (useWorkloadFilter) {
      needsSpilloverFields = keyHasSpillover.get(ikDedupe) === true;
      needsBugFields = keyHasBug.get(ikDedupe) === true;
    } else {
      // Invite created without sprint_threshold: use Issue Type on this row — Bug → PROD/RCA only; anything else → spillover fields only (no bug column for spillover-only people).
      needsBugFields = representativeBug;
      needsSpilloverFields = !representativeBug;
    }
    issues.push({
      issueKey: String(ikRaw).trim(),
      assignee,
      summary,
      rowIndex: row.row_index,
      needsSpilloverFields,
      needsBugFields,
    });
  }

  return { issues, headers };
}

function spilloverFieldsComplete(edit) {
  if (!edit) return false;
  const sr = String(edit.spillover_reason || '').trim();
  const sc = String(edit.spillover_category || '').trim();
  return sr !== '' && sc !== '';
}

function bugFieldsComplete(edit) {
  if (!edit) return false;
  const pr = String(edit.prod || '').trim();
  const rc = String(edit.rca || '').trim();
  return pr !== '' && rc !== '';
}

/**
 * Per Assignee from bulk invites: spillover vs bug submission vs workload (same rules as contribute workspace).
 */
async function assigneeInviteSubmissionStatuses(client, uploadId, thresholdStr, assigneeNames) {
  const editsRes = await client.query(
    `SELECT issue_key, spillover_reason, spillover_category, prod, rca FROM issue_field_edits WHERE upload_id = $1`,
    [uploadId]
  );
  const editsByIk = new Map();
  for (const row of editsRes.rows) {
    editsByIk.set(normalizeIssueKeyDedupe(row.issue_key), row);
  }

  const thresholdRaw =
    thresholdStr != null && String(thresholdStr).trim() !== ''
      ? String(thresholdStr).trim()
      : null;

  const out = [];
  for (const assigneeName of assigneeNames) {
    const invite = {
      upload_id: uploadId,
      sprint_threshold: thresholdRaw,
      filter_by_assignee: true,
      assignee_scope: assigneeName,
      invitee_email: null,
    };
    const { issues } = await buildIssuesForInvite(client, invite);
    let needsSpill = false;
    let needsBug = false;
    let spillOk = true;
    let bugOk = true;
    for (const iss of issues) {
      const ik = normalizeIssueKeyDedupe(iss.issueKey);
      const edit = editsByIk.get(ik);
      if (iss.needsSpilloverFields) {
        needsSpill = true;
        if (!spilloverFieldsComplete(edit)) spillOk = false;
      }
      if (iss.needsBugFields) {
        needsBug = true;
        if (!bugFieldsComplete(edit)) bugOk = false;
      }
    }

    let spilloverState = 'none';
    let bugState = 'none';
    if (needsSpill) spilloverState = spillOk ? 'complete' : 'pending';
    if (needsBug) bugState = bugOk ? 'complete' : 'pending';

    let statusLabel = '—';
    if (!issues.length) {
      statusLabel = 'No issues on link';
    } else if (needsSpill && needsBug) {
      if (spillOk && bugOk) statusLabel = 'Submitted';
      else if (spillOk && !bugOk) statusLabel = 'Spillover submitted · Bug pending';
      else if (!spillOk && bugOk) statusLabel = 'Spillover pending · Bug submitted';
      else statusLabel = 'Spillover pending · Bug pending';
    } else if (needsSpill) {
      statusLabel = spillOk ? 'Submitted' : 'Spillover pending';
    } else if (needsBug) {
      statusLabel = bugOk ? 'Submitted' : 'Bug pending';
    }

    out.push({
      assignee: assigneeName,
      statusLabel,
      spillover: spilloverState,
      bug: bugState,
    });
  }
  return out;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_BYTES },
});

const app = express();
app.set('trust proxy', 1);
app.use(cors());
app.use(express.json());

app.get('/api/health', (req, res) => {
  const p = getPool();
  res.json({
    ok: true,
    databaseConfigured: Boolean(p),
  });
});

/** Deletes every upload; cascades to rows, invites, issue_field_edits. Requires explicit confirmation. */
app.post('/api/database/clear-all', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }

  const body = req.body || {};
  if (body.confirm !== true) {
    return res.status(400).json({
      error: 'Confirmation required: send JSON body { "confirm": true }.',
    });
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const r = await client.query('DELETE FROM file_uploads');
    res.json({
      ok: true,
      deletedUploads: r.rowCount,
      message:
        'All uploads and related rows (invites, issue edits) were removed.',
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error:
        process.env.NODE_ENV === 'development'
          ? String(e.message || e)
          : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.post('/api/upload', upload.single('file'), async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }

  if (!req.file || !req.file.buffer) {
    return res.status(400).json({ error: 'Missing file field "file"' });
  }

  const originalName = req.file.originalname || 'upload';
  let parsed;
  try {
    parsed = parseUploadBuffer(
      originalName,
      req.file.buffer,
      req.file.mimetype || ''
    );
  } catch (e) {
    return res.status(400).json({ error: String(e.message || e) });
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);

    await client.query('BEGIN');

    const ins = await client.query(
      `INSERT INTO file_uploads
        (original_name, content_type, byte_size, row_count, sheet_names, first_sheet_name, column_headers)
       VALUES ($1, $2, $3, $4, $5, $6, $7)
       RETURNING id`,
      [
        originalName,
        req.file.mimetype || null,
        req.file.size,
        parsed.rows.length,
        parsed.sheetNames,
        parsed.firstSheetName || null,
        parsed.headerNames || [],
      ]
    );

    const uploadId = ins.rows[0].id;
    await insertRowsChunked(client, uploadId, parsed.rows);

    await client.query('COMMIT');

    res.json({
      ok: true,
      uploadId: String(uploadId),
      rowCount: parsed.rows.length,
      firstSheet: parsed.firstSheetName,
    });
  } catch (e) {
    await client.query('ROLLBACK');
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.get('/api/uploads', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const client = await p.connect();
  try {
    await ensureSchema(client);
    const r = await client.query(
      `SELECT id, original_name, created_at, row_count, first_sheet_name
       FROM file_uploads
       ORDER BY id DESC
       LIMIT 100`
    );
    res.json({
      uploads: r.rows.map((row) => ({
        id: String(row.id),
        originalName: row.original_name,
        createdAt: row.created_at,
        rowCount: row.row_count,
        firstSheetName: row.first_sheet_name,
      })),
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.get('/api/uploads/:id/data', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const id = req.params.id;
  if (!/^\d+$/.test(id)) {
    return res.status(400).json({ error: 'Invalid upload id' });
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const meta = await client.query(
      `SELECT id, original_name, sheet_names, first_sheet_name, column_headers, row_count
       FROM file_uploads WHERE id = $1`,
      [id]
    );
    if (!meta.rows.length) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    const m = meta.rows[0];
    const dataRows = await client.query(
      `SELECT cells FROM file_upload_rows WHERE upload_id = $1 ORDER BY row_index ASC`,
      [id]
    );

    let headers = m.column_headers;
    if (!headers || !headers.length) {
      const first = dataRows.rows[0];
      if (first && first.cells && typeof first.cells === 'object') {
        headers = Object.keys(first.cells);
      } else {
        headers = [];
      }
    }

    const rows = dataRows.rows.map((r) => r.cells);
    res.json({
      ok: true,
      uploadId: String(m.id),
      originalName: m.original_name,
      sheetNames: m.sheet_names || [],
      firstSheetName: m.first_sheet_name || 'Sheet1',
      headers,
      rows,
      rowCount: rows.length,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.get('/api/uploads/:id/issue-edits', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const id = req.params.id;
  if (!/^\d+$/.test(id)) {
    return res.status(400).json({ error: 'Invalid upload id' });
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const r = await client.query(
      `SELECT issue_key, spillover_reason, spillover_category, prod, rca, updated_at
       FROM issue_field_edits WHERE upload_id = $1
       ORDER BY issue_key ASC`,
      [id]
    );
    res.json({ ok: true, edits: r.rows });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.get('/api/uploads/:id/issue-edits/recent', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const id = req.params.id;
  if (!/^\d+$/.test(id)) {
    return res.status(400).json({ error: 'Invalid upload id' });
  }

  let limit = parseInt(String(req.query.limit != null ? req.query.limit : '20'), 10);
  if (!Number.isFinite(limit) || limit < 1) limit = 20;
  if (limit > 100) limit = 100;

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const check = await client.query(`SELECT id FROM file_uploads WHERE id = $1`, [id]);
    if (!check.rows.length) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    const r = await client.query(
      `SELECT issue_key, spillover_reason, spillover_category, prod, rca, updated_at
       FROM issue_field_edits WHERE upload_id = $1
       ORDER BY updated_at DESC NULLS LAST
       LIMIT $2`,
      [id, limit]
    );
    const assigneeMap = await assigneeByIssueKeyMap(client, id);
    const edits = r.rows.map((row) => {
      const ik = normalizeIssueKeyDedupe(row.issue_key);
      const kinds = contributorEditKindsFromRow(row);
      return {
        issue_key: row.issue_key,
        assignee_name:
          ik && assigneeMap.has(ik) ? assigneeMap.get(ik) : '—',
        update_kinds: kinds,
        updated_at: row.updated_at,
      };
    });
    res.json({ ok: true, edits });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.get('/api/uploads/:id/people-insights', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const id = req.params.id;
  if (!/^\d+$/.test(id)) {
    return res.status(400).json({ error: 'Invalid upload id' });
  }
  const threshold = req.query.threshold != null ? String(req.query.threshold).trim() : '';

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const check = await client.query(`SELECT id FROM file_uploads WHERE id = $1`, [id]);
    if (!check.rows.length) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    const insights = await computePeopleInsights(client, id, threshold);
    if (!insights) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    res.json({
      ok: true,
      uploadId: id,
      ...insights,
      spilloverHelp:
        insights.sprintColumnCount === 0
          ? 'No Sprint columns detected from column F onward (same rule as the main app).'
          : !threshold
            ? 'Enter sprint threshold to see who has spillover (same number as Upload & configure).'
            : null,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

/** POST body: { assignees: string[], threshold?: string } — spillover/bug completion vs bulk-invite workload. */
app.post('/api/uploads/:id/invite-submit-status', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const id = req.params.id;
  if (!/^\d+$/.test(id)) {
    return res.status(400).json({ error: 'Invalid upload id' });
  }

  const body = req.body || {};
  const rawList = body.assignees;
  if (!Array.isArray(rawList) || rawList.length === 0) {
    return res.status(400).json({ error: 'assignees must be a non-empty array' });
  }

  const assignees = [];
  const seen = new Set();
  for (const a of rawList) {
    const name = String(a || '').trim();
    if (!name || seen.has(name)) continue;
    seen.add(name);
    assignees.push(name);
  }
  if (!assignees.length) {
    return res.status(400).json({ error: 'No valid assignee names' });
  }

  const threshold =
    body.threshold != null ? String(body.threshold).trim() : '';

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const check = await client.query(`SELECT id FROM file_uploads WHERE id = $1`, [id]);
    if (!check.rows.length) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    const statuses = await assigneeInviteSubmissionStatuses(client, id, threshold, assignees);
    res.json({ ok: true, uploadId: id, statuses });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.post('/api/invites', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }

  const body = req.body || {};
  const uploadId = body.uploadId;
  if (!uploadId || !/^\d+$/.test(String(uploadId))) {
    return res.status(400).json({ error: 'uploadId required (numeric)' });
  }

  const token = crypto.randomUUID();
  const days = Math.min(365, Math.max(1, parseInt(body.expiresInDays, 10) || 30));
  const expires = new Date();
  expires.setDate(expires.getDate() + days);

  const assigneeScopeRaw =
    body.assigneeScope != null ? String(body.assigneeScope).trim() : '';
  const assigneeScope = assigneeScopeRaw || null;
  let filterByAssignee = Boolean(body.filterByAssignee);
  if (assigneeScope) filterByAssignee = true;
  if (filterByAssignee && !assigneeScope) {
    const em = body.inviteeEmail != null ? String(body.inviteeEmail).trim() : '';
    if (!em) {
      return res.status(400).json({
        error:
          'Assignee filtering is on but no exact assignee name was given. Enter a contributor email (Assignee column substring match) or enter an exact assignee name.',
      });
    }
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const check = await client.query(`SELECT id FROM file_uploads WHERE id = $1`, [uploadId]);
    if (!check.rows.length) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    const sprintThreshold =
      body.threshold != null && String(body.threshold).trim() !== ''
        ? String(body.threshold).trim()
        : body.sprintThreshold != null && String(body.sprintThreshold).trim() !== ''
          ? String(body.sprintThreshold).trim()
          : null;

    await client.query(
      `INSERT INTO invites (token, upload_id, invitee_email, filter_by_assignee, assignee_scope, note, expires_at, sprint_threshold)
       VALUES ($1::uuid, $2, $3, $4, $5, $6, $7, $8)`,
      [
        token,
        uploadId,
        body.inviteeEmail ? String(body.inviteeEmail).trim() : null,
        filterByAssignee,
        assigneeScope,
        body.note ? String(body.note).trim() : null,
        expires,
        sprintThreshold,
      ]
    );

    const base = process.env.PUBLIC_APP_URL || `${req.protocol}://${req.get('host')}`;
    const baseClean = base.replace(/\/$/, '');
    const inviteUrl = `${baseClean}/user-dashboard.html?token=${encodeURIComponent(token)}&api=${encodeURIComponent(baseClean)}`;

    res.json({
      ok: true,
      token,
      inviteUrl,
      expiresAt: expires.toISOString(),
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.post('/api/invites/bulk', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }

  const body = req.body || {};
  const uploadId = body.uploadId;
  if (!uploadId || !/^\d+$/.test(String(uploadId))) {
    return res.status(400).json({ error: 'uploadId required (numeric)' });
  }

  const rawList = body.assignees;
  if (!Array.isArray(rawList) || rawList.length === 0) {
    return res.status(400).json({ error: 'assignees must be a non-empty array of names' });
  }

  const seen = new Set();
  const assignees = [];
  for (const a of rawList) {
    const name = String(a || '').trim();
    if (!name || seen.has(name)) continue;
    seen.add(name);
    assignees.push(name);
  }
  if (!assignees.length) {
    return res.status(400).json({ error: 'No valid assignee names after deduplication' });
  }

  const days = Math.min(365, Math.max(1, parseInt(body.expiresInDays, 10) || 30));
  const expires = new Date();
  expires.setDate(expires.getDate() + days);
  const note = body.note != null ? String(body.note).trim() || null : null;
  const sprintThreshold =
    body.threshold != null && String(body.threshold).trim() !== ''
      ? String(body.threshold).trim()
      : null;

  const emailMap =
    body.assigneeEmails && typeof body.assigneeEmails === 'object' && !Array.isArray(body.assigneeEmails)
      ? body.assigneeEmails
      : {};

  const base = process.env.PUBLIC_APP_URL || `${req.protocol}://${req.get('host')}`;
  const baseClean = base.replace(/\/$/, '');

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const check = await client.query(`SELECT id FROM file_uploads WHERE id = $1`, [uploadId]);
    if (!check.rows.length) {
      return res.status(404).json({ error: 'Upload not found' });
    }

    await client.query('BEGIN');
    const invites = [];
    try {
      for (const assigneeName of assignees) {
        const token = crypto.randomUUID();
        let inviteeEmail = null;
        if (Object.prototype.hasOwnProperty.call(emailMap, assigneeName)) {
          const em = String(emailMap[assigneeName] ?? '').trim();
          if (em) inviteeEmail = em;
        }
        await client.query(
          `INSERT INTO invites (token, upload_id, invitee_email, filter_by_assignee, assignee_scope, note, expires_at, sprint_threshold)
           VALUES ($1::uuid, $2, $3, true, $4, $5, $6, $7)`,
          [token, uploadId, inviteeEmail, assigneeName, note, expires, sprintThreshold]
        );
        invites.push({
          assignee: assigneeName,
          inviteeEmail: inviteeEmail,
          token,
          inviteUrl: `${baseClean}/user-dashboard.html?token=${encodeURIComponent(token)}&api=${encodeURIComponent(baseClean)}`,
          expiresAt: expires.toISOString(),
        });
      }
      await client.query('COMMIT');
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    }

    res.json({
      ok: true,
      uploadId: String(uploadId),
      expiresAt: expires.toISOString(),
      invites,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

/** PATCH body: { uploadId, token (uuid string), inviteeEmail } — empty inviteeEmail clears stored email. */
app.patch('/api/invites/email', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({ error: 'DATABASE_URL is not set' });
  }
  const body = req.body || {};
  const uploadId = body.uploadId;
  const token = body.token != null ? String(body.token).trim() : '';
  if (!uploadId || !/^\d+$/.test(String(uploadId))) {
    return res.status(400).json({ error: 'uploadId required (numeric)' });
  }
  if (!token) {
    return res.status(400).json({ error: 'token required' });
  }
  const rawEm = body.inviteeEmail != null ? String(body.inviteeEmail).trim() : '';
  const inviteeEmail = rawEm === '' ? null : rawEm;

  const client = await p.connect();
  try {
    const q = await client.query(
      `UPDATE invites SET invitee_email = $1
       WHERE token = $2::uuid AND upload_id = $3
       RETURNING invitee_email`,
      [inviteeEmail, token, uploadId]
    );
    if (!q.rows.length) {
      return res.status(404).json({ error: 'Invite not found for this upload' });
    }
    const row = q.rows[0];
    res.json({
      ok: true,
      inviteeEmail: row.invitee_email != null ? String(row.invitee_email) : '',
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

function getSmtpTransport() {
  if (!process.env.SMTP_HOST) return null;
  const port = parseInt(process.env.SMTP_PORT, 10) || 587;
  const secure = String(process.env.SMTP_SECURE || '').toLowerCase() === 'true';
  const user = process.env.SMTP_USER;
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port,
    secure,
    auth: user ? { user, pass: process.env.SMTP_PASS || '' } : undefined,
  });
}

/** MailerSend API token (https://www.mailersend.com/) — same value as in the dashboard. */
function getMailerSendApiToken() {
  const t = (
    process.env.MAILERSEND_API_TOKEN ||
    process.env.MAILERSEND_API_KEY ||
    ''
  ).trim();
  return t || null;
}

/**
 * Verified sender for MailerSend: MAILERSEND_FROM_EMAIL + MAILERSEND_FROM_NAME,
 * or parse "Name <email@verified-domain.com>" from MAILERSEND_FROM / SMTP_FROM / MAIL_FROM.
 */
function getMailerSendFrom() {
  const explicit = (process.env.MAILERSEND_FROM_EMAIL || '').trim();
  if (explicit && explicit.includes('@')) {
    const name = (process.env.MAILERSEND_FROM_NAME || '').trim();
    return {
      email: explicit,
      name: name || 'Jira Spillover Analyzer',
    };
  }
  const combined = (
    process.env.MAILERSEND_FROM ||
    process.env.SMTP_FROM ||
    process.env.MAIL_FROM ||
    ''
  ).trim();
  if (!combined) return null;
  const angle = combined.match(/<([^>]+)>/);
  const email = angle ? angle[1].trim() : /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(combined) ? combined : '';
  if (!email || !email.includes('@')) return null;
  let name = combined.replace(/<[^>]*>/, '').trim().replace(/^["']|["']$/g, '');
  if (!name) name = 'Jira Spillover Analyzer';
  return { email, name };
}

async function sendInviteViaMailerSend({ apiToken, fromEmail, fromName, to, toName, subject, text, html }) {
  const res = await fetch('https://api.mailersend.com/v1/email', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiToken}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
    body: JSON.stringify({
      from: { email: fromEmail, name: fromName || 'Jira Spillover Analyzer' },
      to: [{ email: to, name: toName || to }],
      subject,
      text,
      html,
    }),
  });
  if (res.ok) {
    return { ok: true };
  }
  let detail = '';
  try {
    const j = await res.json();
    detail = j.message || (j.errors && JSON.stringify(j.errors)) || JSON.stringify(j);
  } catch (_) {
    try {
      detail = await res.text();
    } catch (_) {
      detail = '';
    }
  }
  return { ok: false, error: detail || `HTTP ${res.status}` };
}

function escapeEmailHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function getInviteAppName() {
  const x = (process.env.INVITE_APP_NAME || '').trim();
  return x || 'Jira Spillover Analyzer';
}

let inviteEmailTemplateCache = null;
function getInviteEmailTemplateString() {
  if (inviteEmailTemplateCache) return inviteEmailTemplateCache;
  const fp = path.join(ROOT, 'templates', 'contributor-invite-email.html');
  inviteEmailTemplateCache = fs.readFileSync(fp, 'utf8');
  return inviteEmailTemplateCache;
}

/** Same layout as the Send invite “Styled” preview — fills templates/contributor-invite-email.html */
function renderInviteEmailHtml(assigneeName, inviteUrl) {
  const appName = getInviteAppName();
  const a = String(assigneeName || '').trim() || 'Contributor';
  const u = String(inviteUrl || '').trim();
  let html = getInviteEmailTemplateString();
  const esc = escapeEmailHtml;
  html = html.replace(/\{\{APP_NAME\}\}/g, esc(appName));
  html = html.replace(/\{\{ASSIGNEE_NAME\}\}/g, esc(a));
  html = html.replace(/\{\{INVITE_LINK\}\}/g, esc(u));
  return html;
}

function buildServerInvitePlainText(assigneeName, inviteUrl) {
  const appName = getInviteAppName();
  const n = String(assigneeName || '').trim();
  const link = String(inviteUrl || '').trim();
  const hi = n && n !== 'Contributor' ? `Hi ${n},` : 'Hi,';
  return [
    hi,
    '',
    `You're invited to use the ${appName} contributor workspace to enter spillover and bug information for the issues assigned to you in the shared export.`,
    '',
    'YOUR LINK (private — do not share)',
    '-----------------------------------',
    link,
    '',
    'What to do',
    '----------',
    '1. Open the link in your browser (Chrome, Edge, or Firefox recommended).',
    '2. Complete the spillover and bug fields shown for your tickets.',
    '3. Your progress is saved in the team database when you use the app online.',
    '',
    "If the link doesn't open, paste it into your browser's address bar.",
    '',
    'Thanks,',
  ].join('\n');
}

/** POST body: { uploadId, tokens: string[] } — sends each invite link to the email stored on that invite (never trusts client recipient). */
app.post('/api/invites/send-by-tokens', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({ error: 'DATABASE_URL is not set' });
  }
  const msToken = getMailerSendApiToken();
  const msFrom = getMailerSendFrom();
  if (msToken && !msFrom) {
    return res.status(503).json({
      error:
        'MAILERSEND_API_TOKEN is set but sender is missing. Set MAILERSEND_FROM_EMAIL (verified domain) or MAILERSEND_FROM / SMTP_FROM as "Name <email@domain.com>".',
    });
  }
  const useMailerSend = Boolean(msToken && msFrom);
  const transport = getSmtpTransport();
  if (!useMailerSend && !transport) {
    return res.status(503).json({
      error:
        'Email not configured. Set MAILERSEND_API_TOKEN and MAILERSEND_FROM_EMAIL (MailerSend), or SMTP_HOST (and typically SMTP_USER, SMTP_PASS, SMTP_FROM) — see .env.example.',
    });
  }
  const body = req.body || {};
  const uploadId = body.uploadId;
  const rawTokens = body.tokens;
  if (!uploadId || !/^\d+$/.test(String(uploadId))) {
    return res.status(400).json({ error: 'uploadId required (numeric)' });
  }
  if (!Array.isArray(rawTokens) || !rawTokens.length) {
    return res.status(400).json({ error: 'tokens must be a non-empty array' });
  }
  const from =
    process.env.SMTP_FROM || process.env.MAIL_FROM || process.env.SMTP_USER || 'noreply@localhost';
  const base = process.env.PUBLIC_APP_URL || `${req.protocol}://${req.get('host')}`;
  const baseClean = base.replace(/\/$/, '');

  const client = await p.connect();
  const results = [];
  try {
    for (const raw of rawTokens) {
      const token = String(raw || '').trim();
      if (!token) {
        results.push({ token: '', ok: false, error: 'empty token' });
        continue;
      }
      const q = await client.query(
        `SELECT invitee_email, assignee_scope
         FROM invites
         WHERE token = $1::uuid AND upload_id = $2`,
        [token, uploadId]
      );
      if (!q.rows.length) {
        results.push({ token, ok: false, error: 'Invite not found for this upload' });
        continue;
      }
      const row = q.rows[0];
      const to = row.invitee_email != null ? String(row.invitee_email).trim() : '';
      if (!to) {
        results.push({ token, ok: false, error: 'No email stored for this invite' });
        continue;
      }
      const name =
        row.assignee_scope != null && String(row.assignee_scope).trim() !== ''
          ? String(row.assignee_scope).trim()
          : 'Contributor';
      const inviteUrl = `${baseClean}/user-dashboard.html?token=${encodeURIComponent(token)}&api=${encodeURIComponent(baseClean)}`;
      const appName = getInviteAppName();
      const subject = `Your contributor link — ${name} — ${appName}`;
      const text = buildServerInvitePlainText(name, inviteUrl);
      const html = renderInviteEmailHtml(name, inviteUrl);
      try {
        if (useMailerSend) {
          const msResult = await sendInviteViaMailerSend({
            apiToken: msToken,
            fromEmail: msFrom.email,
            fromName: msFrom.name,
            to,
            toName: name,
            subject,
            text,
            html,
          });
          if (!msResult.ok) {
            throw new Error(msResult.error || 'MailerSend rejected request');
          }
        } else {
          await transport.sendMail({ from, to, subject, text, html });
        }
        results.push({ token, ok: true, to });
      } catch (sendErr) {
        console.error(sendErr);
        results.push({
          token,
          ok: false,
          error:
            process.env.NODE_ENV === 'development' ? String(sendErr.message || sendErr) : 'Send failed',
        });
      }
    }
    const sent = results.filter((r) => r.ok).length;
    res.json({ ok: true, sent, results });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Server error',
    });
  } finally {
    client.release();
  }
});

/**
 * Merge duplicate issue_field_edits rows that share the same normalized issue key
 * (e.g. legacy casing variants), taking non-empty values from each.
 */
function mergeIssueEditRowsForInvite(rows) {
  if (!rows || !rows.length) return [];
  const pickField = (a, b, key) => {
    const va = String(a[key] ?? '').trim();
    const vb = String(b[key] ?? '').trim();
    return va || vb ? va || vb : '';
  };
  const byNorm = new Map();
  for (const row of rows) {
    const n = normalizeIssueKeyDedupe(row.issue_key);
    const prev = byNorm.get(n);
    if (!prev) {
      byNorm.set(n, { ...row });
      continue;
    }
    byNorm.set(n, {
      issue_key: n,
      spillover_reason: pickField(prev, row, 'spillover_reason'),
      spillover_category: pickField(prev, row, 'spillover_category'),
      prod: pickField(prev, row, 'prod'),
      rca: pickField(prev, row, 'rca'),
      updated_at:
        new Date(prev.updated_at || 0) > new Date(row.updated_at || 0)
          ? prev.updated_at
          : row.updated_at,
    });
  }
  return Array.from(byNorm.values());
}

app.get('/api/invite/:token/session', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const raw = String(req.params.token || '').trim();
  if (!/^[0-9a-f-]{36}$/i.test(raw)) {
    return res.status(400).json({ error: 'Invalid token' });
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const got = await fetchInviteOr404(client, raw);
    if (!got) {
      return res.status(404).json({ error: 'Invite not found' });
    }
    if (got.expired) {
      return res.status(410).json({ error: 'Invite expired' });
    }

    const inv = got.invite;
    const { issues, headers } = await buildIssuesForInvite(client, inv);

    const edits = await client.query(
      `SELECT issue_key, spillover_reason, spillover_category, prod, rca, updated_at
       FROM issue_field_edits WHERE upload_id = $1`,
      [inv.upload_id]
    );

    res.json({
      ok: true,
      uploadId: String(inv.upload_id),
      originalName: inv.original_name,
      inviteeEmail: inv.invitee_email,
      assigneeScope: inv.assignee_scope || null,
      filterByAssignee: inv.filter_by_assignee,
      note: inv.note,
      expiresAt: inv.expires_at,
      sprintThreshold: inv.sprint_threshold || null,
      headers,
      issues,
      edits: mergeIssueEditRowsForInvite(edits.rows),
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.put('/api/invite/:token/issue', async (req, res) => {
  const p = getPool();
  if (!p) {
    return res.status(503).json({
      error: 'DATABASE_URL is not set. Add your database connection URI to .env',
    });
  }
  const raw = String(req.params.token || '').trim();
  const body = req.body || {};
  const issueKey = String(body.issueKey || '').trim();
  if (!issueKey) {
    return res.status(400).json({ error: 'issueKey required' });
  }

  const client = await p.connect();
  try {
    await ensureSchema(client);
    const got = await fetchInviteOr404(client, raw);
    if (!got) {
      return res.status(404).json({ error: 'Invite not found' });
    }
    if (got.expired) {
      return res.status(410).json({ error: 'Invite expired' });
    }

    const uploadId = got.invite.upload_id;
    const ikNorm = normalizeIssueKeyDedupe(issueKey);
    let sr = String(body.spilloverReason ?? '');
    let sc = String(body.spilloverCategory ?? '');
    let prod = String(body.prod ?? '');
    let rca = String(body.rca ?? '');
    const kind = String(body.kind || '').trim();

    const existingRes = await client.query(
      `SELECT issue_key, spillover_reason, spillover_category, prod, rca
       FROM issue_field_edits WHERE upload_id = $1`,
      [uploadId]
    );
    let ex = null;
    for (const row of existingRes.rows) {
      if (normalizeIssueKeyDedupe(row.issue_key) === ikNorm) {
        ex = row;
        break;
      }
    }
    if (ex) {
      if (kind === 'spill') {
        if (!String(prod ?? '').trim()) prod = String(ex.prod ?? '');
        if (!String(rca ?? '').trim()) rca = String(ex.rca ?? '');
      } else if (kind === 'bug') {
        if (!String(sr ?? '').trim()) sr = String(ex.spillover_reason ?? '');
        if (!String(sc ?? '').trim()) sc = String(ex.spillover_category ?? '');
      }
    }

    if (ex) {
      await client.query(
        `UPDATE issue_field_edits SET
           spillover_reason = $1,
           spillover_category = $2,
           prod = $3,
           rca = $4,
           updated_at = NOW()
         WHERE upload_id = $5 AND issue_key = $6`,
        [sr, sc, prod, rca, uploadId, ex.issue_key]
      );
    } else {
      await client.query(
        `INSERT INTO issue_field_edits
          (upload_id, issue_key, spillover_reason, spillover_category, prod, rca, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, NOW())`,
        [uploadId, ikNorm, sr, sc, prod, rca]
      );
    }

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: process.env.NODE_ENV === 'development' ? String(e.message || e) : 'Database error',
    });
  } finally {
    client.release();
  }
});

app.get('/index.html', (req, res) => {
  const i = req.url.indexOf('?');
  res.redirect(301, '/' + (i >= 0 ? req.url.slice(i) : ''));
});

/** Root HTML (not named index.html) so Vercel does not serve it as static before rewrites to this app. */
app.get('/', (_req, res) => {
  res.sendFile(path.join(ROOT, 'analyzer.html'));
});

app.use(express.static(ROOT));

module.exports = app;

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`http://localhost:${PORT}`);
    if (!process.env.DATABASE_URL) {
      console.warn('Warning: DATABASE_URL not set — /api/upload will return 503.');
    }
  });
}
