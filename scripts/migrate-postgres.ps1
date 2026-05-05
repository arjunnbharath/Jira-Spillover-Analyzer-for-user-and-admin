<#
.SYNOPSIS
  Copy one PostgreSQL database to another (e.g. Aiven -> Neon / Vercel Postgres).

.DESCRIPTION
  Uses pg_dump (custom format) and pg_restore. Install PostgreSQL client tools first:
  https://www.postgresql.org/download/windows/  (add bin to PATH)

  The TARGET database should be EMPTY (new Neon/Vercel DB). pg_restore may fail if
  the target already has conflicting objects.

.EXAMPLE
  $src = "postgres://..."   # current Aiven URI (from Aiven console)
  $dst = "postgres://..."   # new Neon URI (from Vercel Storage / Neon dashboard)
  .\scripts\migrate-postgres.ps1 -SourceUrl $src -TargetUrl $dst

  Then set DATABASE_URL to $dst in Vercel (Environment Variables) and in local .env, redeploy, and delete the Aiven service when satisfied.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SourceUrl,

    [Parameter(Mandatory = $true)]
    [string] $TargetUrl
)

$ErrorActionPreference = 'Stop'

function Test-PostgresTool {
    param([string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $cmd) {
        throw "Missing '$Name' in PATH. Install PostgreSQL client tools and restart the terminal."
    }
}

Test-PostgresTool 'pg_dump'
Test-PostgresTool 'pg_restore'

$dumpFile = Join-Path $env:TEMP ("jira-spillover-migrate-{0}.dump" -f [Guid]::NewGuid().ToString('N'))

try {
    Write-Host "Dumping source database to $dumpFile ..."
    & pg_dump --format=custom --no-owner --file="$dumpFile" "$SourceUrl"
    if ($LASTEXITCODE -ne 0) { throw "pg_dump failed with exit $LASTEXITCODE" }

    Write-Host "Restoring into target database (empty DB recommended) ..."
    & pg_restore --dbname="$TargetUrl" --no-owner --no-acl --verbose "$dumpFile"
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne 1) {
        # pg_restore sometimes returns 1 for non-fatal warnings
        throw "pg_restore failed with exit $LASTEXITCODE"
    }

    Write-Host ""
    Write-Host "Done. Next steps:"
    Write-Host "  1. Set DATABASE_URL in Vercel + local .env to the NEW (target) connection string."
    Write-Host "  2. Redeploy on Vercel; run npm start locally and open /api/health (databaseReachable should be true)."
    Write-Host "  3. When verified, remove or power down the old Aiven service."
}
finally {
    if (Test-Path $dumpFile) {
        Remove-Item -LiteralPath $dumpFile -Force -ErrorAction SilentlyContinue
    }
}
