(function () {
    window.firebaseReady = false;
    window.firebaseDb = null;
    window.firebaseInitError = null;

    if (typeof firebase === 'undefined') {
        return;
    }

    var cfg = typeof window.FIREBASE_CONFIG === 'object' && window.FIREBASE_CONFIG ? window.FIREBASE_CONFIG : null;
    if (!cfg || !cfg.apiKey || !cfg.projectId) {
        window.firebaseSkippedReason =
            'firebase-config.js: set apiKey and projectId from Firebase Console (Project settings → Web app).';
        return;
    }

    var key = String(cfg.apiKey);
    var pid = String(cfg.projectId);
    if (key.indexOf('YOUR_') === 0 || pid.indexOf('YOUR_') === 0) {
        window.firebaseSkippedReason =
            'Replace YOUR_* placeholders in firebase-config.js with your Firebase project values.';
        return;
    }

    try {
        if (!firebase.apps || !firebase.apps.length) {
            firebase.initializeApp(cfg);
        }
        window.firebaseReady = true;
        if (typeof firebase.firestore === 'function') {
            window.firebaseDb = firebase.firestore();
        }
    } catch (err) {
        window.firebaseInitError = (err && err.message) || String(err);
    }
})();
