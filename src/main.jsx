import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

// ── Centralized Database via Firebase Realtime Database ──────────────
// All devices share the SAME database — data is synced in real time.
// Setup: create a free Firebase project at https://console.firebase.google.com
// Then replace the config below with your own project's config.

const FIREBASE_CONFIG = {
  apiKey: "REPLACE_WITH_YOUR_API_KEY",
  authDomain: "REPLACE_WITH_YOUR_PROJECT.firebaseapp.com",
  databaseURL: "https://REPLACE_WITH_YOUR_PROJECT-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "REPLACE_WITH_YOUR_PROJECT",
  storageBucket: "REPLACE_WITH_YOUR_PROJECT.appspot.com",
  messagingSenderId: "REPLACE_WITH_YOUR_SENDER_ID",
  appId: "REPLACE_WITH_YOUR_APP_ID"
};

const FIREBASE_READY = !FIREBASE_CONFIG.apiKey.startsWith("REPLACE");

if (FIREBASE_READY) {
  // Load Firebase dynamically
  Promise.all([
    import('https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js'),
    import('https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js')
  ]).then(([{ initializeApp }, { getDatabase, ref, set, get, child }]) => {
    const app = initializeApp(FIREBASE_CONFIG);
    const db = getDatabase(app);

    window.storage = {
      get: async (key) => {
        const snapshot = await get(child(ref(db), `jln/${key}`));
        if (!snapshot.exists()) throw new Error(`Key not found: ${key}`);
        return { key, value: snapshot.val(), shared: true };
      },
      set: async (key, value, shared = true) => {
        await set(ref(db, `jln/${key}`), value);
        return { key, value, shared };
      },
      delete: async (key) => {
        await set(ref(db, `jln/${key}`), null);
        return { key, deleted: true, shared: true };
      },
      list: async (prefix = '') => {
        const snapshot = await get(ref(db, 'jln'));
        const keys = snapshot.exists()
          ? Object.keys(snapshot.val()).filter(k => k.startsWith(prefix))
          : [];
        return { keys, prefix, shared: true };
      }
    };

    mountApp();
  }).catch(err => {
    console.error('Firebase failed, falling back to localStorage:', err);
    setupLocalStorage();
    mountApp();
  });
} else {
  // Firebase not configured yet — use localStorage as fallback
  setupLocalStorage();
  mountApp();
}

function setupLocalStorage() {
  window.storage = {
    get: async (key) => {
      const value = localStorage.getItem(key);
      if (value === null) throw new Error(`Key not found: ${key}`);
      return { key, value, shared: false };
    },
    set: async (key, value, shared = false) => {
      localStorage.setItem(key, value);
      return { key, value, shared };
    },
    delete: async (key) => {
      localStorage.removeItem(key);
      return { key, deleted: true, shared: false };
    },
    list: async (prefix = '') => {
      const keys = Object.keys(localStorage).filter(k => k.startsWith(prefix));
      return { keys, prefix, shared: false };
    }
  };
}

function mountApp() {
  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}
