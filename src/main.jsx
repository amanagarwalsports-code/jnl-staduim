import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { initializeApp } from 'firebase/app'
import { getDatabase, ref, set, get, child } from 'firebase/database'

const firebaseConfig = {
  apiKey: "AIzaSyBALO1EPurLK6OW0aqKzpr-oh9mmuN8emY",
  authDomain: "jln-stadium.firebaseapp.com",
  databaseURL: "https://jln-stadium-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "jln-stadium",
  storageBucket: "jln-stadium.firebasestorage.app",
  messagingSenderId: "153659490043",
  appId: "1:153659490043:web:e0c05c16b38866360c3a5e"
};

function setupLocalStorage() {
  window.storage = {
    get: async (key) => {
      const value = localStorage.getItem(key);
      if (value === null) throw new Error(`Key not found: ${key}`);
      return { key, value, shared: false };
    },
    set: async (key, value) => {
      localStorage.setItem(key, value);
      return { key, value, shared: false };
    },
    delete: async (key) => {
      localStorage.removeItem(key);
      return { key, deleted: true };
    },
    list: async (prefix = '') => {
      const keys = Object.keys(localStorage).filter(k => k.startsWith(prefix));
      return { keys, prefix };
    }
  };
}

function mountApp() {
  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode><App /></React.StrictMode>
  );
}

try {
  const app = initializeApp(firebaseConfig);
  const db = getDatabase(app);

  window.storage = {
    get: async (key) => {
      const snapshot = await get(child(ref(db), `jln/${key}`));
      if (!snapshot.exists()) throw new Error(`Key not found: ${key}`);
      return { key, value: snapshot.val(), shared: true };
    },
    set: async (key, value) => {
      await set(ref(db, `jln/${key}`), value);
      return { key, value, shared: true };
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
} catch (err) {
  console.error('Firebase failed, using localStorage:', err);
  setupLocalStorage();
  mountApp();
}
