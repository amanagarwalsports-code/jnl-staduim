import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { db } from './firebase.js'
import { ref, set, get, child, onValue, off } from 'firebase/database'

// ── Centralized Firebase storage with real-time sync ──────────────
window.storage = {
  get: async (key) => {
    try {
      const snapshot = await get(child(ref(db), `jln/${key}`))
      if (!snapshot.exists()) throw new Error(`Key not found: ${key}`)
      return { key, value: snapshot.val(), shared: true }
    } catch (err) {
      // Fallback to localStorage
      const value = localStorage.getItem(key)
      if (value === null) throw new Error(`Key not found: ${key}`)
      return { key, value, shared: false }
    }
  },
  set: async (key, value) => {
    try {
      await set(ref(db, `jln/${key}`), value)
      localStorage.setItem(key, value) // keep local copy as backup
      return { key, value, shared: true }
    } catch (err) {
      console.warn('Firebase write failed, saving locally:', err)
      localStorage.setItem(key, value)
      return { key, value, shared: false }
    }
  },
  delete: async (key) => {
    try {
      await set(ref(db, `jln/${key}`), null)
      localStorage.removeItem(key)
      return { key, deleted: true, shared: true }
    } catch (err) {
      localStorage.removeItem(key)
      return { key, deleted: true, shared: false }
    }
  },
  list: async (prefix = '') => {
    try {
      const snapshot = await get(ref(db, 'jln'))
      const keys = snapshot.exists()
        ? Object.keys(snapshot.val()).filter(k => k.startsWith(prefix))
        : []
      return { keys, prefix, shared: true }
    } catch (err) {
      const keys = Object.keys(localStorage).filter(k => k.startsWith(prefix))
      return { keys, prefix, shared: false }
    }
  },
  // Real-time listener — calls callback whenever data changes on any device
  listen: (key, callback) => {
    const dbRef = ref(db, `jln/${key}`)
    onValue(dbRef, (snapshot) => {
      if (snapshot.exists()) {
        try {
          callback(JSON.parse(snapshot.val()))
        } catch {
          callback(snapshot.val())
        }
      }
    }, (err) => {
      console.warn('Firebase listener error:', err)
    })
    // Return unsubscribe function
    return () => off(dbRef)
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
)
