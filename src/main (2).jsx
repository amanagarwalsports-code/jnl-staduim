import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { db } from './firebase.js'
import { ref, set, get, child } from 'firebase/database'

window.storage = {
  get: async (key) => {
    try {
      const snapshot = await get(child(ref(db), `jln/${key}`))
      if (!snapshot.exists()) throw new Error(`Key not found: ${key}`)
      return { key, value: snapshot.val(), shared: true }
    } catch (err) {
      const value = localStorage.getItem(key)
      if (value === null) throw new Error(`Key not found: ${key}`)
      return { key, value, shared: false }
    }
  },
  set: async (key, value) => {
    try {
      await set(ref(db, `jln/${key}`), value)
      localStorage.setItem(key, value)
      return { key, value, shared: true }
    } catch (err) {
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
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
)
