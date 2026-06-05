import { initializeApp } from 'firebase/app'
import { getDatabase } from 'firebase/database'

const firebaseConfig = {
  apiKey: "AIzaSyBALO1EPurLK6OW0aqKzpr-oh9mmuN8emY",
  authDomain: "jln-stadium.firebaseapp.com",
  databaseURL: "https://jln-stadium-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "jln-stadium",
  storageBucket: "jln-stadium.firebasestorage.app",
  messagingSenderId: "153659490043",
  appId: "1:153659490043:web:e0c05c16b38866360c3a5e"
}

const app = initializeApp(firebaseConfig)
export const db = getDatabase(app)
