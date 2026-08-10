import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyA9pvMee-kkVFIUXE4islcwlkdkNwFcEq4",
  authDomain: "stok-takip-f84a2.firebaseapp.com",
  projectId: "stok-takip-f84a2",
  storageBucket: "stok-takip-f84a2.firebasestorage.app",
  messagingSenderId: "572028161303",
  appId: "1:572028161303:web:dd9a75fac084a1ee65bb52",
  measurementId: "G-N4HS6D87BC",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
