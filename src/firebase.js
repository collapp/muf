import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";

// Firebase config for the "datacollection-95473" project.
// Note: this apiKey is not a secret — Firebase apps are protected by
// Authentication + security rules, not by hiding this config. It is
// safe to commit this file to a public repo.
const firebaseConfig = {
  apiKey: "AIzaSyBW94cO5T4NZY2ipIkjPOQ22lpvZTA7bJc",
  authDomain: "datacollection-95473.firebaseapp.com",
  projectId: "datacollection-95473",
  storageBucket: "datacollection-95473.firebasestorage.app",
  messagingSenderId: "161764908214",
  appId: "1:161764908214:web:cc2c0059df11a24d269a17",
};

export const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
