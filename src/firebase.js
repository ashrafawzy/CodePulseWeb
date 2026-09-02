// Firebase project config — reads from environment variables so real
// credentials never get committed to GitHub. Copy .env.example to .env.local
// and fill in your own Firebase project's values (Firebase Console →
// Project Settings → General → "Your apps" → SDK setup and configuration).
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

if (!firebaseConfig.apiKey) {
  // Fails loudly instead of Firestore silently hanging on every read/write —
  // the single most common "why isn't anything loading" issue with this setup.
  console.error(
    "Firebase config is missing. Copy .env.example to .env.local and fill in your Firebase project's values, then restart the dev server."
  );
}

export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
