# Deploying CodePulse (Firebase + GitHub + Vercel)

The data layer is now Firebase Firestore — every collection (leads,
products, sales orders, everything) is real-time and shared, not
per-browser-session state anymore. This still builds as a normal static
site, so Vercel/Netlify/Firebase Hosting all work the same way as before;
Firebase is specifically for the *data*, not for hosting the app (though
Firebase Hosting is included as an option below too, since it's one less
account to manage if you're already using Firebase for data).

## Step 1 — Create a Firebase project (you do this — I can't create
accounts on your behalf)

1. Go to https://console.firebase.google.com → **Add project**
2. Once created, go to **Build → Firestore Database → Create database**
   — start in **test mode** for now (see the security note below)
3. Go to **Project Settings** (gear icon) → **General** → scroll to
   **"Your apps"** → click the **Web** icon (`</>`) → register the app
4. Copy the `firebaseConfig` values it gives you

## Step 2 — Configure this project with those values

```bash
cp .env.example .env.local
```
Paste the 6 values from Step 1 into `.env.local`. This file is gitignored
— it never gets committed.

```bash
npm install
npm run dev
```
Open it locally and confirm data loads. The first time it connects, it
seeds Firestore from this project's existing demo data (products,
customers, etc.) — after that, Firestore is the source of truth and the
seed data is never touched again.

## ⚠️ Before you put any real data in — read this

The current login is a **demo picker** (choose one of 5 seeded accounts) —
it's a piece of local UI state, not real authentication. **Firestore
Security Rules cannot enforce your app's role permissions (Admin-only
edits, etc.) without real Firebase Authentication wired in.** Right now,
anyone with your deployed URL could open browser dev tools and read/write
your Firestore data directly, bypassing the app's UI entirely.

`firestore.rules` in this project is deliberately left wide open
(`allow read, write: if true`) with a big comment explaining exactly this,
so it's obvious rather than hidden. This is fine for testing. **It is not
fine for real business data or real users.** Making it real requires
adding actual Firebase Authentication + a Cloud Function to assign roles —
genuine additional work I didn't build in this pass. Ask if you want that
built before you go live with real data.

## Step 3 — Push to GitHub

This project is already a git repository with an initial commit. You need
your own GitHub repo to push it to:

```bash
# Create a new repo at https://github.com/new first, then:
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git branch -M main
git push -u origin main
```

## Step 4 — Deploy (pick one)

### Option A: Vercel, connected to GitHub (recommended — auto-deploys on every push)
1. https://vercel.com → **Add New Project** → import your GitHub repo
2. Vercel auto-detects the Vite build — no config needed (`vercel.json` is
   already in this repo)
3. **Add your 6 Firebase env vars** in Vercel's project settings →
   Environment Variables (same names as `.env.example`) — this step is
   easy to miss and the #1 reason a Vercel deploy shows a blank/broken app
4. Deploy — from now on, every `git push` auto-deploys

### Option B: Firebase Hosting (keeps everything in one Firebase project)
```bash
npm install -g firebase-tools
firebase login
firebase init hosting   # point it at this project, dist as the public dir
npm run build
firebase deploy
```
`firebase.json` is already included and configured for this.

### Option C: Netlify
Same as Vercel — connect the GitHub repo, add the same 6 env vars in
Netlify's dashboard (`netlify.toml` is already included).

## What I actually verified before handing this to you

- Ran a real production build (`npm run build`) after the Firestore
  migration — succeeds cleanly, 52 modules, PWA artifacts still generate
  correctly
- Diffed the exact set of values exposed to the rest of the app before and
  after the migration — all 145 match exactly, so none of the ~3,000 lines
  of page components needed to change; they consume data the same way
  regardless of where it comes from
- Initialized a real git repository with a real commit — this project is
  ready to push, not just described as "should work with git"

## What I could not verify

I have no Firebase project, no GitHub account, and no network access to
either platform from this sandbox. Everything above is genuinely built and
locally validated — the actual first connection to a real Firestore
database, the actual push to a real GitHub remote, and the actual Vercel
deploy are steps only you can take and verify.
