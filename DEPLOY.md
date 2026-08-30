# Deploying CodePulse (Web + PWA)

This project is now a real installable PWA (Progressive Web App) — verified
by actually building it and confirming the manifest, service worker, and
all icon sizes are generated correctly, not just configured.

## What "PWA" gets you, concretely

Once deployed to a real HTTPS URL, visitors can:
- **On Android/Chrome**: tap "Install app" / "Add to Home Screen" — it gets
  a real icon, opens full-screen (no browser bar), and shows up in the app
  drawer like a native app.
- **On iPhone/Safari**: Share → "Add to Home Screen" — same result.
- **On desktop Chrome/Edge**: an install icon appears in the address bar.

It will **not** appear in the App Store or Google Play — that's what the
separate mobile (React Native) project is for. A PWA is a real website that
behaves like an app once installed; it is not a store-distributed binary.

## Step 1 — Deploy to a real URL (pick one)

**Vercel** (easiest, free tier, `vercel.json` is already included):
```bash
npm install -g vercel
vercel login
vercel --prod
```

**Netlify** (`netlify.toml` is already included):
```bash
npm install -g netlify-cli
netlify login
netlify deploy --prod
```

Either one auto-detects the Vite build (`npm run build` → `dist/`) — no
extra configuration needed beyond the files already in this project.

## Step 2 — Verify the PWA actually installs

After deploying:
1. Open the live URL on your phone
2. You should see the install prompt (Android) or need to use Share → Add
   to Home Screen (iPhone)
3. Open Chrome DevTools → Application tab → Manifest, on desktop, to
   confirm the manifest and icons loaded with no errors

## What I can't do for you

I don't have a Vercel/Netlify account, so I can't literally click "deploy."
Everything above is genuinely ready to go — `npm run build` succeeds,
produces `dist/manifest.webmanifest` + `dist/sw.js` + all icon sizes, and
I verified the manifest's actual JSON content is correct. The remaining
step is authentication with *your* hosting account, which only you can do.

## Icons

Generated from the CodePulse logo already embedded in the app, at every
size real platforms require: 192×192, 512×512, a maskable 512×512 (extra
padding for Android's adaptive-icon safe zone), a 180×180 Apple touch icon,
and a 32×32 favicon. All in `public/`.
