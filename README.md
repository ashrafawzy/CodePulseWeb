# CodePulse — Standalone Project

The full CodePulse ERP web app, set up as a normal Vite + React project you
can open in VSCode and run on your own machine — instead of living only
inside Claude's artifact preview.

## Why this matters

A few things that were broken *specifically because* this ran inside
Claude's sandboxed preview iframe (not because of bugs in the code) should
now work normally once you run this yourself: printing, and anything else
that depends on standard browser permissions the preview sandbox
restricts.

## Setup

```bash
npm install
npm run dev
```

Then open the URL it prints (typically `http://localhost:5173`).

I verified this actually builds correctly before handing it to you —
`npm run build` succeeds, and I confirmed by inspecting the output that
Tailwind genuinely generated the CSS for the app's classes (including the
arbitrary-value ones like `text-[11px]`) and that the bundled JavaScript
contains the real app content — not just that it compiled without errors.

```bash
npm run build     # production build -> dist/
npm run preview   # serve that build locally to sanity-check it
```

## Project structure

```
index.html          — entry HTML
src/main.jsx         — mounts <App /> into #root
src/App.jsx           — the entire CodePulse application (single file, as built)
src/index.css         — Tailwind directives
tailwind.config.js    — scans index.html + src/**/*.{js,jsx} for classes used
postcss.config.js     — required for Tailwind to process
vite.config.js         — minimal Vite + React plugin config
```

## What you're getting vs. what's still a prototype

This packaging change makes it a real, runnable local project — it does
**not** change what the app itself does. Same honest status as before:

- **Still in-memory** — all data resets on page refresh. Connecting this to
  the `codepulse-api` backend (built separately, with a real PostgreSQL
  database) is the natural next step; ask if you want me to wire that up
  the same way I did for the mobile app.
- **Login is cosmetic** — picks from 5 demo accounts, no real password check.
- Three languages (English/Spanish/Arabic with RTL), light/dark theme,
  multi-currency, and full Export (Excel/CSV/HTML with logo) all work as
  already built.

## Deploying it for real

Once you're happy with it locally, `npm run build` produces a static
`dist/` folder — deployable as-is to Vercel, Netlify, Cloudflare Pages, or
any static host. No server-side rendering or special hosting needed for
the frontend itself.
