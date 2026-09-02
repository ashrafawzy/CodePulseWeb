import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["favicon.png", "apple-touch-icon.png"],
      manifest: {
        name: "CodePulse Enterprise Console",
        short_name: "CodePulse",
        description: "CodePulse ERP — CRM, Sales, Purchasing, Inventory, Manufacturing, Accounting, HR, and more.",
        theme_color: "#1B2421",
        background_color: "#1B2421",
        display: "standalone",
        orientation: "portrait-primary",
        start_url: "/",
        scope: "/",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "maskable-icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
      workbox: {
        // App shell + assets cached for offline load; this app's own data
        // is in-memory only (or fetched from your API in the -db variant),
        // so this caches the UI, not live business data.
        globPatterns: ["**/*.{js,css,html,png,svg,woff2}"],
        // Without these, a new service worker installs after each deploy
        // but stays in the browser's "waiting" state until every open tab
        // of the app is fully closed — so a reload alone can still show
        // stale content days after a real update shipped. This is what
        // caused the Manufacturing menu to look missing even though it
        // was already in the deployed code. skipWaiting + clientsClaim
        // make a new deployment take over immediately instead.
        skipWaiting: true,
        clientsClaim: true,
      },
    }),
  ],
});
