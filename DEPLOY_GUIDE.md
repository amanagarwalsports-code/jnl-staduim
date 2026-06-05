# JLN Stadium — Deployment Guide
## From Files → Live on Android & iPhone in ~20 Minutes

---

## WHAT YOU'LL NEED
- A computer (Windows, Mac, or Linux)
- Internet connection
- A free GitHub account → https://github.com
- A free Vercel account → https://vercel.com
- Android (Chrome) or iPhone (Safari) to install

---

## STEP 1 — Install Node.js (5 min)

1. Go to https://nodejs.org
2. Download the **LTS** version (left button)
3. Install it — click Next through all the defaults
4. Open Terminal (Mac/Linux) or Command Prompt (Windows)
5. Confirm it worked:
   ```
   node --version
   npm --version
   ```
   Both should print a version number (e.g. v20.x.x)

---

## STEP 2 — Set Up the Project (3 min)

1. Create a folder on your computer, e.g. `jln-stadium`
2. Copy ALL the files from this download into that folder. Structure:
   ```
   jln-stadium/
   ├── index.html
   ├── package.json
   ├── vite.config.js
   ├── vercel.json
   ├── .gitignore
   ├── public/
   │   └── manifest.json
   └── src/
       ├── main.jsx
       └── App.jsx
   ```

3. Open Terminal/Command Prompt, navigate to the folder:
   ```
   cd path/to/jln-stadium
   ```

4. Install dependencies:
   ```
   npm install
   ```
   (This downloads ~150MB of packages into node_modules — takes 1-2 min)

5. Test it locally:
   ```
   npm run dev
   ```
   Open http://localhost:5173 in your browser — you should see the JLN Stadium app!
   Press Ctrl+C to stop.

---

## STEP 3 — Add App Icons (Optional but Recommended)

The app needs 3 icon files in the /public folder:
- `icon-192.png` — 192×192 pixels
- `icon-512.png` — 512×512 pixels
- `apple-touch-icon.png` — 180×180 pixels

You can create these at https://favicon.io or use any image editor.
Use a stadium/car/shield image. Dark navy background (#080f1e) with gold icon matches the app theme.

If you skip this step, the app still works — just without a custom home screen icon.

---

## STEP 4 — Build the Production App (1 min)

```
npm run build
```

This creates a `/dist` folder — the optimized, deployable version of your app.

---

## STEP 5 — Push to GitHub (5 min)

1. Go to https://github.com → Sign up (free) or log in
2. Click the **+** icon → **New repository**
3. Name it `jln-stadium`, set to **Private**, click **Create repository**
4. Back in your Terminal:

   ```
   git init
   git add .
   git commit -m "JLN Stadium v6 initial deploy"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/jln-stadium.git
   git push -u origin main
   ```
   (Replace YOUR_USERNAME with your GitHub username)

---

## STEP 6 — Deploy on Vercel (2 min)

1. Go to https://vercel.com → Sign up with GitHub (free)
2. Click **Add New Project**
3. Find and select your `jln-stadium` repository → Click **Import**
4. Vercel auto-detects Vite. Leave all settings as-is.
5. Click **Deploy**

   ⏳ Wait ~60 seconds...

6. You'll get a live URL like: **https://jln-stadium.vercel.app**

   ✅ Your app is now live on the internet with HTTPS!

---

## STEP 7 — Install on Android (1 min)

1. Open **Chrome** on your Android phone
2. Go to your Vercel URL (e.g. https://jln-stadium.vercel.app)
3. Tap the **⋮ menu** (top right)
4. Tap **"Add to Home screen"** or **"Install app"**
5. Tap **Add**

The app icon appears on your home screen. Tap it — it opens full screen, no browser bar, exactly like a native app.

---

## STEP 8 — Install on iPhone (1 min)

1. Open **Safari** on your iPhone (must be Safari, not Chrome)
2. Go to your Vercel URL
3. Tap the **Share button** (square with arrow pointing up)
4. Scroll down and tap **"Add to Home Screen"**
5. Tap **Add**

The app icon appears on your home screen. Tap it — it opens full screen in standalone mode.

---

## FUTURE UPDATES

When you make changes to the app:

```
npm run build
git add .
git commit -m "Update: description of changes"
git push
```

Vercel automatically re-deploys in ~30 seconds. All phones get the update automatically next time they open the app.

---

## SHARING WITH STAFF

Just send them your Vercel URL. They can install it on their own phones using Steps 7 or 8.
No App Store, no download, no installation file to manage.

---

## DATA STORAGE

The app stores all data (master vehicle list, entry logs, users, etc.) in the phone's browser localStorage.

⚠️ IMPORTANT: Data is per-device. Each phone has its own database.
If you need shared data across multiple security staff devices, that requires a backend (Firebase, Supabase, etc.) — ask if you'd like a guide for that.

---

## DEFAULT CREDENTIALS

- Admin login: admin / admin123
- Staff login:  staff / staff123

Change these immediately after first login via Admin → Change Passwords.

---

## TROUBLESHOOTING

**"npm: command not found"** → Node.js not installed. Repeat Step 1.

**"Module not found: xlsx"** → Run `npm install` again.

**App not loading on phone** → Make sure you're using Chrome (Android) or Safari (iPhone). Other browsers may not support PWA installation.

**Icon not showing** → Add the PNG icons to /public and redeploy (Step 4 → Step 6).

**Data disappeared** → localStorage was cleared (phone storage cleared, or private/incognito mode). This is expected behaviour — data lives in the browser.
