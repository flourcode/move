# Putting this online

You don't need a language you don't already have. GitHub Pages serves plain
files, and the "backend" is a small script that lives inside your own
spreadsheet. About ten minutes, once.

| File | Where it goes |
|---|---|
| `index.html` | GitHub repo |
| `manifest.json` | GitHub repo |
| `apple-touch-icon.png`, `icon-192.png`, `icon-512.png` | GitHub repo |
| `Code.gs` | Google Sheets (a copy in the repo is fine too) |

---

## 1. Put the script in your spreadsheet

1. Open your sheet → **Extensions ▸ Apps Script**.
2. Select everything in the editor and paste in the contents of `Code.gs`. Save.
3. Choose `testConnection` from the function dropdown and press **Run**.
   Google asks for permission — click **Advanced ▸ Go to (project name)** and
   allow it. That warning is normal; you're authorising your own script.
4. The log at the bottom should read `10 items · closing 2026-10-06`.

## 2. Publish the script

1. **Deploy ▸ New deployment**.
2. Click the gear next to "Select type" → **Web app**.
3. Execute as: **Me**. Who has access: **Anyone**.
   ("Anyone", not "Anyone with a Google account" — otherwise your wife hits a
   Google login wall.)
4. **Deploy**, then **copy the web app URL**. It ends in `/exec`.

## 3. Make the repo

1. On GitHub: **New repository**. Call it something like `move-in`. Public.
   Don't add a README.
2. On the empty repo page, click **uploading an existing file**.
3. Drag in `index.html`, `manifest.json` and the three `.png` files.
4. **Commit changes**.

## 4. Paste your URL into the app

1. In the repo, click `index.html`, then the **pencil** icon to edit.
2. Find **line 384**: `const SCRIPT_URL = '';`
   (Cmd+F / Ctrl+F works in the editor — search for `SCRIPT_URL`.)
3. Paste your URL between the quotes:
   `const SCRIPT_URL = 'https://script.google.com/macros/s/AKfyc.../exec';`
4. **Commit changes**.

## 5. Turn on Pages

1. Repo **Settings ▸ Pages**.
2. Source: **Deploy from a branch**. Branch: **main**, folder: **/ (root)**. **Save**.
3. Wait a minute, refresh, and the address appears at the top:
   `https://YOURNAME.github.io/move-in/`

Open it. Your ten items load straight away — no sign-in, nothing to type.

## 6. Send it to your wife

Text her the link. That's it.

On iPhone, tell her to open it in **Safari** and tap **Share ▸ Add to Home
Screen** — it gets the house icon and opens full-screen with no browser bars,
like a real app. Android: Chrome menu ▸ **Add to Home screen**.

---

## What "no passcode" actually means

The page is public, so anyone who stumbles across the URL can add, edit and
delete rows. They can't open your actual spreadsheet — the script runs as you,
so the sheet itself stays private, and the app only exposes the ten columns in
that one tab.

Realistically nobody will find a random `github.io` URL. But if it ever becomes
a nuisance, the lock is already built in: put any word on **line 19** of
`Code.gs`, then **Deploy ▸ Manage deployments ▸ pencil ▸ Version: New version ▸
Deploy**. The app starts asking for it on its own. Nothing to change in the
HTML.

## When something doesn't work

**"That URL didn't return sheet data."** The deployment's access is set to
"Only myself". Redeploy with **Anyone**.

**Your wife hits a Google sign-in page.** Access was set to "Anyone with a
Google account". Redeploy with plain **Anyone**.

**You edited `Code.gs` and nothing changed.** Google keeps serving the old
version until you publish a new one: **Deploy ▸ Manage deployments ▸ pencil ▸
Version: New version ▸ Deploy**. This catches everyone at least once.

**You changed `index.html` but the site looks the same.** GitHub caches for a
few minutes. Hard refresh: Cmd+Shift+R (Mac) or Ctrl+F5 (Windows).

**It loads but edits don't stick.** Check the tab is still named
`Move-In Purchases` and row 7 still says `Room / Space`. The script finds your
table by that header.

**You renamed a column.** Update the matching line in the `H` block near the
top of the script section in `index.html` (around line 390).
