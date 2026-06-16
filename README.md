# Custom Video Request — MyGeotab Add-In

Recreates the Surfsight portal's *"Request custom video recording"* picker (event start time + duration) **inside MyGeotab**, for both **GoFocus (Smarter AI)** and **Surfsight** cameras.

## How it works — no proxy

A classic MyGeotab add-in runs in the **`my.geotab.com` origin** (its HTML/JS is injected into the MyGeotab page). That's the same origin the native Cameras view uses to call `media-services.geotab.com`, so the add-in calls Camera-Services **directly** — no CORS issue, no proxy. The MyGeotab session is already authenticated; we just read it with `api.getSession()` and build the four `X-MyGeotab-*` headers from it.

```
List cameras:   GET  /DeviceMappings
Request clip:   POST /Media
```

The camera dropdown is built from `/DeviceMappings`, joined to MyGeotab `Device` names for friendly labels. The add-in branches on `partnerId` automatically:

| Camera | `partnerId` | `partnerDeviceId` |
|--------|-------------|-------------------|
| GoFocus / GO Focus Plus | `smarterai` | camera serial |
| Surfsight | `surfsight` | camera IMEI |
| Sensata | `sensata` | recorderId |

## Two variants

| Variant | Folder | Where it appears |
|---------|--------|------------------|
| **Page add-in** | root (`customVideoRequest.*`) | Its own menu item. Self-contained / embeddable via `config.packed.json`. |
| **Trips History map add-in** | `trips-history/` | A launcher in the Trips History side panel that opens the request modal — like the native dialog. Hosted (GitHub Pages). |

### Trips History map add-in (`trips-history/`)

A Map add-in with `"page": "tripsHistory"` (`configuration.json`). It registers `geotab.addin.request = (elt, service) => …` and renders a "Request custom video recording" launcher in the Trips History side panel; clicking opens a modal styled like the native dialog. Hovering/clicking a camera-equipped vehicle on the map pre-selects its camera.

- **Hosting:** map add-ins load from a URL (no embedded/packed option). Push the repo to GitHub Pages — `configuration.json`'s `mapScript.url` points at `Austin-GPSFMS.github.io/custom-video-request-addin/trips-history/addin.html`. Update that to your repo path.
- **Same direct-to-media-services approach** — map add-ins run in the `my.geotab.com` context, so the `X-MyGeotab-*` headers from `service.api.getSession()` work without a proxy.
- **Note:** the vehicle→camera pre-select uses map `over`/`click` events (guarded). If a future MyGeotab version changes those event hooks, the dropdown still works manually; nothing else breaks.

## Files

```
custom-video-request-addin/
├── config.json              # Embedded-files manifest (paste into MyGeotab)
├── config.hosted.json       # Hosted-URL manifest variant
├── config.packed.json       # Self-contained: HTML/JS/CSS base64-embedded
├── customVideoRequest.html
├── customVideoRequest.js
├── styles.css
├── translations/en.json
└── README.md
```

## Deploy to GitHub Pages (recommended — edit once, push to update)

Hosting the files means you paste the config into MyGeotab **once**; after that, every change is just a `git push` — no re-uploading.

1. Create an **empty** repo on GitHub named **`custom-video-request-addin`** (no README/license, to avoid conflicts).
2. From this folder, run the included setup script (it removes the partial `.git`, inits cleanly, commits, sets the remote, and pushes):
   ```powershell
   .\init-repo.ps1
   ```
   Or do it manually:
   ```bash
   rm -rf .git              # remove the partial .git from the sandbox
   git init && git add -A && git commit -m "Custom Video Request add-in"
   git branch -M main
   git remote add origin https://github.com/Austin-GPSFMS/custom-video-request-addin.git
   git push -u origin main
   ```
3. GitHub → repo **Settings → Pages** → Source: **Deploy from a branch**, Branch: **main / (root)**, Save. It publishes at `https://Austin-GPSFMS.github.io/custom-video-request-addin/`.
4. In MyGeotab → **Administration → System… → Add-Ins → New Add-In**, paste:
   - **`config.hosted.json`** for the page add-in (own menu item), and/or
   - **`trips-history/configuration.json`** for the Trips History modal.

After that, editing any file + `git push` updates the live add-in on next reload. No more file uploads.

> `.nojekyll` is included so GitHub Pages serves the files as-is.

## Install (embedded alternative)

**Hosted:** host the three files over HTTPS, point `config.hosted.json`'s `url` at the HTML, then MyGeotab → **Administration → System… → Add-Ins → New Add-In** and paste that JSON.

**Self-contained:** paste `config.packed.json` directly — no hosting needed (HTML/JS/CSS travel inside the config as base64). Regenerate it after edits with:

```bash
# from this folder. NOTE: translations/en.json MUST be embedded — MyGeotab probes
# translations/<lang>.json on load, and a 404 there causes "Issue Loading This Page".
node -e '
  const fs=require("fs"); const c=JSON.parse(fs.readFileSync("config.json","utf8"));
  const mime={html:"text/html",js:"application/javascript",css:"text/css",json:"application/json"};
  const enc=f=>"data:"+mime[f.split(".").pop()]+";base64,"+fs.readFileSync(f).toString("base64");
  c.files={};["customVideoRequest.html","customVideoRequest.js","styles.css","translations/en.json"].forEach(f=>c.files[f]=enc(f));
  fs.writeFileSync("config.packed.json",JSON.stringify(c,null,2));
'
```

> **"Issue Loading This Page"?** MyGeotab fetches `translations/en.json` relative to the add-in root on load; if it 404s, the page fails. The packed config embeds it, so re-paste `config.packed.json`. If hosting, make sure the `translations/` folder is served too.

## Required roles

The logged-in user needs Camera-Services clearances — at minimum `ListAssets` (read camera list) and `ViewRecordedVideo` (request clips). A `403` means the session is valid but missing a role; a `401` means it expired (the add-in re-reads the session and retries once).

## Notes

- **Duration cap:** clips are limited to 120s client-side (`MAX_DURATION_SECONDS`) to avoid accidental large/expensive pulls. Adjust in `customVideoRequest.js`.
- **AddInData:** recent requests are kept in memory this session. If you want them (and the last-used duration) persisted across sessions, MyGeotab's `AddInData` storage entity is the right place — easy to add next.
- **Map variant:** this is a page add-in. It can be reworked as a **map add-in** so clicking a vehicle pre-fills the camera.
- **Service account / governance:** only needed for *unattended* automation or central logging/duration enforcement — not for this interactive add-in.

---
*GPS Fleet Management Solutions · Geotab Authorized Partner & Reseller*
