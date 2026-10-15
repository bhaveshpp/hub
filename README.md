# Hub — private launcher (repo `hub`)

Upload the contents of `site/` to the root of the GitHub repo **hub** → Pages → `https://bhaveshpp.github.io/hub/`. Repo is public but holds no links or data; cards come from your Sheet only after the PIN.

## Setup
1. New Google Sheet → Apps Script → paste `backend/Code.gs` → change `HUB_PIN` → run `setupSheets` (Cards + Groups tabs + your repo links as examples).
2. Deploy → Web app (Execute as *Me*, Access *Anyone*) → copy URL → `API_URL` in `site/index.html`.
3. Open the hub on the phone → enter PIN → browser menu → **Install / Add to Home screen**.
Changing `HUB_PIN` (and re-deploying a new version) logs every phone out.

## Sheet
- **Cards**: Title · Note · Icon (emoji) · URL (https only) · Group · Sort · Status (સક્રિય / નિષ્ક્રિય). New row = new card; no code change.
- **Groups** (optional): Group · Icon · Sort → section order and icon. Groups not listed come after, A–Z; blank group = અન્ય.
- Cards open in the same tab. List is cached on the phone (shows instantly / offline, refreshes in background; ⟳ forces it).
- Service worker caches only the hub's own files — never the apps or the API.

## Tests
`node tests/hub_test.js` (20; needs `npm i jsdom`).
