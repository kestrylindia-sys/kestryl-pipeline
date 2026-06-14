# Kestryl RFQ Pipeline — Netlify Deployment

## Deploy in 3 steps

### 1 — Set environment variables in Netlify
Go to Site Settings → Environment Variables → Add:

| Variable | Value |
|---|---|
| `VITE_AIRTABLE_PAT` | Your Airtable Personal Access Token |
| `VITE_WORKER_URL` | `https://kestrylindia-pricing.kestryl-india.workers.dev` |

To get your Airtable PAT: airtable.com/account → Developer Hub → Personal Access Tokens
Scopes needed: `data.records:read`, `data.records:write`, `schema.bases:read`
Base access: Kestryl Pipeline Config + Operations Hub

### 2 — Connect to GitHub (easiest) or drag-drop
**Option A — GitHub (auto-deploys on every push):**
1. Push this folder to a GitHub repo
2. netlify.com → Add new site → Import from GitHub
3. Build command: `npm run build` | Publish directory: `dist`
4. Add env vars above → Deploy

**Option B — Drag & Drop (instant, one-time):**
1. Run `npm install && npm run build` locally (needs Node 18+)
2. Drag the `dist/` folder to netlify.com/drop
3. Done — you get a URL instantly

### 3 — Set password protection (recommended)
Netlify dashboard → Site settings → Access control → Password protection
One password keeps it private without building auth.

## Architecture
```
Browser (your-site.netlify.app)
  → Cloudflare Worker (live pricing — Mouser + E14 + Nexar)
  → Airtable REST API (config + RFQ logging)
  → Claude API (PDF extraction via anthropic.com)
```

## Local development
```bash
npm install
npm run dev
# Opens at http://localhost:5173
```

## Files
- `src/App.jsx` — Main pipeline application
- `src/main.jsx` — React entry point  
- `index.html` — HTML shell
- `vite.config.js` — Vite bundler config
- `netlify.toml` — Netlify build config
- `package.json` — Dependencies
