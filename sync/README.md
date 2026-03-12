# Sync Folder

This folder contains the Cloudflare Worker scaffold for future NTN database sync.

## Contents
- `cloudflare/worker.js` - API worker scaffold
- `cloudflare/wrangler.toml` - Wrangler config template

## Planned endpoints
- `GET /health` -> service check
- `GET /db` -> load NTN database
- `POST /db` -> save NTN database

## Notes
This scaffold is added now, but the website app is not yet wired to call these endpoints.
That can be done later when deploying on Cloudflare.
