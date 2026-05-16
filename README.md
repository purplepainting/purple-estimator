# Purple Estimator

Vite + React + Supabase + Vercel scaffold around `src/JobWalkProposal.jsx`, a 3,400-line React component built originally as a Claude.ai artifact.

## Stack
- Vite + React 18
- Supabase (email magic-link auth + Postgres persistence)
- Vercel (static hosting + a single serverless function that proxies to Anthropic)
- xlsx, mathjs

## Local setup

1. `npm install`
2. `cp .env.example .env.local` and fill in the two `VITE_SUPABASE_*` values
3. Open the Supabase dashboard -> SQL Editor -> paste `supabase/schema.sql` -> Run
4. `npm run dev`

`ANTHROPIC_API_KEY` is only needed when you want the chat features to actually call the model — for pure UI work locally, it can be left unset and `/api/chat` will return 500.

## Deploy to Vercel

1. Push to GitHub.
2. Import the repo in Vercel.
3. In the Vercel project settings -> Environment Variables, set:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `ANTHROPIC_API_KEY` (server-side only — **do not** prefix with `VITE_`)
4. Deploy.

## How the pieces fit

- `src/JobWalkProposal.jsx` is unchanged from the artifact.
- `src/main.jsx` monkey-patches `window.fetch` so the component's direct calls to `https://api.anthropic.com/v1/messages` get rerouted through `/api/chat`, which adds the server-side `x-api-key` and `anthropic-version` headers.
- `src/storage.js` swaps the artifact's `window.storage` for Supabase. The three known keys (`scope_library_v3`, `catalog_ids_v1`, `tier_multipliers_v1`) live in dedicated singleton tables visible to the whole team; every other key falls into per-user rows in `kv_store`.
- `src/chatStore.js` is the persistence layer for chat sessions/messages, used wherever the app needs cross-session chat history.
