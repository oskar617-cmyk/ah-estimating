# AH Estimating — Classifier Worker (Cloudflare)

A small Cloudflare Worker that proxies Google Gemini for the AH Estimating PWA.
Holds the Gemini API key as a Cloudflare Worker secret so it never reaches the browser.

## Endpoint

POST `/` (Worker root URL) with JSON body:

```json
{ "task": "classify" | "extractAmount" | "summarizeFilename", "payload": { ... } }
```

Returns: `{ "result": { ... } }` on success, `{ "error": "..." }` on failure.

## Secrets / Variables (set in Cloudflare dashboard)

| Setting          | Where                | Value                                                  |
|------------------|----------------------|--------------------------------------------------------|
| `GEMINI_API_KEY` | Worker → Secrets     | Your Gemini API key from <https://aistudio.google.com> |
| `CORS_ORIGINS`   | Worker → Variables   | `https://oskar617-cmyk.github.io`                      |

## Deploy

Easiest path: paste the contents of `worker.js` into Cloudflare's Quick edit UI on the Worker page.

## Free tier

Cloudflare Workers free plan: 100,000 requests/day, 10ms CPU per request.
Gemini calls are fetch-bound — wait time isn't billed as CPU — so we fit comfortably.
