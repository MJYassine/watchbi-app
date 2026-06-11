# Horometrics — WatchBI

Luxury watch inventory & sales intelligence with an AI chat assistant powered by Google Gemini.

## Getting a free Gemini API key

1. Go to [Google AI Studio](https://aistudio.google.com/app/apikey).
2. Sign in with your Google account.
3. Click **Create API key** → copy the key.
4. The free tier (Gemini 2.5 Flash) has generous quotas — no billing required.

## Setup

```bash
cd watchbi-app
cp .env.example .env
# Paste your key into .env:
# GEMINI_API_KEY=AIza...
npm install
```

## Running

```bash
npm start          # starts Express (port 3001) + Vite (port 5173) together
```

Or separately:

```bash
npm run server     # Express API server  → http://localhost:3001
npm run dev        # Vite dev server     → http://localhost:5173
```

Open **http://localhost:5173** in your browser.
