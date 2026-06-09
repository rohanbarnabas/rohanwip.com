# Gut Feelings

Personal meal planning and gut health tracking app.

## Setup

1. Copy `.env.example` to `.env` and add your Anthropic API key:
```
VITE_ANTHROPIC_API_KEY=your_key_here
```

2. Install and run:
```
npm install
npm run dev
```

## Deploying on Vercel

1. Connect this repo to Vercel
2. Add `VITE_ANTHROPIC_API_KEY` in Vercel → Settings → Environment Variables
3. Deploy

Data is stored in localStorage on the device.
