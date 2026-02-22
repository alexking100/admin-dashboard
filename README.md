# Admin Tasks Dashboard

A local Python dashboard that scans Google Docs for to-do items, displays them in a browser UI with categories, and syncs completions/deletions back to the doc.

## Features

- Scans `.gdoc` and `.docx` files for tasks using Claude Haiku (LLM) or rule-based fallback
- Auto-categorises tasks into Holiday / Work / Finances / Other
- Marks items complete or removed directly in the source Google Doc
- Google Calendar integration — add tasks as events from the dashboard
- Inline editing synced back to the Google Doc
- Drag-and-drop category management

## Setup

### 1. Install dependencies

```bash
pip3 install -r requirements.txt
```

### 2. Google OAuth credentials

- Go to [Google Cloud Console](https://console.cloud.google.com)
- Create a project and enable the **Google Docs API**, **Google Drive API**, and **Google Calendar API**
- Create an OAuth 2.0 Desktop App credential
- Download the JSON and save it as `credentials.json` in this folder (use `credentials.example.json` as a reference for the expected shape)

### 3. Anthropic API key (optional — enables LLM task extraction)

Create a `.env` file in this folder:

```
ANTHROPIC_API_KEY=your_key_here
```

Get a key at [console.anthropic.com](https://console.anthropic.com). Without it the dashboard falls back to rule-based task extraction.

### 4. Run

```bash
python3 dashboard.py
```

The dashboard opens automatically at `http://localhost:5678`. On first run a browser tab will open for Google sign-in.

## Files not in this repo

These are gitignored and must be created locally:

| File | Why excluded |
|---|---|
| `credentials.json` | Google OAuth client secret |
| `token.json` | Google OAuth token (auto-generated on first run) |
| `.env` | Contains your Anthropic API key |
| `dashboard_tasks.json` | Your personal task data (auto-generated) |
