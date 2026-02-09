## Realtime Spreadsheet App

Minimal Google-Sheets-like web app with:

- **Backend**: Node.js + Express
- **Realtime**: Socket.IO
- **Database**: SQLite (cell-level storage only)
- **Frontend**: Vanilla JS + HTML + CSS
- **Spreadsheet engine**: Luckysheet (via CDN)
- **Auth**: JWT

### Getting started

1. Install dependencies:

```bash
npm install
```

2. Optionally copy `.env.example` to `.env` and adjust:

```bash
cp .env.example .env
```

3. Run the server:

```bash
npm start
```

4. Open `http://localhost:3000` in your browser.

### Default admin

On first run a default admin is seeded:

- **username**: `admin`
- **password**: `admin123`

Use this account to create leaders and users via the API or extend the UI.

