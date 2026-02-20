# Splitify

Excel file splitter with profile management and auto-updates.

## First-time Setup

```bash
npm install
npm start          # test the app
```

## Building & Publishing a Release

1. Copy `.env.example` to `.env` and paste your GitHub token:
   ```
   GH_TOKEN=ghp_xxxxxxxxxxxx
   ```

2. Bump the version in `package.json` (e.g. `1.0.0` → `1.1.0`)

3. Build and publish:
   ```bash
   npm run electron:build
   ```
   This creates the installer in `dist/` AND publishes a GitHub Release automatically.

4. Users with the app installed will see the update banner next time they open Splitify.

## Build Without Publishing (local test only)

```bash
npm run electron:build:local
```

## How Updates Work

- App checks GitHub for a new version 3 seconds after launch
- If found, a green banner appears at the top: "Downloading update..."
- Progress bar fills as it downloads in the background
- When ready: "Restart & Install" button appears
- User clicks → app restarts with the new version installed
