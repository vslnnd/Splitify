# Splitify

**Split Excel and CSV files based on a value, row condition, or rule you define.**

Splitify scans your file, finds the condition you set, and splits it into separate output files automatically. Save your split configurations as profiles and reuse them instantly — no repeating setup every time.

> Part of the [Vellova Apps](https://github.com/vslnnd) suite.

---

## What it does

You load a file, define a split condition (a specific column value, row range, delimiter, etc.), and Splitify produces the output files. You can save that configuration as a named profile and load it again next time in one click.

**Key features:**

- Split by value, row, or custom condition
- Profile management — save, load, and switch between configurations
- Auto-updates — new versions install in the background
- macOS and Windows support

---

## Installation

Download the latest release for your platform from the [Releases](../../releases) page.

| Platform | File |
|----------|------|
| Windows | `.exe` installer |

No setup required — download, install, open.

---

## Development

**Requirements:** Node.js 18+

```bash
# Install dependencies
npm install

# Run in development
npm start
```

## Building a Release

1. Copy `.env.example` to `.env` and add your GitHub token:
   ```
   GH_TOKEN=ghp_xxxxxxxxxxxx
   ```
2. Bump the version in `package.json` (e.g. `1.4.17` → `1.4.18`)
3. Build and publish:
   ```bash
   npm run electron:build
   ```
   This builds the installer and publishes a GitHub Release automatically.

## Build Without Publishing (local test only)

```bash
npm run electron:build:local
```

---

## Auto-updates

Splitify checks for a new version 3 seconds after launch. If one is available, a banner appears at the top of the app. It downloads in the background — when ready, a "Restart & Install" button appears. One click and the app restarts with the new version.

---

## Stack

- [Electron](https://www.electronjs.org/)
- HTML / CSS / JavaScript

---

## License

See [LICENSE](./LICENSE) for details.

---

*Built by [Nenad](https://github.com/vslnnd) · Vellova Apps*
