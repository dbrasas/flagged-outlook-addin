# Flagged by Due Date

Outlook add-in that helps a user triage flagged email across the entire mailbox by grouping messages into `Overdue`, `Today`, `Upcoming`, and `No due date`.

## Why this add-in exists

The business purpose is narrow and practical: Outlook already lets a user flag messages, but it does not provide a focused side panel that lists all flagged mail from all folders sorted by flag due date. This add-in exists to reduce missed follow-ups and make due-date-based email review faster than manually hunting through folders or using generic Outlook views.

That purpose drives the current design:

- Read-only Outlook add-in command in message read mode.
- Microsoft Graph is used only to read flagged messages for the signed-in user.
- The add-in stays lightweight and client-side because it only needs mailbox read access and simple sorting/grouping.

## Current architecture

- `manifest.xml`
  Production Outlook add-in manifest for the GitHub Pages deployment.
- `manifest.local.xml`
  Local-development manifest that points Outlook to `https://localhost:3000`.
- `server.js`
  Minimal HTTPS dev server with a strict file allowlist and baseline security headers.
- `src/taskpane.html`
  Task pane shell that applies Starting Point UI component classes.
- `src/taskpane.css`
  Tailwind CSS v4 source file. Imports `tailwindcss` and `starting-point-ui`, defines theme tokens, and contains the app-specific component styling.
- `src/taskpane.generated.css`
  Compiled stylesheet served by the add-in at runtime.
- `src/taskpane.js`
  MSAL auth flow, Microsoft Graph access, grouping/sorting logic, and DOM rendering.

## Design system integration

The task pane now uses [Starting Point UI](https://www.startingpointui.com/) on top of Tailwind CSS v4. In this repository that requires a small build step because the add-in is a plain static HTML/CSS/JS app, not a React or Next.js project.

Installed build-time packages:

- `tailwindcss`
- `@tailwindcss/cli`
- `starting-point-ui`

Key implementation detail:

- `src/taskpane.css` is the source stylesheet.
- `npm run build:styles` compiles it into `src/taskpane.generated.css`.
- `src/taskpane.html` loads the generated file, not the source file.
- `npm run dev` and `npm start` now rebuild the generated stylesheet before starting the HTTPS server.

## Authentication model

The current implementation uses popup-based MSAL SPA authentication against Microsoft Entra ID and then calls Microsoft Graph directly from the task pane.

Important notes:

- The Azure Application (client) ID is a public identifier, not a secret.
- No client secret should ever be added to this project.
- Access tokens are now cached in `sessionStorage` instead of `localStorage` to reduce persistence risk on shared devices.
- Current code keeps the existing `Mail.Read` scope to avoid an unreviewed breaking auth change for the already-registered Azure app.
- Current Microsoft guidance recommends Nested App Authentication (NAA) for supported Outlook hosts. That migration is documented in [docs/AUDIT.md](docs/AUDIT.md) as a recommended follow-up, not an in-place auth rewrite.

## Azure app registration

Configure the existing Azure app registration with:

1. Supported account types:
   `Accounts in any organizational directory and personal Microsoft accounts`
2. SPA redirect URIs for the current popup flow:
   `https://localhost:3000/src/taskpane.html`
   `https://dbrasas.github.io/flagged-outlook-addin/src/taskpane.html`
3. Microsoft Graph delegated permissions:
   `Mail.Read`

Notes:

- `Mail.Read` is what the current implementation requests today.
- A later least-privilege review should evaluate moving to `Mail.ReadBasic` if tenant consent and required message fields are confirmed for this exact scenario.
- If you change the Azure client ID, update `CLIENT_ID` in [src/taskpane.js](src/taskpane.js).

## Local development

Prerequisites:

- Windows 10/11 with new Outlook or Outlook on the web.
- Node.js 18+.
- Microsoft 365 account.

### 1. Install dev certificates

```bash
npm run certs:install
mkdir certs
copy "%USERPROFILE%\.office-addin-dev-certs\localhost.key" certs\server.key
copy "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" certs\server.crt
```

### 2. Install packages and build styles

```bash
npm install
npm run build:styles
```

If you are actively editing Tailwind utilities or Starting Point UI markup, run a watcher in a second terminal:

```bash
npm run watch:styles
```

### 3. Start the local HTTPS server

```bash
npm run dev
```

Expected output:

```text
✅ Add-in serveris veikia: https://localhost:3000
📋 Local manifest: https://localhost:3000/manifest.local.xml
🔧 Taskpane: https://localhost:3000/src/taskpane.html
```

### 4. Sideload the local manifest

Use `manifest.local.xml` for local work.

- Outlook Web / new Outlook:
  Open `https://aka.ms/olksideload`
- Choose `Add a custom add-in`
- Upload `manifest.local.xml`

## Production deployment

The production manifest currently points to GitHub Pages:

- Task pane: `https://dbrasas.github.io/flagged-outlook-addin/src/taskpane.html`
- Manifest: `https://dbrasas.github.io/flagged-outlook-addin/manifest.xml`

Production guidance:

- Keep the hosting endpoint HTTPS-only.
- Treat GitHub Pages as acceptable for a small static add-in, but move to a host that supports managed response headers if stricter production CSP/HSTS policy control is required.
- Validate the manifest before release.
- Keep Azure redirect URIs aligned with the exact deployed task pane URL.

## Security and operations

Already implemented in the repository:

- `certs/` and certificate files are ignored by Git.
- The local dev server serves only an allowlisted set of files.
- The add-in manifest permission is limited to `ReadItem`.
- Task pane scripting is now externalized so the page can use a stricter CSP.
- Graph calls now honor timeouts and retry throttling-related responses.

Operational rules:

- Never commit private keys, `.env` files, or bearer tokens.
- Never add a client secret to a browser-based Outlook add-in.
- Keep permissions least-privileged and review them before widening scope.
- Prefer production monitoring that avoids logging mailbox content or message subjects unless explicitly justified and approved.

## Validation Helpers

```bash
npm run build:styles
npm run validate:manifest
npm run validate:manifest:local
```

These scripts use `npx office-addin-manifest validate ...` as recommended by current Microsoft documentation.

Quick UI verification:

- Confirm `src/taskpane.generated.css` exists after the build step.
- Open `https://localhost:3000/src/taskpane.html` and verify the page loads with the new card-based dark theme.
- Check that the buttons, badges, separators, and cards render with Starting Point UI styling.
- Sign in and verify flagged messages render as styled cards grouped by `Vėluoja`, `Šiandien`, `Artimiausi`, and `Be termino`.

## Common Issues & Troubleshooting

**1. Desktop Add-in opens with blank white space**
If the add-in pane appears blank or distorts its size pushing content out of view on desktop clients, it's typically caused by viewport vertical height units (e.g., `100vh` or Tailwind `min-h-screen`). 
- **Fix:** Desktop's Edge WebView often miscalculates `100vh`. Constrain the `html` and `body` strictly using `h-full overflow-hidden`, and apply explicitly contained scrolling (`h-full overflow-y-auto`) to the topmost application wrapper `div`.

**2. Outlook throws "Add-in Error" asking to Retry/Start**
If the add-in strictly refuses to load initially, showing an error screen, but works seamlessly after clicking "Retry", the Office host timed out waiting for `Office.onReady()` to resolve.
- **Fix 1 (CSP):** `office.js` silently fetches `MicrosoftAjax.js` from `ajax.aspnetcdn.com` in OWA environments. Ensure your `taskpane.html` Content-Security-Policy `script-src` directive explicitly allows `https://ajax.aspnetcdn.com`. A blocked script permanently halts Office initialization.
- **Fix 2 (Lifecycle Race Conditions):** Do not manually assign an empty placeholder to `Office.initialize` before calling `Office.onReady`. Defining both creates a host race condition that fails the strict initialization checks.

## Project structure

```text
.
├── assets/
├── certs/
├── docs/
│   └── AUDIT.md
├── manifest.local.xml
├── manifest.xml
├── package.json
├── package-lock.json
├── server.js
└── src/
    ├── taskpane.css
    ├── taskpane.generated.css
    ├── taskpane.html
    └── taskpane.js
```

## Audit and source-backed recommendations

See [docs/AUDIT.md](docs/AUDIT.md) for:

- business-purpose summary
- architecture audit
- prioritized risks
- security checklist
- current-doc-validated recommendations
- deferred items that need Azure or hosting decisions
