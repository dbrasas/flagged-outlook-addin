# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.5] - 2026-04-07

### Changed

- Split the monolithic `src/taskpane.html` into `src/taskpane.html`, `src/taskpane.css`, and `src/taskpane.js` without changing the existing UI design.
- Reworked the task pane UI to use Starting Point UI component classes on top of a Tailwind CSS v4 build pipeline while keeping the Outlook/MSAL/Graph behavior intact.
- Added Tailwind build scripts and a generated runtime stylesheet (`src/taskpane.generated.css`) so the design system works in this plain static add-in.
- Updated the task pane layout into a card-based dashboard with styled auth, loading, stats, and flagged-message states.
- Reworked the task pane to use external scripts and styles so a stricter CSP can be applied without relying on inline handlers.
- Switched MSAL browser cache persistence from `localStorage` to `sessionStorage` to reduce unnecessary token persistence.
- Added timeout handling and throttling-aware retry logic for Microsoft Graph requests.
- Added `manifest.local.xml` so local development uses a real localhost manifest instead of the production GitHub Pages manifest.
- Updated `server.js` to support the new static files, enforce `GET` and `HEAD` only, and return baseline security headers.
- Added manifest validation helper scripts to `package.json`.
- Rewrote `README.md` around the actual business purpose, deployment model, and security assumptions.
- Added `docs/AUDIT.md` with a source-backed architecture and security audit, implementation status labels, and deferred recommendations.

## [1.0.4] - 2026-04-07

### Security

- **CRITICAL OVERSIGHT**: An AI development error led to the accidental inclusion of SSL private keys (`server.key`) and certificates (`server.crt`) in a public commit. This was an extremely bad practice and an unacceptable oversight of security protocols.
- Removed all SSL certificates from Git tracking and updated the project index to ignore future certificate commits.
- Added a comprehensive `.gitignore` file to prevent the leakage of sensitive keys, environment variables, and local certificates.
- Updated `server.js` to explicitly list certificates in the ignore list and provided instructions for local certificate installation to keep keys outside of the codebase.

## [1.0.3] - 2026-04-03

### Changed

- Prepared add-in for production GitHub Pages deployment. Updated `manifest.xml` source locations (`SourceLocation`, `IconUrl`, `HighResolutionIconUrl`, `AppDomains`) to point to the live `dbrasas.github.io` URL instead of `localhost`.
- Replaced obsolete `.svg` placeholders with high-quality PNG icons (16x16, 32x32, 64x64, 80x80, 128x128) configured specifically for Office Add-in manifest validation.
- Cleaned up unneeded log outputs and temporary files (e.g., `validation.log`).

## [1.0.2] - 2026-04-02

### Fixed

- Fixed directory traversal vulnerability in `server.js` by explicitly whitelisting allowed static files and enforcing root-path boundaries.
- Reconstructed and repaired corrupted HTML structure in `src/taskpane.html`.
- Fixed auth setup state by adding an explicit guard message if the Azure Client ID placeholder is not replaced.
- Corrected `manifest.xml` to fix `Mailbox` casing and replace the obsolete `CustomPane` with the supported `MessageReadCommandSurface` extension point.
- Modified Add-in permissions from `ReadWriteMailbox` to a safer `ReadItem` baseline, as full message lookup relies on Graph auth state.

## [1.0.1] - 2026-04-02

### Added

- `.geminirules` file to strictly enforce AI development standards for changelogs.
- `CHANGELOG.md` to track project history.
- Placeholder icon files in the `assets/` directory to satisfy Outlook Manifest requirements.

### Changed

- Replaced the flawed `Office.auth.getAccessToken` mechanism with MSAL.js Single Page Application (SPA) popup flow for proper Graph API token retrieval.
- Updated `manifest.xml` source locations from the root mapping to explicitly point to the actual `/src/taskpane.html` file location.
- Updated `manifest.xml` to include an `<AppDomains>` section, whitelisting `login.microsoftonline.com` to allow MSAL popups within the Outlook sandbox.
- Refactored `src/taskpane.html` code to gracefully handle the new MSAL authentication flow.
- Changed the `README.md` setup instructions to guide the user in setting up an "SPA" platform in Azure App Registration.

### Fixed

- Fixed bug in `server.js` where requests to `/taskpane.html` returned a 404 error instead of serving the file from `/src/`.
- Fixed the TimeZone shift bug where Graph API UTC `dueDateTime` outputs were incorrectly shifted to the previous day; the request now specifies `Prefer: outlook.timezone="system local"` to enforce timezone alignment natively.
