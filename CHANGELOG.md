# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
