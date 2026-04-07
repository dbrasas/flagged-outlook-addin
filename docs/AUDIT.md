# Audit Summary

Date: 2026-04-07

## Business purpose

This add-in exists to solve one specific workflow problem: users need a fast way to review all flagged mail across all folders, sorted by follow-up due date, without manually switching folders or building custom Outlook views. That purpose justifies a minimal Outlook command surface, a read-focused Microsoft Graph integration, and a lightweight task pane UI.

## Current implementation summary

- Outlook mail add-in with a task pane button on the message read surface.
- Static client-side UI hosted from GitHub Pages in production and from a local HTTPS server in development.
- MSAL browser popup flow for Microsoft Graph delegated access.
- Microsoft Graph query to list flagged messages, then client-side grouping into:
  `Overdue`, `Today`, `Upcoming`, `No due date`.
- No backend API, database, or server-side secret handling.

## Architecture assessment

### Already present

- `ReadItem` manifest permission instead of a broader mailbox permission.
- No client secret in the codebase.
- Basic static-file allowlisting in the local server.
- Public certificate exposure incident was already corrected and documented in the changelog.

### Implemented now

- Split `src/taskpane.html` into:
  `src/taskpane.html`, `src/taskpane.css`, `src/taskpane.js`
- Removed inline event handlers and inline script blocks so the task pane can use a stricter CSP.
- Added a local-only manifest: [manifest.local.xml](/c:/laragon/www/blank/manifest.local.xml)
- Aligned local-development instructions with an actual localhost manifest instead of the production manifest.
- Added basic response hardening in the local server:
  CSP, `Referrer-Policy`, `X-Content-Type-Options`, `Permissions-Policy`, `Cache-Control`
- Added Graph timeout handling and retry behavior for throttling-related failures.
- Switched MSAL cache persistence from `localStorage` to `sessionStorage`.

### Should be improved

- Auth model is still popup-based SPA auth rather than Nested App Authentication.
- Production hosting still relies on GitHub Pages, which limits response-header control.
- There is still no automated lint/test/build pipeline.
- CDN-hosted dependencies are still loaded directly instead of being pinned through a package lock and bundled build.

### Deferred

- Migration to Nested App Authentication (NAA).
- Scope reduction from `Mail.Read` to `Mail.ReadBasic`.
- Dedicated lightweight auth redirect page for popup/NAA flows.
- Production observability and privacy-reviewed telemetry.

## Prioritized risks and recommendations

| Priority | Area | Status | Details |
|---|---|---|---|
| High | Auth modernization | Deferred | Current Microsoft guidance favors NAA for supported Outlook hosts. The current popup flow works, but it is no longer the preferred long-term model. |
| High | Production header control | Deferred | GitHub Pages is convenient, but it cannot enforce the full response-header posture you would want for a hardened production add-in. |
| Medium | Least-privilege Graph scope | Deferred | The add-in currently asks for `Mail.Read`. A review should verify whether `Mail.ReadBasic` is sufficient for the required message fields and tenant consent path. |
| Medium | Supply-chain control | Deferred | `office.js` and `msal-browser` are loaded from Microsoft CDNs. This is common, but package-managed and version-pinned delivery is more controllable. |
| Medium | Validation automation | Implemented now | Added manifest validation scripts and a dedicated local manifest to reduce release mistakes. |
| Medium | Token persistence | Implemented now | Tokens now use `sessionStorage`, reducing unnecessary persistence on shared machines. |
| Medium | Graph resilience | Implemented now | Timeout and `Retry-After` aware retry logic reduce brittle behavior during throttling or transient failures. |
| Low | Project structure | Implemented now | Task pane code is no longer monolithic, which improves maintainability and CSP enforcement. |

## Security checklist for this add-in

### Implemented now

- Keep Outlook manifest permission at `ReadItem`.
- Keep Graph access delegated and user-scoped.
- Do not store secrets in the browser.
- Store tokens in `sessionStorage` instead of `localStorage`.
- Use external scripts and styles so CSP can block inline execution.
- Keep the local dev server on an allowlist-only model.
- Use HTTPS locally and in production.
- Do not commit certificates or key material.
- Use `noopener,noreferrer` when opening external message links.
- Avoid logging access tokens or message content.

### Recommended next

- Add NAA support with fallback detection for older hosts.
- Reconfirm whether `Mail.ReadBasic` can replace `Mail.Read` for this scenario.
- Move production hosting to a platform that can set response headers explicitly.
- Add release validation and dependency review to CI.
- Add privacy-reviewed diagnostics that exclude message bodies, tokens, and sensitive mailbox metadata.

## Proposed project structure

```text
.
├── assets/
├── certs/
├── docs/
│   └── AUDIT.md
├── manifest.local.xml
├── manifest.xml
├── package.json
├── server.js
└── src/
    ├── taskpane.css
    ├── taskpane.html
    └── taskpane.js
```

If the add-in grows, the next clean split should be:

- `src/auth/`
- `src/graph/`
- `src/ui/`
- `src/config/`

## Documentation gaps fixed

The updated README now explains:

- why the app exists
- the difference between local and production manifests
- how the Azure app registration maps to the current auth flow
- which redirect URIs are required
- which permission is currently requested and why it remains unchanged for now
- how to validate manifests
- which security assumptions apply to this project

## Validation against current documentation

### Confirmed by current Microsoft documentation

- Outlook add-ins can authenticate in several ways, and NAA is the recommended SSO path for supported Outlook hosts:
  https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/authentication
- Enabling NAA requires specific SPA redirect configuration and `createNestablePublicClientApplication`:
  https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- Microsoft recommends requesting minimum scopes needed for the task:
  https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- MSAL browser caching guidance documents the tradeoff between `sessionStorage`, `localStorage`, and memory storage:
  https://learn.microsoft.com/en-us/entra/msal/javascript/browser/caching
- Outlook manifest permissions should follow least privilege, and `ReadItem` is the lower mail-read baseline inside the add-in manifest:
  https://learn.microsoft.com/en-us/javascript/api/manifest/permissions
- Microsoft recommends validating Office add-in manifests with `office-addin-manifest`:
  https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest
- Microsoft Graph guidance treats `Retry-After` handling as the correct throttling recovery behavior:
  https://learn.microsoft.com/ko-kr/previous-versions/azure/ad/graph/howto/azure-ad-graph-api-throttling

### Inferred from general best practices

- Externalizing scripts/styles before tightening CSP.
- Using `Referrer-Policy: no-referrer` for a mailbox-oriented add-in.
- Adding `X-Content-Type-Options: nosniff`.
- Adding a restrictive `Permissions-Policy` on the local dev server.
- Preferring a host with production response-header control over static GitHub Pages when the app moves beyond a lightweight deployment.

### Recommended but not implemented

- Full NAA migration.
- CI-based manifest validation.
- Dependency pinning and bundling for `msal-browser`.
- Auth redirect page isolation.

## MCP and research note

The user asked for MCP-backed validation, including `perplexity-ask` or an equivalent. In this workspace there was no relevant research connector available through MCP discovery, so up-to-date validation was done with official Microsoft documentation through the available web tool instead. That means the recommendations above are grounded in current Microsoft sources, but not via a dedicated Perplexity MCP connector in this session.
