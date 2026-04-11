# Changelog

## [1.4.0] - 2026-04-11

### Added
- Comprehensive README.md with badges, quick start, and full documentation
- Build pipeline with esbuild, clean-css, html-minifier-terser (34% size reduction)
- CI/CD via GitHub Actions: lint → build → validate manifest → deploy to Pages
- Complete i18n coverage (42% → 100%) — all 75+ UI strings translated
- `data-i18n-placeholder` and `data-i18n-title` support in applyTranslations()
- `package.json` with build, lint, dev, and clean scripts

### Fixed
- Dark mode contrast: `--text-muted` #888 → #a0a0a0, `--text-secondary` #bbb → #c8c8c8 (WCAG AA compliant)
- `--text-primary` #e0e0e0 → #e8e8e8 for better readability
- `--border-color` #444 → #4a4a4a for clearer separation
- Duplicate Spanish translation keys (noFileSelected, orDivider)

## [1.3.0] - 2026-04-11

### Security
- Added DOMPurify for HTML sanitization (XSS prevention)
- Wrapped all JSON.parse calls in safe try/catch with fallback defaults
- Removed inline onclick handler (CSP compliance)
- Added ARIA roles to all modal dialogs
- Added form validation (type="email") on CC/BCC/unsubscribe fields

### Added
- Real rate limiting engine (token bucket, 30/min enforced)
- Retry logic with exponential backoff for Graph API calls (500, 502, timeout)
- 30-second request timeout on all Graph API calls
- Send progress checkpointing to localStorage (resume on page refresh)
- Retry-After header support for 429 responses
- `.gitignore`, `LICENSE` (MIT), and `CHANGELOG.md`

### Fixed
- MSAL interaction_in_progress stuck state (proactive clearing)
- Dark mode toggle button not working
- appState.fileName not set on upload
- Missing i18n data attributes on step labels

### Improved
- Accessibility: ARIA labels, roles, aria-modal, aria-live regions
- Content Security Policy readiness (no inline handlers)
- Form inputs: proper type="email", required attributes, aria-labels
- Error messages now escaped (XSS-safe)

## [1.2.0] - 2026-04-11

### Fixed
- 5 bugs found by QA integration testing
- MSAL single-tenant authority configuration

## [1.0.0] - Initial Release
- 44 features: mail merge, A/B testing, scheduling, templates, etc.
- MSAL authentication with Microsoft Graph API
- 6-language UI (EN, DE, FR, ES, JA, ZH)
- Dark mode with OS detection
