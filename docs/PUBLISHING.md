# GitHub Pages deployment

The published application is https://wieslawsoltes.github.io/Gridline/.

## Automatic publishing

`.github/workflows/pages.yml` runs on every push to `main` and supports manual `workflow_dispatch`. It runs the unit tests, builds the standalone application, checks JavaScript syntax, runs the browser integration suite, and uploads only `dist/` using GitHub's official Pages actions. The deployment job then verifies that the public HTTPS response has exactly the SHA-256 fingerprint of the tested build.

`.github/workflows/ci.yml` validates pull requests and non-main branches without deployment permissions. Official actions are pinned to verified commit SHAs. The application has no npm runtime dependencies; Python Playwright is installed only into an isolated temporary environment for browser tests.

Repository Settings > Pages > Build and deployment uses **GitHub Actions**. The deployment job uses the `github-pages` environment with `pages: write`, `id-token: write`, and read-only repository access. Build and pull-request jobs do not receive deployment credentials. Concurrent Pages deployments are serialized.

## Reproduce locally

```sh
npm test
npm run build
node --check dist/gridline.js
# Linux: installs isolated development-only browser dependencies.
bash scripts/browser-tests.sh
```

There is no application server in production. The generated HTML contains its styles, scripts, and sample workbook inline, so project-path hosting under `/Gridline/` requires no URL rewriting. A `.nojekyll` marker is included in the published artifact.

Local autosave is scoped to the browser origin. The status bar reports the renderer actually initialized. Browser integration tests use an opaque-origin harness with an in-memory Storage fixture; deployment success does not certify hardware WebGPU execution or persistent browser storage.

See `GITHUB_VALIDATION.md` for the initial GitHub test run and the Actions run summaries for subsequent deployment verification.
