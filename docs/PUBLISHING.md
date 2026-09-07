# GitHub Pages deployment

The published application is https://wieslawsoltes.github.io/Gridline/.

The Pages workflow tests the engine, builds the standalone application, and publishes only `dist/` using GitHub's official Pages actions. Pushes to `main` deploy automatically; `workflow_dispatch` supports manual deployment. Pull requests run validation without deployment credentials.

Repository Settings → Pages → Build and deployment must use **GitHub Actions**. The deployment job uses the `github-pages` environment and grants only `pages: write` and `id-token: write`, with read-only repository access. Build jobs do not receive deployment credentials.

```sh
npm test
npm run build
```

There is no runtime dependency installation and no application server. The artifact contains all styles, scripts, and sample workbook inline. Project-path hosting under `/Gridline/` therefore requires no URL rewriting. A `.nojekyll` marker is included in the published artifact.

Local autosave is scoped to the browser origin. The app's status bar reports the renderer actually initialized; deployment success does not prove hardware WebGPU execution.
