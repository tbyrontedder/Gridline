# Gridline 0.1.0 — release test report

Formatting update validation: 2026-09-08. Results describe this build and this environment, not certification of Excel equivalence or general hardware performance.

## Results

| Suite | Result | Scope |
|---|---|---|
| Node built-in test runner | **132 passed, 0 failed** | Formula semantics, references, dependency invalidation, transactions, history, structural operations, sorting/fills, JSON, CSV, ZIP/CRC, XLSX emission and axis geometry |
| Browser integration | **45 checks passed** | Real DOM/input interactions in Chromium 149.0.7827.53; no uncaught page errors |
| Standalone JavaScript syntax | Passed | `node --check dist/gridline.js` |
| XLSX self-round-trip | Passed | All populated sample-cell values compared, 3 sheets, 24 merges, dimensions, pane freezes and named ranges |
| DEFLATE-compressed XLSX import | Passed | Python ZIP-compressed sample imported through the browser's native decompressor |
| Unsupported date mode | Passed | 1904-date-system fixture rejected with an explicit error |
| GPU pipeline execution | **Not verified** | The test harness exposed an opaque origin with no WebGPU API; Canvas2D fallback was the executed renderer |
| Persistent browser storage | **Not verified** | Snapshot writes/restoration exercised through an in-memory Storage fixture |
| Microsoft Excel interoperability | **Not verified** | No Excel process opened the exported workbook during this validation |
| GitHub CI execution | **Not run** | Workflow is included in source, ready for a repository |

## Browser interaction coverage

The suite boots the bundled HTML through Playwright, checks the fictional demo total and chart overlays, navigates through the address box, types a numeric value, commits it, checks dependent recalculation, undoes/redoes, edits a formula through the formula bar, adds/renames a sheet, dispatches an actual clipboard paste event, fills formulas downward, uses toolbar copy/paste through the internal clipboard fallback, applies bold to a range, freezes panes, adds a note, sorts with a preserved header, applies and clears a value filter, replaces text, exports/imports XLSX, exercises native DEFLATE decompression, tests explicit date-mode rejection, and toggles theme/zoom. Formatting checks exercise right-click row/column selection, sparse defaults, the Format Cells dialog and keyboard shortcut, readable date editing, invalid-date rejection, Cancel, unchanged-value preservation, and native/XLSX round-trips for dates, times, currencies, row defaults and General overrides. New entries after loading inherit their column formats. Security regressions reject malicious decimal settings in cell, row and column styles before a file replaces the active workbook, and verify that the formatting dialog cannot interpret a hostile precision value as HTML even when import validation is bypassed.

The synthetic native paste event verifies the application's paste-event handler and parsing. It does not verify operating-system clipboard permissions or interoperability with every external application. Tests of in-app clipboard snapshots do not imply that asynchronous system clipboard access was granted.

A browser test exposed an Enter bubbling issue in the address box: Enter selected the requested cell and then bubbled into worksheet navigation, moving down one more row. The handler now stops propagation, and subsequent browser runs pass. Unit tests additionally exercise defined-name retargeting during worksheet rename and prevention of imported JSON overriding Sheet methods.

## Stress exercise

The browser created **80,008 stored cells**: 10,000 data rows × 8 columns plus headers. This was not 80,008 DOM elements. It navigated to **D10001**, evaluated that row's revenue and contribution, and displayed a viewport containing **576 cells**.

The machine-readable report records creation wall time and last-frame CPU time for the current run. These are diagnostic observations, not reproducible benchmark claims, GPU timing, steady-state FPS or comparisons with Excel. The full sheet's 50,000 formula cells were not all forced into an exhaustive recalculation benchmark.

## Reproduce

```sh
npm test
npm run build
python tests/browser-smoke.py --chromium /usr/bin/chromium
```

For a served-origin run on a host with an available WebGPU adapter:

```sh
# Terminal 1
npm start
# Terminal 2
python tests/browser-smoke.py --url 'http://localhost:8080/?fresh=1' --chromium /path/to/chromium
```

Inspect the report's actual `backend` field. A Canvas2D run is not a successful GPU validation run. Capture browser GPU validation errors, visually inspect glyph output, exercise device loss on supported hardware, and profile CPU/GPU separately before making GPU performance guarantees.

## Evidence files

- `unit-test-results.txt`: output from the Node test run.
- `browser-test-results.json`: machine-readable successful browser checks and diagnostic observations.
- `browser-test-output.txt`: browser test console summary.
- `preview.png`: screenshot captured from the running application, not a design mockup. The status bar correctly shows Canvas2D fallback.

## Remaining validation work

Cross-browser/OS coverage, real IME sessions, bidirectional and complex-script text, assistive-technology workflows, extensive random formula differential testing against Excel, adversarial file fuzzing, GPU driver/device matrices, real storage quota/failure behavior, actual OS clipboard permissions, and printing are not covered by this release run. The published limits and unsupported features are documented in the README and architecture notes rather than hidden behind passing smoke tests.
