# Gridline 0.1 — architecture and implementation contract

## 1. Document model

A workbook owns ordered worksheets, a stable active-sheet ID, workbook-scoped names, calculation state, and history. A worksheet owns a sparse `Map` whose key is `"row,column"`. Public coordinates are zero-based; user-facing A1 coordinates are converted at the boundary. Cells store raw input separately from calculated values. Raw strings beginning with `=` enter formula evaluation; apostrophe-prefixed input remains literal.

A cell can carry a style and local note even when its input is empty. Dimensions, merges, chart definitions, conditional rules, frozen panes and filtering are worksheet metadata. Numeric dimension overrides are separate maps. JSON serialization expands maps/sets into arrays; reconstruction restores their concrete types. Imported worksheet metadata is whitelisted so it cannot replace worksheet methods. This is basic validation, not an assertion that every adversarial JSON shape has been security-audited.

`Workbook.transaction(label, fn)` records cell before-images once per touched key. Successful completion captures after-images, clears the redo branch, invalidates dependents and emits a change. Exceptions restore the transaction's before-images. Cell edits also invalidate affected cache entries immediately so reads inside an edit do not intentionally rely on old cached results.

`Workbook.mutate(label, fn)` handles structural/metadata operations with before/after document snapshots. Snapshot history is simple and comprehensive but is a deliberate memory and latency tradeoff. History retains at most 100 commands, not a byte-budgeted number of bytes. Consumers should avoid creating enormous metadata commands. There is no disk journal or crash-safe multi-version document store.

After snapshot undo/redo, sheet instances may be replaced. Extensions should retain sheet IDs and resolve them again instead of retaining a `Sheet` object indefinitely across structural history operations.

## 2. Formula language and evaluation

The tokenizer handles numeric/string literals, quoted worksheet names, error literals, identifiers, operators, parentheses and reference syntax. A Pratt parser produces an AST. Exponentiation and unary precedence follow the implementation's tested spreadsheet conventions; comparison and concatenation are handled explicitly. No expression is compiled as JavaScript.

The evaluator uses typed JavaScript values plus `FormulaError` and `RangeValue`. Numbers are IEEE-754 doubles. A range has shape and coordinates and is materialized subject to the operation cap. Whole-column references evaluate to a bounded used-range extent, but register a dependency over the logical column span so editing a previously empty later row still invalidates the formula. This distinction prevents the usual stale-total bug in `SUM(A:A)` after appending data.

ASTs are cached by raw formula string. Formula results are cached by sheet ID plus coordinate. During evaluation, the current formula records each referenced cell; a reverse graph maps a cell to formulas that depend on it. Editing a cell walks reverse edges transitively. Existing edges are disconnected before reevaluating a formula so replacing `=A1` with `=B1` stops treating A1 as a dependency. A separate range-dependency collection handles logical whole-column regions. Volatile function-containing formulas are invalidated on edits and explicit recalculation.

`IF`, `IFS`, `IFERROR`, `IFNA`, `CHOOSE` and error predicates have dedicated lazy evaluation. Unselected branches do not execute, so `IF(FALSE,1/0,42)` returns 42. Cycles use an active-evaluation set and produce `#CYCLE!`. The stack guard bounds recursion. Formula evaluation runs synchronously and on demand; this is not a parallel worker scheduler or a topologically sorted whole-workbook execution engine.

References preserve absolute/relative row and column flags. Copy/fill translates relative components; structural insertion/deletion translates supported numbered references irrespective of absolute flags. Token-span replacement avoids rewriting A1-like text inside strings. Sheet rename rewrites cell formulas and defined-name targets. Structural editing is intentionally conservative about dependent layout metadata and does not implement every Excel reference category.

## 3. Viewport geometry and drawing

The renderer does not allocate a row object or DOM element per logical row. `AxisLayout` represents a default cell extent plus sorted overrides. Prefix deltas calculate offsets; binary searching translates an offset back into a row or column. Hidden rows become zero-sized axis entries. The full logical scroll range is mapped onto custom scrollbar rails, avoiding reliance on an enormous physical DOM scroller.

A frame derives visible body and frozen-band spans. It emits an ordered display list of rectangles and text with explicit rectangular clip bounds. Cell fills, conditional fills/bars, text, headers, selection outlines, copy outlines and the fill handle preserve paint order. Only visible cells are considered for cell drawing. Chart overlays are anchored to worksheet coordinates and moved with the same view transform.

Canvas2D consumes the display list through clipped rectangle and text operations. It also provides the fallback when adapter acquisition, shader/pipeline setup, device operation or atlas capacity cannot support the GPU path. Event listeners are attached to the stable grid host, not the replaceable drawing canvas.

### WebGPU instance ABI

Each instance is 16 floats / 64 bytes:

| Float offsets | Data | Meaning |
|---|---|---|
| 0–3 | `rect: vec4f` | CSS-pixel x, y, width, height |
| 4–7 | `uv: vec4f` | Atlas u, v, width, height; zero width indicates a solid quad |
| 8–11 | `color: vec4f` | Normalized RGBA |
| 12–15 | `clip: vec4f` | CSS-pixel left, top, right, bottom |

The vertex shader derives six vertices from `vertex_index` and maps CSS positions to NDC using a 16-byte viewport uniform. The fragment shader clips, reads the glyph mask with explicit level-zero sampling, and emits premultiplied color. The blend factors are `one` / `one-minus-src-alpha`. All instances use a single atlas texture and bind group, so ordered cell and glyph instances can share one render pass and `draw(6, instanceCount)` call.

The persistent instance array and GPU vertex buffer grow by powers of two. The atlas is a 2048-square RGBA texture. Individual glyph masks are rasterized on a native offscreen 2D canvas, cached by font descriptor and code point, and uploaded when dirty. This is real quad/glyph GPU rendering; it is not Canvas2D rendering of the final worksheet followed by presentation of that entire bitmap.

There is no multi-page atlas, glyph eviction, shaping engine, MSDF representation, subpixel glyph positioning guarantee or arbitrary-language typography certification. Atlas exhaustion requests Canvas fallback rather than entering a reset/rebuild loop. Future production work should introduce shaped glyph runs, a multi-page atlas with lifetime tracking, and text-layout equivalence tests between backends.

Device pixel ratio changes update backing dimensions. Zoom remains a separate logical transform. Device loss triggers fallback; this release does not automatically reacquire an adapter. GPU setup and draw execution are implemented but were not executed in the release's restricted test environment.

## 4. UI interaction and commands

Normal HTML controls implement the ribbon, panels, dialogs, file input and formula bar. A textarea overlays the active cell during editing, keeping caret behavior, selection and composition in the browser's text system. Keyboard dispatch distinguishes text-editing mode from worksheet navigation and command shortcuts. Address-box Enter explicitly stops propagation so it cannot also trigger grid Enter navigation.

The clipboard path supports native copy/cut/paste events. The toolbar attempts the asynchronous clipboard API but falls back to an in-app clipboard snapshot when permissions or API exposure do not allow it. Internal snapshots carry cell input and styles in addition to serialized TSV. Formula copy translates relative references. Cut moves data but does not implement a workbook-wide reference-retargeting transaction.

Chart definitions are saved in the model; presentation uses escaped SVG generated from calculated values. Charts are data-bound in the sense that workbook changes rebuild their data views. They are not GPU plot primitives. The simple chart implementations have limited data-point counts and are not intended as a general negative-value/stacked/logarithmic chart engine.

## 5. IO and trust boundaries

CSV parsing handles delimiters, quotes, embedded line breaks and BOMs. CSV import treats `=`-leading values as literals. CSV-value export neutralizes literal strings that resemble formula injection. Deliberate spreadsheet clipboard paste can enter formulas, which are evaluated only by the restricted interpreter.

XLSX export writes SpreadsheetML XML inside an uncompressed standards-style ZIP container. Import reads central-directory entries and supports stored and DEFLATE members through native `DecompressionStream`. It rejects encrypted archives and unsupported compression, checks CRCs, refuses DTD/entity declarations and skips path traversal entries. Compressed input is capped at 32 MiB; aggregate expanded data at 64 MiB, with a 32 MiB individual-member limit and 4,096-part cap. These limits bound common allocation hazards but do not substitute for a full hostile-document fuzzing campaign.

The XLSX path supports ordinary/shared formulas, inline/shared strings, basic styles and selected layout metadata. It deliberately omits document parts for comments/notes, chart drawings, pivot tables, macros, external connections and advanced conditional formatting. Export sets up a document that receiving spreadsheet software may recalculate, but that software was not used for the release validation. Stored and compressed round-trip checks validate Gridline's own supported subset, not full Office compliance.

No telemetry, networking, collaboration or authentication is implemented. Local storage is neither encrypted nor shared. Serving this app does not create a backend workbook database. Loading untrusted source code or granting third-party scripts same-origin access is outside the application security model.

## 6. Explicit limits and next production work

Logical worksheet capacity is not an allocation or throughput guarantee. The principal implementation boundaries are 200,000 cells per bounded rectangular operation/evaluated range, 256 nested cell evaluations, 128 parser nesting levels, 100 history commands, 32,767 characters per cell written through `setCell`, and a fixed-size glyph atlas. Print output is limited to 10,000 selected cells. Imported JSON permits up to 256 worksheets and 1,000,000 stored cell records per sheet; practical memory and autosave limits can be reached much earlier.

Priority production extensions are worker-based revisioned calculation; byte-budgeted history and incremental durable storage; complete structural-reference transforms; broad Excel differential tests; hostile-file fuzzing; language shaping and accessibility virtualization; paged glyph atlases; GPU validation/performance runs on real devices; and explicit chart and formatting compatibility policies. These are not hidden dependencies of this build: the delivered application runs without them, but it should not be represented as having their guarantees.
