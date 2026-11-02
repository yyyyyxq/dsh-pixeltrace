# dsh-pixeltrace — implementation contract (v1)

Single source of truth for everyone building this plugin. Read it fully before writing code.
Repo root: `D:\pixeltrace` (this file lives there). Lead owns every file except the ones named in
"Write scopes" below.

## 1. Product

A DeepSeek Harness (Cordis) plugin, panel in the Harness Web UI: the user drops a raster image
(screenshot, paper figure, diagram), the plugin rebuilds it on a PowerPoint slide out of
**individually editable objects** (solid-colour rectangles, cropped picture objects, optional
text boxes) whose rendered result is **pixel-identical to the source image**, and downloads a
`.pptx`. Multiple images produce multiple slides.

Three output modes: `smart` (default: flat areas → colour rectangles, detailed/text areas →
exactly cropped picture objects, one object per region, whole slide still pixel-exact),
`vector` (everything becomes colour rectangles; may need a correction layer and a shape budget),
`image` (one picture object per slide: guaranteed exact).

## 2. Hard rules

1. **Zero runtime dependencies.** No npm packages anywhere. Browser decode uses `createImageBitmap`,
   ZIP compression uses `CompressionStream`/`DecompressionStream` when present, else
   `await import('node:zlib')`. No `fs`/`Buffer` at the browser hot path.
2. **Plain script fragments, not ES modules.** Every file under `src/` is a *fragment*: a bare
   JavaScript body with **no `import`/`export` statements and no top-level `await`**. The Lead
   concatenates fragments into one client bundle inside a single closure
   (`window.__ModuleLoader__.load({ id, factory })`), and Node tests evaluate the fragments with
   `new Function(...)`. Therefore every fragment must run when simply concatenated: only
   function/const declarations at top level.
3. **Unique top-level names.** All top-level identifiers must start with the fragment's prefix:
   `scene*` for `src/scene.frag.js`, `pptx*` for `src/pptx.frag.js`, `ui*` for `src/ui.frag.js`,
   `core*` for `src/entry.frag.js`. Never shadow `React`, `ctx`, `require`.
4. **Determinism.** Same input → byte-identical scene (and, given a fixed PNG encoder, identical
   PPTX apart from nothing: ZIP timestamps are fixed constants).
5. **No DOM in `scene.frag.js`.** It must run in Node 22 unchanged. `pptx.frag.js` may touch
   `CompressionStream` / `node:zlib` only, never `document`.
6. **Integer pixel geometry.** Every shape coordinate/size is an integer in source-image pixels,
   origin top-left, y down, `x + w <= scene.width`, `y + h <= scene.height`.
7. **Colour format** is lowercase `#rrggbb` everywhere (`/^#[0-9a-f]{6}$/`).

## 3. Data model

```js
Scene = {
  width:  int,          // source image pixels
  height: int,
  background: '#rrggbb',// composites transparent source pixels
  shapes: Shape[],      // paint order: later shapes are drawn on top
  palette: string[],    // distinct colours actually used by rect shapes, by descending area
  stats: Stats
}

Shape =
  | { k:'r', x, y, w, h, c:'#rrggbb', role:'bg'|'flat'|'fix' }      // opaque filled rect, no outline
  | { k:'i', x, y, w, h, px:{ data:Uint8ClampedArray, width:int, height:int }, role:'photo'|'text'|'detail' }
  | { k:'t', x, y, w, h, s:string, sz:number, c:'#rrggbb', al:'l'|'ctr'|'r', b:boolean, f?:string, role:'text' }
```
* `k:'i'`: `px` is the exact RGBA crop of the source (w*h*4), row-major, straight (non-premultiplied)
  alpha. The crop must equal `source[y:y+h, x:x+w]` exactly.
* `k:'t'`: `sz` is the font size in **points** (`px * 0.75`), `c` is the fill colour used to erase the
  region, `s` may be `''` (empty editable text box, the normal case until the user types). `f`
  defaults to `'Arial'`. Text boxes must have integer geometry, zero insets, no wrap, anchored top.
* `Stats = { shapes, rects, images, texts, paletteSize, exactRatio, mismatchPixels, meanAbs, maxAbs,
  ms, notes:string[] }`.

## 4. `src/scene.frag.js` (owner: scene-dev)

```js
/** @param {{width:number,height:number,data:Uint8ClampedArray}} image  RGBA, row-major
 *  @param {SceneOptions} options
 *  @returns {{scale:number, scene:Scene}}  // scale = px→slide scale (1 unless over the 5376px cap) */
function sceneBuild(image, options)

/** Painter's-algorithm rasteriser of a scene back to RGBA (used for fidelity metrics + panel preview).
 *  @returns {{width:number,height:number,data:Uint8ClampedArray}} */
function sceneRasterize(scene)

/** @param {{width,height,data}} a  @param {{width,height,data}} b  @param {{ignoreRects?:Array<{x,y,w,h}>}} [opts]
 *  @returns {{pixels:number, mismatchPixels:number, exactRatio:number, meanAbs:number, maxAbs:number,
 *             exactRatioIgnoring:number|null, heatmap:Uint8ClampedArray}} // heatmap = RGBA, red on error */
function sceneDiff(a, b, opts)

const sceneDefaults = { /* see table */ }
```

### Options (all optional, defaults shown)

| key | default | meaning |
|---|---|---|
| `mode` | `'smart'` | `'smart'` \| `'vector'` \| `'image'` |
| `maxColors` | `0` | `0` = original colours; `16/32/64/128` = quantise (allowed to hurt fidelity) |
| `mergeTolerance` | `0` | max per-channel delta when merging rect colours; `0` = exact |
| `shapeBudget` | `8000` | max shapes per slide; `0` = unlimited |
| `detailToImage` | `true` | smart mode: high-detail areas become exactly cropped pictures |
| `detailMinAreaPx` | `24*24` | a detail area smaller than this is merged into neighbours / traced |
| `regionMinPx` | `2` | minimum rect side |
| `alphaBackground` | `'#ffffff'` | colour used to composite transparent source pixels |
| `textRegions` | `false` | detect text blocks and emit `k:'t'` boxes instead of pixels |
| `textMinHeightPx` | `8` | ignore text blocks shorter than this |
| `maxSlidePx` | `5376` | longest slide side in px (56 in cap); larger images are scaled |

### Required behaviour

* `mode:'image'` → exactly one `k:'i'` shape covering the whole (composited) image. `exactRatio === 1`.
* `mode:'vector'` with `mergeTolerance:0`, `maxColors:0`, `shapeBudget:0`, `textRegions:false` →
  `exactRatio === 1`, only `k:'r'` shapes. `vector` = exact maximal-rectangle decomposition of runs of
  identical colour (row runs merged vertically), then an optional correction layer.
* `mode:'smart'` → flat regions stay as rectangles, detail/text regions become `k:'i'` crops.
  With `shapeBudget:0` it must also be exactly reconstructable (`exactRatio === 1`):
  a region is either *entirely* covered by rects of its exact colour, or *entirely* replaced by one
  exact crop. Never both, never neither (no seams: adjacent shapes must not overlap).
* Budget pressure: if the shape count exceeds `shapeBudget`, degrade in this order —
  (a) grow the detail-to-image regions, (b) raise the effective merge tolerance in small steps,
  (c) whole image as one `k:'i'`. Always record what happened in `stats.notes` and let
  `stats.exactRatio` tell the truth. Never exceed the budget (a budget that cannot be met falls
  through to one picture).
* `textRegions:true` → detect text-ish blocks (dark/light ink on a locally uniform background,
  consistent stroke height, line-oriented), each block becomes an empty `k:'t'` box: erase the block
  with its dominant background colour (one rect covering the text ink box, or a rect per line),
  then place the text box with zero insets over the ink bbox, `sz = clamp(round(inkH*0.9*0.75), 6, 96)`.
  Text ink pixels are excluded from the fidelity denominator (`exactRatioIgnoring`, `ignoreRects`).
* Quality targets: `1600×900` realistic screenshot, defaults, **< 4 s** in Node on this machine,
  and fewer than `shapeBudget` shapes. `1920×1080` photo should fall back to a handful of crops
  instead of exploding.
* Never throw on odd input (1×1 image, fully transparent image, all-one-colour image); clamp sizes.

### Algorithm hints (follow unless you have something better)

1. Composite alpha over `alphaBackground`.
2. Optional palette quantisation (median cut or frequency-based) when `maxColors > 0`.
3. Exact maximal rectangles: per row, maximal runs of identical colour; merge a run into the rect
   above when `x/w/colour` match. This is the exact baseline (verifiable with `exactRatio`).
4. Smart split: tile the image (e.g. 8×8), compute per-tile detail (`distinct colours`, `run count`);
   flood/grow detail tiles into connected regions, take bboxes, crop them as `k:'i'`, and decompose
   only the remaining "flat" mask into rectangles (a rect must be fully flat — clip it at the mask).
5. Budget/merging: merge rects whose colours differ within tolerance (weighted-average colour,
   area-weighted), then a correction pass: pixels still differing by more than `mergeTolerance` get
   small exact rects, until the budget is reached.
6. Text detection (optional, own heuristic): ink = pixel far from the local median background,
   connected components with height in `[textMinHeightPx, height/3]`, group components into lines by
   baseline proximity, then into blocks; filter out components that are large filled areas.
7. Compute `stats` by `sceneRasterize` + `sceneDiff` before returning; `ms` from `Date.now()`
   (use a monotonic source, no `performance` in Node? `performance.now` exists in Node 22 too).

## 5. `src/pptx.frag.js` (owner: pptx-dev)

```js
/** PNG encoder (browser + Node). @returns {Promise<Uint8Array>} */
async function pptxEncodePng(rgba, width, height)

/** PNG decoder, used by Node tooling and tests only. @returns {Promise<{width,height,data}>} */
async function pptxDecodePng(bytes)

/** Build a complete .pptx package.
 *  @param {Array<{scene:Scene, name?:string, background?:string}>} slides
 *  @param {{title?:string, subject?:string, creator?:string, emuPerPx?:number, maxSlideEmu?:number,
 *           layout16x9?:boolean}} [options]
 *  @returns {Promise<Uint8Array>} */
async function pptxBuild(slides, options)

/** ZIP helpers (exported for tests): @returns {Promise<Map<string,Uint8Array>>} */
async function pptxUnzip(bytes)
async function pptxZip(files /* Array<{name, data:Uint8Array}> */)
```

### Package requirements

* Valid OOXML PresentationML that opens in PowerPoint, WPS, LibreOffice Impress and Keynote.
  Minimum parts: `[Content_Types].xml`, `_rels/.rels`, `docProps/core.xml`, `docProps/app.xml`,
  `ppt/presentation.xml` (+ rels), `ppt/slideMasters/slideMaster1.xml` (+ rels),
  `ppt/slideLayouts/slideLayout1.xml` (+ rels), `ppt/theme/theme1.xml`,
  `ppt/slides/slideN.xml` (+ rels), `ppt/media/imageN.png`.
  A blank slide layout (`type="blank"`, no placeholders) is enough: every object is placed absolutely.
* Slide size: `cx = round(scene.width * emuPerPx * scale)`, `cy = round(scene.height * emuPerPx * scale)`,
  `emuPerPx` default `9525` (96 dpi ⇒ 1 px = 9525 EMU, exact), `scale = min(1, maxSlideEmu / maxPxEdge / 9525)`
  with `maxSlideEmu = 51206400` (56 in). Set `<p:sldSz cx cy type="custom"/>`.
* Slide background from `scene.background` via `<p:bg><p:bgPr><a:solidFill>…`, plus a solid background
  rectangle behind everything is *not* needed (background alone is enough).
* Mapping: `k:'r'` → `<p:sp>` with `<a:prstGeom prst="rect">`, `<a:solidFill>`, `<a:ln><a:noFill/>`;
  `k:'i'` → `<p:pic>` with a media part + `<a:blipFill><a:stretch><a:fillRect/></a:stretch>`,
  `<p:spPr>` geometry exactly the shape box; `k:'t'` → `<p:sp>` with `<p:txBody>`
  (`<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>`,
  `<a:p><a:pPr algn="l|ctr|r"/><a:r><a:rPr lang="zh-CN" sz="pt*100" b="1|0" dirty="0">…`).
* Every shape gets a stable unique `id` (start at 2, increment by 1, never 0/1), a human `name`
  (`矩形 12`, `图片 3`, `文本 5` are fine — the panel renames them later if needed), and
  `<a:xfrm>` with `<a:off>`/`<a:ext>`. EMU must be `Math.round(px * emuPerPx * scale)`.
* Media parts are de-duplicated by byte equality, numbered `image1.png…` in slide order.
* ZIP: local headers + central directory, method 8 (`deflate-raw`) when compression is available,
  method 0 (store) otherwise; CRC-32 required; fixed DOS date/time `1980-01-01 00:00:00`
  (time 0, date 0x0021) so output is deterministic; no ZIP64 (files < 4 GB); UTF-8 flag off for
  ASCII names (all names are ASCII).
* `pptxBuild` must not mutate the scene.

### Acceptance for the writer

* `node --test test/pptx.test.mjs` passes, including: round-trip of `pptxZip`/`pptxUnzip`,
  PNG encode→decode equality, and **external validation**: the produced bytes load in
  `@office-kit/pptx` and pass `validatePresentation(pres)` with zero errors, imported by absolute
  path in Node:
  `await import('file:///C:/Users/yxq/.dsh/profiles/web/node_modules/@office-kit/pptx/dist/index.js')`
  (that package is installed in the harness profile; never add it as a dependency of this repo).
* A demo file written to `D:\pixeltrace\out\pptx-dev-demo.pptx` (2 slides: 300×200 with a few rects,
  one cropped picture and one text box) that the Lead renders with LibreOffice.

## 6. Panel (owner: Lead)

`src/ui.frag.js` (React, from `require('react')`), `src/entry.frag.js` (the loader payload),
`lib/client.js` (generated bundle, committed), `lib/index.js` (host entry, does nothing besides
declaring the plugin), `package.json`, `cordis.patch.yml`, `locale/{zh,en}.json`, `icon.svg`,
`tools/build.mjs` (concatenate fragments → `lib/client.js`), `tools/verify.mjs` (Node end-to-end),
`README.md`.

## 7. Testing rules

* Tests are `node:test` files `test/scene.test.mjs`, `test/pptx.test.mjs`. Load the fragments with
  `new Function`, do not duplicate the sources: `test/harness.mjs` (Lead) exposes
  `loadfrags()` returning the evaluated API. If it does not exist yet, create a local loader in your
  own test file; the Lead will reconcile.
* Run tests with `node --test test/` from `D:\pixeltrace`.
* Fixtures are generated procedurally inside tests (checkerboards, gradients, antialiased "text"
  strokes, photos via deterministic noise). Do not add binary fixture files.
* Every claim in your report must be backed by a command you actually ran; paste the command and the
  relevant output lines.

## 8. Write scopes (do not touch other files)

| owner | files |
|---|---|
| scene-dev | `src/scene.frag.js`, `test/scene.test.mjs`, `out/` (its own scratch) |
| pptx-dev | `src/pptx.frag.js`, `test/pptx.test.mjs`, `out/` (its own scratch) |
| Lead | everything else, plus `lib/client.js` generation and the browser round |

Report back when done: files written, test command + result, public functions with one-line
signatures, known limitations, and anything the Lead must do (e.g. a UI hook you needed).
