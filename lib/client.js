/* dsh-pixeltrace — 0.1.0
 *
 * GENERATED FILE — do not edit by hand. Source of truth: src/*.frag.js.
 * Rebuild with: node tools/build.mjs
 */
window.__ModuleLoader__.load({
  id: 'dsh-pixeltrace',
  factory(require) {
    const React = require('react');
    const { useState, useEffect, useLayoutEffect, useRef, useMemo, useCallback } = React;
    const h = React.createElement;

/* ==================== src/entry.frag.js ==================== */
/* PixelTrace loader payload: option normalisation, scene building and deck assembly.
 *
 * Plain fragment: no import/export statements, no top-level await. Every top-level
 * name starts with `core`. This file is concatenated into `lib/client.js` by
 * tools/build.mjs and evaluated directly by Node tooling/tests.
 */

const coreVersion = '0.1.0';
const coreModes = ['smart', 'vector', 'image'];
const coreMaxColorChoices = [0, 8, 16, 32, 64, 128, 256];

/** Defaults for the scene engine (mirrors CONTRACT.md §4) plus a few deck defaults. */
const coreDefaults = {
  mode: 'smart',
  maxColors: 0,
  mergeTolerance: 0,
  shapeBudget: 8000,
  detailToImage: true,
  textRegions: false,
  alphaBackground: '#ffffff',
  maxSlidePx: 5376,
};

const coreDeckDefaults = {
  title: 'PixelTrace',
  creator: 'PixelTrace (dsh-pixeltrace)',
};

function coreIsHexColor(value) {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value.trim());
}

function coreClampInt(value, low, high, fallback) {
  const n = Math.round(Number(value));
  if (!Number.isFinite(n)) {
    return fallback;
  }
  return Math.min(high, Math.max(low, n));
}

/**
 * Coerce arbitrary UI state into options the scene engine accepts.
 * Unknown or malformed values fall back to the documented defaults.
 */
function coreNormalizeOptions(raw) {
  const src = raw && typeof raw === 'object' ? raw : {};
  const mode = coreModes.indexOf(src.mode) >= 0 ? src.mode : coreDefaults.mode;
  const rawColors = coreClampInt(src.maxColors, 0, 256, coreDefaults.maxColors);
  const maxColors = coreMaxColorChoices.indexOf(rawColors) >= 0 ? rawColors : coreDefaults.maxColors;
  return {
    mode,
    maxColors,
    mergeTolerance: coreClampInt(src.mergeTolerance, 0, 255, coreDefaults.mergeTolerance),
    shapeBudget: coreClampInt(src.shapeBudget, 0, 200000, coreDefaults.shapeBudget),
    detailToImage: src.detailToImage !== false,
    textRegions: src.textRegions === true,
    alphaBackground: coreIsHexColor(src.alphaBackground)
      ? src.alphaBackground.trim().toLowerCase()
      : coreDefaults.alphaBackground,
    maxSlidePx: coreClampInt(src.maxSlidePx, 256, 8192, coreDefaults.maxSlidePx),
  };
}

/** Build one scene from an RGBA image. Returns the engine result unchanged. */
function coreScene(image, options) {
  const opts = coreNormalizeOptions(options);
  const built = sceneBuild(image, opts);
  return { scale: built.scale, scene: built.scene, options: opts };
}

/** Aggregate per-slide numbers for the panel summary. */
function coreSummary(items) {
  const acc = {
    slides: 0,
    shapes: 0,
    rects: 0,
    images: 0,
    texts: 0,
    palette: 0,
    exactRatio: 1,
    meanAbs: 0,
    maxAbs: 0,
    ms: 0,
    notes: [],
  };
  for (const item of items || []) {
    const scene = item && item.scene;
    if (!scene || !scene.stats) {
      continue;
    }
    const stats = scene.stats;
    acc.slides += 1;
    acc.shapes += stats.shapes || 0;
    acc.rects += stats.rects || 0;
    acc.images += stats.images || 0;
    acc.texts += stats.texts || 0;
    acc.palette = Math.max(acc.palette, stats.paletteSize || 0);
    acc.exactRatio = Math.min(acc.exactRatio, stats.exactRatio == null ? 1 : stats.exactRatio);
    acc.meanAbs = Math.max(acc.meanAbs, stats.meanAbs || 0);
    acc.maxAbs = Math.max(acc.maxAbs, stats.maxAbs || 0);
    acc.ms += stats.ms || 0;
    for (const note of stats.notes || []) {
      if (acc.notes.indexOf(note) < 0) {
        acc.notes.push(note);
      }
    }
  }
  return acc;
}

/** Filesystem/URL friendly base name, no extension. */
function coreSafeName(name, fallback) {
  const base = String(name == null ? '' : name)
    .replace(/\.[a-z0-9]{1,8}$/i, '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
  return base || fallback || 'pixeltrace';
}

/** Deck file name; single image keeps its own name, several get an index. */
function coreDeckFileName(name, total) {
  const base = coreSafeName(name || coreDeckDefaults.title, 'pixeltrace');
  if ((total || 1) > 1) {
    return base + '-' + String(total) + 'p.pptx';
  }
  return base + '.pptx';
}

/**
 * Assemble the .pptx bytes for a list of processed items.
 * @param {Array<{name?:string, scene:object}>} items
 * @param {object} [deckOptions]
 * @returns {Promise<{bytes: Uint8Array, slides: number}>}
 */
async function coreBuildDeck(items, deckOptions) {
  const list = (items || []).filter((item) => item && item.scene);
  if (!list.length) {
    throw new Error('no processed image to export');
  }
  const opts = deckOptions && typeof deckOptions === 'object' ? deckOptions : {};
  const slides = list.map((item, index) => ({
    scene: item.scene,
    name: (item.name || 'slide-' + (index + 1)).slice(0, 120),
    background: item.scene.background,
  }));
  const bytes = await pptxBuild(slides, {
    title: opts.title || coreDeckDefaults.title,
    creator: opts.creator || coreDeckDefaults.creator,
    subject: opts.subject || '',
  });
  return { bytes, slides: slides.length };
}

/** Trigger a browser download; in Node (tools, tests) just return the bytes. */
async function coreSaveBytes(bytes, fileName, mimeType) {
  const type = mimeType || 'application/octet-stream';
  if (typeof document === 'undefined' || typeof Blob === 'undefined') {
    return { saved: false, bytes, fileName };
  }
  const blob = new Blob([bytes], { type });
  const picker = typeof window !== 'undefined' ? window.showSaveFilePicker : null;
  if (typeof picker === 'function') {
    try {
      const handle = await picker({
        suggestedName: fileName,
        types: [
          {
            description: 'Microsoft PowerPoint',
            accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return { saved: true, fileName };
    } catch (error) {
      if (error && (error.name === 'AbortError' || error.name === 'NotAllowedError')) {
        return { saved: false, cancelled: true, fileName };
      }
      // fall through to the anchor strategy
    }
  }
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = 'noopener';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  setTimeout(() => {
    try {
      anchor.remove();
      URL.revokeObjectURL(url);
    } catch {
      /* ignore */
    }
  }, 4000);
  return { saved: true, fileName };
}

/* ==================== src/scene.frag.js ==================== */
/* dsh-pixeltrace — scene engine fragment.
 * Plain script fragment: no import/export, no top-level await, no DOM/browser API.
 * Every top-level identifier starts with `scene`.
 *
 * Public API (see CONTRACT.md §4):
 *   sceneBuild(image, options) -> {scale, scene}
 *   sceneRasterize(scene)      -> {width, height, data}
 *   sceneDiff(a, b, opts)      -> {pixels, mismatchPixels, exactRatio, meanAbs, maxAbs, exactRatioIgnoring, heatmap}
 *   sceneDefaults              -> option defaults
 */

/** Default options (CONTRACT §4). A new object is handed out by sceneBuild via the parser. */
const sceneDefaults = {
  mode: 'smart',
  maxColors: 0,
  mergeTolerance: 0,
  shapeBudget: 8000,
  detailToImage: true,
  detailMinAreaPx: 24 * 24,
  regionMinPx: 2,
  alphaBackground: '#ffffff',
  textRegions: false,
  textMinHeightPx: 8,
  maxSlidePx: 5376,
};

/* ------------------------------------------------------------------ *
 * internal tuning constants (not part of the public options table)    *
 * ------------------------------------------------------------------ */

/** smart-mode analysis tile edge, in source pixels. */
const sceneTileSize = 8;
/** an 8x8 tile whose neighbour pairs differ this often is "detail". */
const sceneDetailTransitions = 12;
/** an 8x8 tile with at least this many distinct colours is "detail". */
const sceneDetailDistinct = 8;
/** memory guard: exact tiling never materialises more rects than this. */
const sceneRectHardCap = 1200000;
/** memory guard: sources larger than this are nearest-neighbour downscaled. */
const sceneMaxSourcePixels = 24000000;
/** luminance distance from the local background that counts as text ink. */
const sceneInkThreshold = 40;

/* ------------------------------------------------------------------ *
 * small utilities                                                     *
 * ------------------------------------------------------------------ */

function sceneNow() {
  if (typeof performance !== 'undefined' && performance && typeof performance.now === 'function') {
    return performance.now();
  }
  return Date.now();
}

function sceneNum(value, lo, hi, fallback) {
  let n = NaN;
  if (typeof value === 'number') n = value;
  else if (typeof value === 'string' && value.trim() !== '') n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  if (n < lo) return lo;
  if (n > hi) return hi;
  return n;
}

function sceneInt(value, lo, hi, fallback) {
  return Math.round(sceneNum(value, lo, hi, fallback));
}

function sceneBool(value, fallback) {
  return typeof value === 'boolean' ? value : fallback;
}

function sceneHexPad(v) {
  return v < 16 ? '0' + v.toString(16) : v.toString(16);
}

/** packed 0xrrggbb -> lowercase '#rrggbb'. */
function sceneHexFromPack(p) {
  return '#' + sceneHexPad((p >>> 16) & 255) + sceneHexPad((p >>> 8) & 255) + sceneHexPad(p & 255);
}

function scenePack(r, g, b) {
  return ((r & 255) << 16) | ((g & 255) << 8) | (b & 255);
}

/** '#rrggbb' (or 'rrggbb') -> [r,g,b]; invalid input yields the fallback colour. */
function sceneHexToRgb(value, fallback) {
  let hex = typeof value === 'string' ? value.trim().toLowerCase() : '';
  if (hex[0] !== '#' && /^[0-9a-f]{6}$/.test(hex)) hex = '#' + hex;
  if (/^#[0-9a-f]{3}$/.test(hex)) {
    hex = '#' + hex[1] + hex[1] + hex[2] + hex[2] + hex[3] + hex[3];
  }
  if (!/^#[0-9a-f]{6}$/.test(hex)) hex = fallback;
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
}

/** '#rrggbb' with fallback; always returns lowercase '#rrggbb'. */
function sceneHexOpt(value, fallback) {
  let hex = typeof value === 'string' ? value.trim().toLowerCase() : '';
  if (hex[0] !== '#' && /^[0-9a-f]{6}$/.test(hex)) hex = '#' + hex;
  if (/^#[0-9a-f]{3}$/.test(hex)) {
    hex = '#' + hex[1] + hex[1] + hex[2] + hex[2] + hex[3] + hex[3];
  }
  if (!/^#[0-9a-f]{6}$/.test(hex)) hex = fallback;
  return hex;
}

/** Normalised option bag; every key present, every value clamped to its legal range. */
function sceneParseOptions(options) {
  const o = options && typeof options === 'object' ? options : {};
  const d = sceneDefaults;
  const mode = o.mode === 'vector' || o.mode === 'image' || o.mode === 'smart' ? o.mode : d.mode;
  return {
    mode,
    maxColors: sceneInt(o.maxColors, 0, 256, d.maxColors),
    mergeTolerance: sceneInt(o.mergeTolerance, 0, 255, d.mergeTolerance),
    shapeBudget: sceneInt(o.shapeBudget, 0, 100000000, d.shapeBudget),
    detailToImage: sceneBool(o.detailToImage, d.detailToImage),
    detailMinAreaPx: sceneInt(o.detailMinAreaPx, 0, 100000000, d.detailMinAreaPx),
    regionMinPx: sceneInt(o.regionMinPx, 1, 4096, d.regionMinPx),
    alphaBackground: sceneHexOpt(o.alphaBackground, d.alphaBackground),
    textRegions: sceneBool(o.textRegions, d.textRegions),
    textMinHeightPx: sceneInt(o.textMinHeightPx, 1, 4096, d.textMinHeightPx),
    maxSlidePx: sceneInt(o.maxSlidePx, 16, 1000000, d.maxSlidePx),
  };
}

/* ------------------------------------------------------------------ *
 * input normalisation + alpha compositing                             *
 * ------------------------------------------------------------------ */

/** Never throws. Odd/short/missing pixel data becomes transparent black. */
function sceneNormalizeImage(image) {
  const img = image && typeof image === 'object' ? image : {};
  let width = sceneInt(img.width, 0, 200000, 0);
  let height = sceneInt(img.height, 0, 200000, 0);
  if (width < 1) width = 1;
  if (height < 1) height = 1;
  const need = width * height * 4;
  const data = img.data;
  let src = null;
  if (data && typeof data === 'object' && typeof data.length === 'number' && data.length >= need) src = data;

  if (width * height > sceneMaxSourcePixels) {
    const factor = Math.sqrt((width * height) / sceneMaxSourcePixels);
    const nw = Math.max(1, Math.floor(width / factor));
    const nh = Math.max(1, Math.floor(height / factor));
    const out = new Uint8ClampedArray(nw * nh * 4);
    for (let y = 0; y < nh; y++) {
      const sy = Math.min(height - 1, Math.floor((y * height) / nh));
      for (let x = 0; x < nw; x++) {
        if (!src) break;
        const sx = Math.min(width - 1, Math.floor((x * width) / nw));
        const s = (sy * width + sx) * 4;
        const t = (y * nw + x) * 4;
        out[t] = src[s];
        out[t + 1] = src[s + 1];
        out[t + 2] = src[s + 2];
        out[t + 3] = src[s + 3];
      }
    }
    return { width: nw, height: nh, data: out, shrunk: true };
  }

  if (src) return { width, height, data: src, shrunk: false };

  const blank = new Uint8ClampedArray(need);
  if (data && typeof data === 'object' && typeof data.length === 'number' && data.length > 0) {
    const m = Math.min(need, data.length);
    for (let i = 0; i < m; i++) blank[i] = data[i];
  }
  return { width, height, data: blank, shrunk: false };
}

/** Composite over `bgHex`; returns straight-alpha (a=255) RGBA plus a packed colour plane. */
function sceneCompose(image, bgHex) {
  const width = image.width;
  const height = image.height;
  const src = image.data;
  const bg = sceneHexToRgb(bgHex, '#ffffff');
  const br = bg[0];
  const bgg = bg[1];
  const bb = bg[2];
  const n = width * height;
  const rgba = new Uint8ClampedArray(n * 4);
  const pack = new Uint32Array(n);
  let transparent = 0;
  let semi = 0;
  for (let i = 0, p = 0; i < n; i++, p += 4) {
    const a = src[p + 3];
    let r = src[p];
    let g = src[p + 1];
    let b = src[p + 2];
    if (a !== 255) {
      if (a === 0) {
        r = br;
        g = bgg;
        b = bb;
        transparent++;
      } else {
        const inv = 255 - a;
        r = ((r * a + br * inv + 127) / 255) | 0;
        g = ((g * a + bgg * inv + 127) / 255) | 0;
        b = ((b * a + bb * inv + 127) / 255) | 0;
        semi++;
      }
    }
    rgba[p] = r;
    rgba[p + 1] = g;
    rgba[p + 2] = b;
    rgba[p + 3] = 255;
    pack[i] = (r << 16) | (g << 8) | b;
  }
  return { width, height, rgba, pack, transparent, semi };
}

/* ------------------------------------------------------------------ *
 * palette quantisation (maxColors)                                    *
 * ------------------------------------------------------------------ */

function sceneColorDistance(p, q) {
  const dr = ((p >>> 16) & 255) - ((q >>> 16) & 255);
  const dg = ((p >>> 8) & 255) - ((q >>> 8) & 255);
  const db = (p & 255) - (q & 255);
  return dr * dr + dg * dg + db * db;
}

/**
 * Median-cut quantisation to at most `maxColors` colours.
 * If the image already uses no more colours than that, the input is returned untouched (exact).
 * `exclude` (optional Uint8Array, 1 = pixel will become a crop/text region): those pixels do not
 * take part in the palette ranking, because quantisation may only change flat-rect colours.
 *
 * The histogram is split into `maxColors` boxes, always subdividing the box with the largest pixel
 * weight along its longest colour axis at the weighted median; each box then contributes the
 * area-weighted mean colour of its members. This keeps the palette spread over the actual gamut,
 * unlike a pure frequency ranking (which collapses to near-identical colours when every colour is
 * used equally often).
 */
function sceneQuantizePacked(pack, maxColors, exclude) {
  const n = pack.length;
  const counts = new Map();
  for (let i = 0; i < n; i++) {
    if (exclude && exclude[i]) continue;
    const p = pack[i];
    counts.set(p, (counts.get(p) || 0) + 1);
  }
  if (counts.size === 0) return { pack, changed: false, colors: 0 };
  if (counts.size <= maxColors) return { pack, changed: false, colors: counts.size };
  const target = Math.max(1, maxColors);
  const total = counts.size;
  const keys = new Int32Array(total);
  const weight = new Int32Array(total);
  let k = 0;
  counts.forEach((c, p) => {
    keys[k] = p;
    weight[k] = c;
    k++;
  });
  const cr = new Uint8Array(total);
  const cg = new Uint8Array(total);
  const cb = new Uint8Array(total);
  let sum = 0;
  for (let i = 0; i < total; i++) {
    const p = keys[i];
    cr[i] = (p >>> 16) & 255;
    cg[i] = (p >>> 8) & 255;
    cb[i] = p & 255;
    sum += weight[i];
  }
  const order = new Int32Array(total);
  for (let i = 0; i < total; i++) order[i] = i;
  const boxes = [{ s: 0, e: total, w: sum }];
  while (boxes.length < target) {
    let pick = -1;
    let bestW = -1;
    for (let i = 0; i < boxes.length; i++) {
      const b = boxes[i];
      if (b.e - b.s < 2) continue;
      if (b.w > bestW) {
        bestW = b.w;
        pick = i;
      }
    }
    if (pick < 0) break;
    const box = boxes[pick];
    let rmin = 255;
    let rmax = 0;
    let gmin = 255;
    let gmax = 0;
    let bmin = 255;
    let bmax = 0;
    for (let i = box.s; i < box.e; i++) {
      const q = order[i];
      const r = cr[q];
      const g = cg[q];
      const b = cb[q];
      if (r < rmin) rmin = r;
      if (r > rmax) rmax = r;
      if (g < gmin) gmin = g;
      if (g > gmax) gmax = g;
      if (b < bmin) bmin = b;
      if (b > bmax) bmax = b;
    }
    const rspan = rmax - rmin;
    const gspan = gmax - gmin;
    const bspan = bmax - bmin;
    let ch = cr;
    if (gspan > rspan && gspan >= bspan) ch = cg;
    else if (bspan > rspan && bspan > gspan) ch = cb;
    order.subarray(box.s, box.e).sort((p, q) => ch[p] - ch[q] || keys[p] - keys[q]);
    let acc = 0;
    let split = box.s + 1;
    for (let i = box.s; i < box.e - 1; i++) {
      acc += weight[order[i]];
      if (acc * 2 >= box.w) {
        split = i + 1;
        break;
      }
    }
    let leftW = 0;
    for (let i = box.s; i < split; i++) leftW += weight[order[i]];
    boxes[pick] = { s: box.s, e: split, w: leftW };
    boxes.splice(pick + 1, 0, { s: split, e: box.e, w: box.w - leftW });
  }
  const reps = [];
  for (let i = 0; i < boxes.length; i++) {
    const box = boxes[i];
    let sr = 0;
    let sg = 0;
    let sb = 0;
    let sw = 0;
    for (let j = box.s; j < box.e; j++) {
      const q = order[j];
      const w2 = weight[q];
      sr += cr[q] * w2;
      sg += cg[q] * w2;
      sb += cb[q] * w2;
      sw += w2;
    }
    const r = Math.min(255, Math.max(0, Math.round(sr / sw)));
    const g = Math.min(255, Math.max(0, Math.round(sg / sw)));
    const b = Math.min(255, Math.max(0, Math.round(sb / sw)));
    reps.push((r << 16) | (g << 8) | b);
  }
  const remap = new Map();
  const out = new Uint32Array(n);
  for (let i = 0; i < n; i++) {
    const p = pack[i];
    let q = remap.get(p);
    if (q === undefined) {
      let best = reps[0];
      let bestD = sceneColorDistance(p, best);
      for (let j = 1; j < reps.length; j++) {
        const d = sceneColorDistance(p, reps[j]);
        if (d < bestD) {
          bestD = d;
          best = reps[j];
        }
      }
      q = best;
      remap.set(p, q);
    }
    out[i] = q;
  }
  return { pack: out, changed: true, colors: reps.length };
}

/* ------------------------------------------------------------------ *
 * tiling: exact maximal-rectangle decomposition of colour row runs     *
 * ------------------------------------------------------------------ */

/**
 * Row runs of identical colour, merged downwards into maximal rectangles.
 * `skip` (optional Uint8Array, 1 = do not tile) lets flat-mask tiling cut runs.
 * Returns typed columns; when `cap` is exceeded the tiling aborts (caller degrades).
 */
function sceneTiling(pack, w, h, skip, cap) {
  let capacity = 1024;
  let X = new Int32Array(capacity);
  let Y = new Int32Array(capacity);
  let Wd = new Int32Array(capacity);
  let Hh = new Int32Array(capacity);
  let Pp = new Uint32Array(capacity);
  let n = 0;
  let actA = new Int32Array(w).fill(-1);
  let actB = new Int32Array(w).fill(-1);
  const limit = typeof cap === 'number' && cap > 0 ? cap : sceneRectHardCap;

  for (let y = 0; y < h; y++) {
    actB.fill(-1);
    const row = y * w;
    let x = 0;
    while (x < w) {
      if (skip && skip[row + x]) {
        x++;
        continue;
      }
      const p0 = pack[row + x];
      let x2 = x + 1;
      while (x2 < w) {
        const idx = row + x2;
        if (skip && skip[idx]) break;
        if (pack[idx] !== p0) break;
        x2++;
      }
      const rw = x2 - x;
      const ai = actA[x];
      if (ai >= 0 && Wd[ai] === rw && Pp[ai] === p0) {
        Hh[ai]++;
        actB[x] = ai;
      } else {
        if (n >= limit) return { n: limit, aborted: true, X, Y, Wd, Hh, Pp };
        if (n === capacity) {
          capacity *= 2;
          const nx = new Int32Array(capacity);
          nx.set(X);
          X = nx;
          const ny = new Int32Array(capacity);
          ny.set(Y);
          Y = ny;
          const nw2 = new Int32Array(capacity);
          nw2.set(Wd);
          Wd = nw2;
          const nh = new Int32Array(capacity);
          nh.set(Hh);
          Hh = nh;
          const np = new Uint32Array(capacity);
          np.set(Pp);
          Pp = np;
        }
        X[n] = x;
        Y[n] = y;
        Wd[n] = rw;
        Hh[n] = 1;
        Pp[n] = p0;
        actB[x] = n;
        n++;
      }
      x = x2;
    }
    const swap = actA;
    actA = actB;
    actB = swap;
  }
  return { n, aborted: false, X, Y, Wd, Hh, Pp };
}

/** Area-weighted colour blend of two packed colours. */
function sceneBlendPack(p1, a1, p2, a2) {
  const total = a1 + a2;
  if (total <= 0) return p1;
  const r = Math.round((((p1 >>> 16) & 255) * a1 + ((p2 >>> 16) & 255) * a2) / total);
  const g = Math.round((((p1 >>> 8) & 255) * a1 + ((p2 >>> 8) & 255) * a2) / total);
  const b = Math.round(((p1 & 255) * a1 + (p2 & 255) * a2) / total);
  return (r << 16) | (g << 8) | b;
}

/**
 * Merge adjacent rectangles whose colours differ by <= `tol` per channel
 * (contract: "max per-channel delta when merging rect colours").
 * A merge is only accepted when the blended colour stays within `tol` of every original colour the
 * rectangle now stands for, so the reported tolerance really is the per-pixel error bound.
 * Mutates the tiling in place, compacts, updates `t.n`; returns the new count.
 */
function sceneMergeSimilar(t, tol) {
  let n = t.n;
  if (!(tol > 0) || n <= 1) return n;
  const X = t.X;
  const Y = t.Y;
  const Wd = t.Wd;
  const Hh = t.Hh;
  const Pp = t.Pp;
  const dev = new Int32Array(n);
  const within = (p, q) => {
    const dr = ((p >>> 16) & 255) - ((q >>> 16) & 255);
    if (dr > tol || dr < -tol) return false;
    const dg = ((p >>> 8) & 255) - ((q >>> 8) & 255);
    if (dg > tol || dg < -tol) return false;
    const db = (p & 255) - (q & 255);
    return db <= tol && db >= -tol;
  };
  const chanDiff = (p, q) => {
    const dr = Math.abs(((p >>> 16) & 255) - ((q >>> 16) & 255));
    const dg = Math.abs(((p >>> 8) & 255) - ((q >>> 8) & 255));
    const db = Math.abs((p & 255) - (q & 255));
    const m = dr > dg ? dr : dg;
    return m > db ? m : db;
  };
  for (let pass = 0; pass < 6; pass++) {
    let merged = 0;
    // vertical neighbours: same x and w, next row starts right below.
    const map = new Map();
    for (let i = 0; i < n; i++) {
      if (Hh[i] <= 0) continue;
      const key = X[i] * 2097152 + Wd[i];
      const j = map.get(key);
      let done = false;
      if (j !== undefined && Hh[j] > 0 && Y[j] + Hh[j] === Y[i] && within(Pp[j], Pp[i])) {
        const areaJ = Wd[j] * Hh[j];
        const areaI = Wd[i] * Hh[i];
        const blended = sceneBlendPack(Pp[j], areaJ, Pp[i], areaI);
        const nd = Math.max(dev[j] + chanDiff(blended, Pp[j]), dev[i] + chanDiff(blended, Pp[i]));
        if (nd <= tol) {
          Pp[j] = blended;
          dev[j] = nd;
          Hh[j] += Hh[i];
          Hh[i] = 0;
          merged++;
          done = true;
        }
      }
      if (!done) map.set(key, i);
    }
    // horizontal neighbours: index order is (y asc, x asc).
    let prev = -1;
    for (let i = 0; i < n; i++) {
      if (Hh[i] <= 0) continue;
      let done = false;
      if (
        prev >= 0 &&
        Y[prev] === Y[i] &&
        Hh[prev] === Hh[i] &&
        X[prev] + Wd[prev] === X[i] &&
        within(Pp[prev], Pp[i])
      ) {
        const areaP = Wd[prev] * Hh[prev];
        const areaI = Wd[i] * Hh[i];
        const blended = sceneBlendPack(Pp[prev], areaP, Pp[i], areaI);
        const nd = Math.max(dev[prev] + chanDiff(blended, Pp[prev]), dev[i] + chanDiff(blended, Pp[i]));
        if (nd <= tol) {
          Pp[prev] = blended;
          dev[prev] = nd;
          Wd[prev] += Wd[i];
          Hh[i] = 0;
          merged++;
          done = true;
        }
      }
      if (!done) prev = i;
    }
    if (!merged) break;
    let m = 0;
    for (let i = 0; i < n; i++) {
      if (Hh[i] <= 0) continue;
      if (m !== i) {
        X[m] = X[i];
        Y[m] = Y[i];
        Wd[m] = Wd[i];
        Hh[m] = Hh[i];
        Pp[m] = Pp[i];
        dev[m] = dev[i];
      }
      m++;
    }
    n = m;
    if (n <= 1) break;
  }
  t.n = n;
  return n;
}

/* ------------------------------------------------------------------ *
 * smart-mode analysis: detail tiles, regions, disjoint crop boxes      *
 * ------------------------------------------------------------------ */

/**
 * Per-tile structure metrics. A tile is "detail" when it has too many colour
 * transitions or too many distinct colours (text, antialiasing, gradients, photos).
 */
function sceneTileMetrics(pack, w, h, tile) {
  const tw = Math.ceil(w / tile);
  const th = Math.ceil(h / tile);
  const detail = new Uint8Array(tw * th);
  const seen = new Int32Array(32);
  for (let ty = 0; ty < th; ty++) {
    const y0 = ty * tile;
    const y1 = Math.min(h, y0 + tile);
    for (let tx = 0; tx < tw; tx++) {
      const x0 = tx * tile;
      const x1 = Math.min(w, x0 + tile);
      let distinct = 0;
      let trans = 0;
      let isDetail = 0;
      for (let y = y0; y < y1 && !isDetail; y++) {
        const row = y * w;
        const downRow = y + 1 < y1 ? row + w : -1;
        for (let x = x0; x < x1; x++) {
          const p = pack[row + x];
          let known = false;
          for (let k = 0; k < distinct; k++) {
            if (seen[k] === p) {
              known = true;
              break;
            }
          }
          if (!known) {
            if (distinct < 32) seen[distinct++] = p;
          }
          if (x + 1 < x1 && pack[row + x + 1] !== p) trans++;
          if (downRow >= 0 && pack[downRow + x] !== p) trans++;
          if (trans >= sceneDetailTransitions || distinct >= sceneDetailDistinct) {
            isDetail = 1;
            break;
          }
        }
      }
      if (isDetail) detail[ty * tw + tx] = 1;
    }
  }
  return { tw, th, detail };
}

/** Chessboard-distance `k` dilation of a tile mask (prefix-sum windows, O(tiles)). */
function sceneDilateTiles(mask, tw, th, k) {
  if (!(k > 0)) return mask;
  const stride = tw + 1;
  const S = new Int32Array(stride * (th + 1));
  for (let y = 1; y <= th; y++) {
    let rowSum = 0;
    for (let x = 1; x <= tw; x++) {
      rowSum += mask[(y - 1) * tw + (x - 1)];
      S[y * stride + x] = S[(y - 1) * stride + x] + rowSum;
    }
  }
  const out = new Uint8Array(tw * th);
  for (let y = 0; y < th; y++) {
    const y0 = Math.max(0, y - k);
    const y1 = Math.min(th - 1, y + k);
    for (let x = 0; x < tw; x++) {
      const x0 = Math.max(0, x - k);
      const x1 = Math.min(tw - 1, x + k);
      const sum =
        S[(y1 + 1) * stride + (x1 + 1)] - S[y0 * stride + (x1 + 1)] - S[(y1 + 1) * stride + x0] + S[y0 * stride + x0];
      if (sum > 0) out[y * tw + x] = 1;
    }
  }
  return out;
}

/** 8-connected components of a tile mask -> bboxes in tile units. */
function sceneTileRegions(mask, tw, th) {
  const boxes = [];
  const seen = new Uint8Array(tw * th);
  const stack = new Int32Array(tw * th);
  const total = tw * th;
  for (let i = 0; i < total; i++) {
    if (!mask[i] || seen[i]) continue;
    let sp = 0;
    stack[sp++] = i;
    seen[i] = 1;
    let minX = tw;
    let minY = th;
    let maxX = -1;
    let maxY = -1;
    while (sp > 0) {
      const cur = stack[--sp];
      const cy = (cur / tw) | 0;
      const cx = cur - cy * tw;
      if (cx < minX) minX = cx;
      if (cx > maxX) maxX = cx;
      if (cy < minY) minY = cy;
      if (cy > maxY) maxY = cy;
      const ya = cy > 0 ? cy - 1 : 0;
      const yb = cy < th - 1 ? cy + 1 : th - 1;
      const xa = cx > 0 ? cx - 1 : 0;
      const xb = cx < tw - 1 ? cx + 1 : tw - 1;
      for (let yy = ya; yy <= yb; yy++) {
        const ro = yy * tw;
        for (let xx = xa; xx <= xb; xx++) {
          const ni = ro + xx;
          if (mask[ni] && !seen[ni]) {
            seen[ni] = 1;
            stack[sp++] = ni;
          }
        }
      }
    }
    boxes.push({ minX, minY, maxX, maxY });
  }
  return boxes;
}

/** Tile bboxes -> clamped pixel boxes. */
function sceneBoxesFromTiles(boxes, tile, w, h) {
  const out = [];
  for (let i = 0; i < boxes.length; i++) {
    const b = boxes[i];
    const x = b.minX * tile;
    const y = b.minY * tile;
    const x1 = Math.min(w, (b.maxX + 1) * tile);
    const y1 = Math.min(h, (b.maxY + 1) * tile);
    if (x1 > x && y1 > y) out.push({ x, y, w: x1 - x, h: y1 - y });
  }
  return out;
}

/**
 * Merge overlapping boxes until the list is pairwise disjoint.
 * Deterministic; used so crop/text regions never overlap (no seams, no double paint).
 */
function sceneDisjointBoxes(boxes) {
  const accepted = [];
  const CELL = 64;
  const grid = new Map();
  const gkey = (cx, cy) => cx * 1048576 + cy;
  const addIndex = (box, idx) => {
    const cx0 = Math.floor(box.x / CELL);
    const cx1 = Math.floor((box.x + box.w - 1) / CELL);
    const cy0 = Math.floor(box.y / CELL);
    const cy1 = Math.floor((box.y + box.h - 1) / CELL);
    for (let cy = cy0; cy <= cy1; cy++) {
      for (let cx = cx0; cx <= cx1; cx++) {
        const k = gkey(cx, cy);
        let arr = grid.get(k);
        if (!arr) {
          arr = [];
          grid.set(k, arr);
        }
        arr.push(idx);
      }
    }
  };
  for (let i = 0; i < boxes.length; i++) {
    const box = { x: boxes[i].x, y: boxes[i].y, w: boxes[i].w, h: boxes[i].h };
    let absorbed = true;
    while (absorbed) {
      absorbed = false;
      const cx0 = Math.floor(box.x / CELL);
      const cx1 = Math.floor((box.x + box.w - 1) / CELL);
      const cy0 = Math.floor(box.y / CELL);
      const cy1 = Math.floor((box.y + box.h - 1) / CELL);
      find: for (let cy = cy0; cy <= cy1; cy++) {
        for (let cx = cx0; cx <= cx1; cx++) {
          const arr = grid.get(gkey(cx, cy));
          if (!arr) continue;
          for (let t = 0; t < arr.length; t++) {
            const o = accepted[arr[t]];
            if (!o || o.dead) continue;
            if (o.x < box.x + box.w && box.x < o.x + o.w && o.y < box.y + box.h && box.y < o.y + o.h) {
              const nx = Math.min(box.x, o.x);
              const ny = Math.min(box.y, o.y);
              const nx1 = Math.max(box.x + box.w, o.x + o.w);
              const ny1 = Math.max(box.y + box.h, o.y + o.h);
              box.x = nx;
              box.y = ny;
              box.w = nx1 - nx;
              box.h = ny1 - ny;
              o.dead = true;
              absorbed = true;
              break find;
            }
          }
        }
      }
    }
    addIndex(box, accepted.length);
    accepted.push(box);
  }
  const out = [];
  for (let i = 0; i < accepted.length; i++) if (!accepted[i].dead) out.push(accepted[i]);
  return out;
}

/* ------------------------------------------------------------------ *
 * text block detection (optional, textRegions:true)                    *
 * ------------------------------------------------------------------ */

function sceneMedianOf(values) {
  if (!values || values.length === 0) return 0;
  const copy = values.slice().sort((a, b) => a - b);
  return copy[(copy.length - 1) >> 1];
}

/**
 * Detect text-ish blocks and return boxes to erase + place an empty text box over.
 * Heuristic: ink = pixel far from its 16x16 local background median; connected
 * components filtered by stroke height/fill; grouped into lines, then blocks.
 */
function sceneTextDetect(comp, o) {
  const W = comp.width;
  const H = comp.height;
  const n = W * H;
  if (n < 64) return [];
  const rgba = comp.rgba;
  const pack = comp.pack;
  const lum = new Uint8Array(n);
  for (let i = 0, p = 0; i < n; i++, p += 4) {
    lum[i] = (rgba[p] * 77 + rgba[p + 1] * 150 + rgba[p + 2] * 29) >> 8;
  }

  // local background luminance: per-block median. A block that is split between two large flat
  // areas (a panel/photo edge, not ink on a uniform background) cannot measure ink reliably, so it
  // is marked ambiguous and contributes no ink at all.
  const B = 16;
  const bw = Math.ceil(W / B);
  const bh = Math.ceil(H / B);
  const bgLum = new Int32Array(bw * bh);
  const bgMixed = new Uint8Array(bw * bh);
  const hist = new Uint32Array(256);
  for (let by = 0; by < bh; by++) {
    const y0 = by * B;
    const y1 = Math.min(H, y0 + B);
    for (let bx = 0; bx < bw; bx++) {
      const x0 = bx * B;
      const x1 = Math.min(W, x0 + B);
      hist.fill(0);
      let count = 0;
      for (let y = y0; y < y1; y++) {
        const row = y * W;
        for (let x = x0; x < x1; x++) {
          hist[lum[row + x]]++;
          count++;
        }
      }
      let acc = 0;
      const half = count >> 1;
      let med = 0;
      for (let v = 0; v < 256; v++) {
        acc += hist[v];
        if (acc > half) {
          med = v;
          break;
        }
      }
      let far = 0;
      for (let y = y0; y < y1; y++) {
        const row = y * W;
        for (let x = x0; x < x1; x++) {
          const d = lum[row + x] - med;
          if (d >= sceneInkThreshold || d <= -sceneInkThreshold) far++;
        }
      }
      bgLum[by * bw + bx] = med;
      if (far > count * 0.42) bgMixed[by * bw + bx] = 1;
    }
  }

  // ink mask
  const ink = new Uint8Array(n);
  let inkCount = 0;
  for (let y = 0; y < H; y++) {
    const row = y * W;
    const brow = ((y / B) | 0) * bw;
    for (let x = 0; x < W; x++) {
      const i = row + x;
      const b = brow + ((x / B) | 0);
      if (bgMixed[b]) continue;
      const d = lum[i] - bgLum[b];
      if (d >= sceneInkThreshold || d <= -sceneInkThreshold) {
        ink[i] = 1;
        inkCount++;
      }
    }
  }
  if (inkCount < 8) return [];

  // connected components (8-connected)
  const minH = Math.max(1, o.textMinHeightPx);
  const maxH = Math.max(minH, Math.floor(H / 3));
  const seen = new Uint8Array(n);
  const stack = new Int32Array(n);
  const glyphs = [];
  for (let start = 0; start < n; start++) {
    if (!ink[start] || seen[start]) continue;
    let sp = 0;
    stack[sp++] = start;
    seen[start] = 1;
    let x0 = W;
    let y0 = H;
    let x1 = -1;
    let y1 = -1;
    let area = 0;
    while (sp > 0) {
      const cur = stack[--sp];
      const cy = (cur / W) | 0;
      const cx = cur - cy * W;
      area++;
      if (cx < x0) x0 = cx;
      if (cx > x1) x1 = cx;
      if (cy < y0) y0 = cy;
      if (cy > y1) y1 = cy;
      const ya = cy > 0 ? cy - 1 : 0;
      const yb = cy < H - 1 ? cy + 1 : H - 1;
      const xa = cx > 0 ? cx - 1 : 0;
      const xb = cx < W - 1 ? cx + 1 : W - 1;
      for (let yy = ya; yy <= yb; yy++) {
        const ro = yy * W;
        for (let xx = xa; xx <= xb; xx++) {
          const ni = ro + xx;
          if (ink[ni] && !seen[ni]) {
            seen[ni] = 1;
            stack[sp++] = ni;
          }
        }
      }
    }
    const cw = x1 - x0 + 1;
    const ch = y1 - y0 + 1;
    if (ch < minH || ch > maxH) continue;
    if (cw > ch * 8 || ch > cw * 8) continue; // rules/borders: too thin to be a glyph
    const fill = area / (cw * ch);
    if (fill > 0.85) continue; // solid blob: icon/panel/rule, not text
    if (fill < 0.1) continue; // sparse curve/line (chart stroke, border trace), not text
    glyphs.push({ x0, y0, x1, y1 });
  }
  if (glyphs.length < 2) return [];

  // glyphs -> text lines
  glyphs.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const lines = [];
  for (let gi = 0; gi < glyphs.length; gi++) {
    const g = glyphs[gi];
    const gh = g.y1 - g.y0 + 1;
    let best = null;
    let bestGap = Infinity;
    for (let li = 0; li < lines.length; li++) {
      const ln = lines[li];
      const ov = Math.min(ln.y1, g.y1) - Math.max(ln.y0, g.y0) + 1;
      const lh = ln.y1 - ln.y0 + 1;
      if (ov < Math.min(lh, gh) * 0.6) continue;
      const gap = Math.max(0, Math.max(g.x0 - ln.x1 - 1, ln.x0 - g.x1 - 1));
      const maxGap = Math.max(6, Math.max(lh, gh) * 2.5);
      if (gap > maxGap) continue;
      if (gap < bestGap) {
        bestGap = gap;
        best = ln;
      }
    }
    if (best) {
      if (g.x0 < best.x0) best.x0 = g.x0;
      if (g.y0 < best.y0) best.y0 = g.y0;
      if (g.x1 > best.x1) best.x1 = g.x1;
      if (g.y1 > best.y1) best.y1 = g.y1;
    } else {
      lines.push({ x0: g.x0, y0: g.y0, x1: g.x1, y1: g.y1 });
    }
  }

  // lines -> blocks
  lines.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const blocks = [];
  for (let li = 0; li < lines.length; li++) {
    const ln = lines[li];
    const lh = ln.y1 - ln.y0 + 1;
    let target = null;
    for (let bi = 0; bi < blocks.length; bi++) {
      const bk = blocks[bi];
      const last = bk.lines[bk.lines.length - 1];
      const lastH = last.y1 - last.y0 + 1;
      const hmax = Math.max(lh, lastH);
      const hmin = Math.min(lh, lastH);
      if (hmax > hmin * 1.8) continue;
      const gap = ln.y0 - last.y1 - 1;
      if (gap < -hmin * 0.5 || gap > hmax * 1.2) continue;
      if (ln.x0 > bk.x1 + hmax || bk.x0 > ln.x1 + hmax) continue;
      const nx = Math.min(bk.x0, ln.x0);
      const nx1 = Math.max(bk.x1, ln.x1);
      const ny = Math.min(bk.y0, ln.y0);
      const ny1 = Math.max(bk.y1, ln.y1);
      if ((nx1 - nx + 1) * (ny1 - ny + 1) > W * H * 0.45) continue;
      target = bk;
      break;
    }
    if (target) {
      target.lines.push(ln);
      if (ln.x0 < target.x0) target.x0 = ln.x0;
      if (ln.y0 < target.y0) target.y0 = ln.y0;
      if (ln.x1 > target.x1) target.x1 = ln.x1;
      if (ln.y1 > target.y1) target.y1 = ln.y1;
    } else {
      blocks.push({ x0: ln.x0, y0: ln.y0, x1: ln.x1, y1: ln.y1, lines: [ln] });
    }
  }

  // block bboxes -> disjoint pixel boxes
  const raw = [];
  for (let bi = 0; bi < blocks.length && bi < 256; bi++) {
    const bk = blocks[bi];
    raw.push({ x: bk.x0, y: bk.y0, w: bk.x1 - bk.x0 + 1, h: bk.y1 - bk.y0 + 1 });
  }
  if (!raw.length) return [];
  const boxes = sceneDisjointBoxes(raw);

  const out = [];
  for (let bi = 0; bi < boxes.length; bi++) {
    const b = boxes[bi];
    const x1 = Math.min(W, b.x + b.w);
    const y1 = Math.min(H, b.y + b.h);
    if (x1 <= b.x || y1 <= b.y) continue;
    // rows that contain ink -> ink bands -> per-line stroke height
    const heights = [];
    let band = 0;
    let inkRows = 0;
    for (let y = b.y; y < y1; y++) {
      const row = y * W;
      let rowInk = 0;
      for (let x = b.x; x < x1; x++) if (ink[row + x]) rowInk++;
      if (rowInk > 0) {
        inkRows += rowInk;
        band++;
      } else if (band > 0) {
        heights.push(band);
        band = 0;
      }
    }
    if (band > 0) heights.push(band);
    if (inkRows < 8 || heights.length === 0) continue;
    const inkH = sceneMedianOf(heights);
    if (inkH < minH) continue;
    // dominant background colour inside the block (everything that is not ink)
    const counts = new Map();
    let bestPack = -1;
    let bestCount = -1;
    for (let y = b.y; y < y1; y++) {
      const row = y * W;
      for (let x = b.x; x < x1; x++) {
        const i = row + x;
        if (ink[i]) continue;
        const p = pack[i];
        const c = (counts.get(p) || 0) + 1;
        counts.set(p, c);
        if (c > bestCount || (c === bestCount && p < bestPack)) {
          bestCount = c;
          bestPack = p;
        }
      }
    }
    if (bestPack < 0 || bestCount <= 0) continue;
    const sz = Math.max(6, Math.min(96, Math.round(inkH * 0.9 * 0.75)));
    out.push({ x: b.x, y: b.y, w: x1 - b.x, h: y1 - b.y, c: sceneHexFromPack(bestPack), sz });
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * masks / crop helpers                                                *
 * ------------------------------------------------------------------ */

function sceneFillMask(mask, w, h, box) {
  const x0 = Math.max(0, Math.floor(box.x));
  const y0 = Math.max(0, Math.floor(box.y));
  const x1 = Math.min(w, Math.floor(box.x + box.w));
  const y1 = Math.min(h, Math.floor(box.y + box.h));
  if (x1 <= x0) return;
  for (let y = y0; y < y1; y++) mask.fill(1, y * w + x0, y * w + x1);
}

function sceneBuildSkip(w, h, crops, texts) {
  const mask = new Uint8Array(w * h);
  for (let i = 0; i < crops.length; i++) sceneFillMask(mask, w, h, crops[i]);
  for (let i = 0; i < texts.length; i++) sceneFillMask(mask, w, h, texts[i]);
  return mask;
}

/** Force tiles that touch a text box back to "flat" (they will be erased, not cropped). */
function sceneForceTextTiles(metrics, textBoxes) {
  const tw = metrics.tw;
  const th = metrics.th;
  const detail = metrics.detail;
  const tile = sceneTileSize;
  for (let i = 0; i < textBoxes.length; i++) {
    const b = textBoxes[i];
    const tx0 = Math.max(0, Math.floor(b.x / tile));
    const tx1 = Math.min(tw - 1, Math.floor((b.x + b.w - 1) / tile));
    const ty0 = Math.max(0, Math.floor(b.y / tile));
    const ty1 = Math.min(th - 1, Math.floor((b.y + b.h - 1) / tile));
    for (let ty = ty0; ty <= ty1; ty++) {
      for (let tx = tx0; tx <= tx1; tx++) detail[ty * tw + tx] = 0;
    }
  }
}

/** Pixel mask (1 = detail tile) from the per-tile detail flags. */
function sceneDetailPixelMask(metrics, w, h) {
  const mask = new Uint8Array(w * h);
  const tile = sceneTileSize;
  const tw = metrics.tw;
  const detail = metrics.detail;
  for (let ty = 0; ty < metrics.th; ty++) {
    const y0 = ty * tile;
    const y1 = Math.min(h, y0 + tile);
    for (let tx = 0; tx < tw; tx++) {
      if (!detail[ty * tw + tx]) continue;
      const x0 = tx * tile;
      const x1 = Math.min(w, x0 + tile);
      for (let y = y0; y < y1; y++) mask.fill(1, y * w + x0, y * w + x1);
    }
  }
  return mask;
}

/**
 * Pixels that will not be painted by flat rects (detail tiles + text boxes).
 * Used to keep palette ranking focused on the pixels quantisation may legally change.
 */
function sceneAnalysisMask(w, h, metrics, textBoxes) {
  let mask = null;
  if (metrics) mask = sceneDetailPixelMask(metrics, w, h);
  if (textBoxes && textBoxes.length) {
    if (!mask) mask = new Uint8Array(w * h);
    for (let i = 0; i < textBoxes.length; i++) sceneFillMask(mask, w, h, textBoxes[i]);
  }
  return mask;
}

/** Exact RGBA crop of the composited image -> k:'i' shape. */
function sceneCropShape(comp, x, y, w, h, role) {
  const data = new Uint8ClampedArray(w * h * 4);
  const src = comp.rgba;
  const sw = comp.width;
  for (let yy = 0; yy < h; yy++) {
    const s = ((y + yy) * sw + x) * 4;
    data.set(src.subarray(s, s + w * 4), yy * w * 4);
  }
  return { k: 'i', x, y, w, h, px: { data, width: w, height: h }, role };
}

/* ------------------------------------------------------------------ *
 * mode implementations                                                *
 * ------------------------------------------------------------------ */

/** mode:'image' — exactly one picture covering the whole (composited) image. */
function scenePictureResult(comp) {
  return { shapes: [sceneCropShape(comp, 0, 0, comp.width, comp.height, 'photo')], palette: [] };
}

/**
 * Rect cap for the exact tiling pass. Unlimited mode only guards memory. A budgeted mode must be
 * able to tile the flat area exactly before the tolerance ladder may degrade it, otherwise budget
 * pressure would skip degradation step (b) and jump straight to a single picture for no reason.
 */
/**
 * Copy of the normalised options with a raised merge tolerance (used by the budget ladder; the
 * object is already normalised, so no re-parsing is needed).
 */
function sceneWithTolerance(o, tol) {
  const c = {};
  for (const key in o) {
    if (Object.prototype.hasOwnProperty.call(o, key)) c[key] = o[key];
  }
  c.mergeTolerance = tol;
  return c;
}

function sceneTilingCap(o) {
  if (o.shapeBudget > 0) {
    const budgeted = Math.max(o.shapeBudget * 256, 16384);
    return budgeted < sceneRectHardCap ? budgeted : sceneRectHardCap;
  }
  return sceneRectHardCap;
}

/**
 * smart-mode attempt with `grow` tiles of detail-region dilation.
 * Returns rect tiling + crop boxes; materialisation happens once, for the winner.
 */
function sceneSmartAttempt(comp, pack, o, metrics, textBoxes, grow) {
  const W = comp.width;
  const H = comp.height;
  const T = sceneTileSize;
  const mask = grow > 0 ? sceneDilateTiles(metrics.detail, metrics.tw, metrics.th, grow) : metrics.detail;
  const regions = sceneTileRegions(mask, metrics.tw, metrics.th);
  let boxes = sceneBoxesFromTiles(regions, T, W, H);
  if (o.detailMinAreaPx > 0) boxes = boxes.filter((b) => b.w * b.h >= o.detailMinAreaPx);
  boxes = sceneDisjointBoxes(boxes);
  // regionMinPx: do not turn hairline detail into its own picture; those pixels stay exact rects.
  if (o.regionMinPx > 1) boxes = boxes.filter((b) => b.w >= o.regionMinPx && b.h >= o.regionMinPx);
  const skip = sceneBuildSkip(W, H, boxes, textBoxes);
  const cap = sceneTilingCap(o);
  const t = sceneTiling(pack, W, H, skip, cap);
  if (t.aborted) return { aborted: true, n: Infinity, t: null, cropBoxes: [] };
  let count = t.n;
  if (o.mergeTolerance > 0 && count > 1) count = sceneMergeSimilar(t, o.mergeTolerance);
  return { aborted: false, n: count, t, cropBoxes: boxes };
}

/** Vectors/flat decomposition + budget ladder (tolerance steps -> drop text -> one picture). */
function sceneFlatResult(comp, pack, o, notes, textBoxes) {
  const W = comp.width;
  const H = comp.height;
  let texts = textBoxes.slice();
  const skip = texts.length ? sceneBuildSkip(W, H, [], texts) : null;
  const cap = sceneTilingCap(o);
  const t = sceneTiling(pack, W, H, skip, cap);
  if (t.aborted) {
    if (o.shapeBudget > 0 && cap < sceneRectHardCap) {
      notes.push(
        'budget ' +
          o.shapeBudget +
          ': exact tiling needs more than ' +
          cap +
          ' rect(s) -> whole image as one picture',
      );
    } else {
      notes.push('exact tiling exceeds the internal rect cap (' + cap + '); whole image as one picture');
    }
    return scenePictureResult(comp);
  }
  let count = t.n;
  if (o.mergeTolerance > 0 && count > 1) count = sceneMergeSimilar(t, o.mergeTolerance);
  const budget = o.shapeBudget;
  if (budget > 0 && count + texts.length * 2 > budget) {
    let done = false;
    const steps = [1, 2, 4, 8, 16, 32, 64, 128];
    for (let i = 0; i < steps.length; i++) {
      const tol = steps[i];
      if (tol <= o.mergeTolerance) continue;
      count = sceneMergeSimilar(t, tol);
      if (count + texts.length * 2 <= budget) {
        notes.push('budget ' + budget + ': raised merge tolerance to ' + tol + ' -> ' + count + ' rect(s)');
        done = true;
        break;
      }
    }
    if (!done && texts.length && count <= budget) {
      notes.push('budget ' + budget + ': text boxes dropped to respect the budget');
      texts = [];
      done = true;
    }
    if (!done) {
      notes.push('budget ' + budget + ': tiling cannot fit the budget -> whole image as one picture');
      return scenePictureResult(comp);
    }
  }
  return sceneAssemble(comp, t, count, [], texts);
}

/** smart mode (detail regions as crops, everything else exact rectangles). */
function sceneSmartResult(comp, pack, o, notes, metrics, textBoxes) {
  const textShapeCount = textBoxes.length * 2;
  if (!o.detailToImage) {
    return sceneFlatResult(comp, pack, o, notes, textBoxes);
  }
  const budget = o.shapeBudget;
  const textShapes = textShapeCount;
  const levels = [0, 1, 2, 3, 4, 6, 8, 12, 16, 24, 32, 48, 64, 96, 128, 192, 256, 384, 512, 768, 1024];
  let chosen = null;
  let chosenGrow = 0;
  let raisedTolerance = false;
  for (let i = 0; i < levels.length; i++) {
    const att = sceneSmartAttempt(comp, pack, o, metrics, textBoxes, levels[i]);
    if (att.aborted) continue;
    const total = att.n + att.cropBoxes.length + textShapes;
    if (budget === 0 || total <= budget) {
      chosen = att;
      chosenGrow = levels[i];
      break;
    }
  }
  if (!chosen) {
    // Degradation step (b): the detail regions are already grown as far as they go, so raise the
    // effective merge tolerance in small steps before giving up on editable flat rects entirely.
    let bestAtt = null;
    let bestGrow = 0;
    for (let i = levels.length - 1; i >= 0; i--) {
      const att = sceneSmartAttempt(comp, pack, o, metrics, textBoxes, levels[i]);
      if (!att.aborted) {
        bestAtt = att;
        bestGrow = levels[i];
        break;
      }
    }
    if (bestAtt) {
      const steps = [1, 2, 4, 8, 16, 32, 64, 128];
      for (let i = 0; i < steps.length; i++) {
        const tol = steps[i];
        if (tol <= o.mergeTolerance) continue;
        const att = sceneSmartAttempt(comp, pack, sceneWithTolerance(o, tol), metrics, textBoxes, bestGrow);
        if (att.aborted) continue;
        if (att.n + att.cropBoxes.length + textShapes <= budget) {
          notes.push(
            'budget ' +
              budget +
              ': raised merge tolerance to ' +
              tol +
              ' (maxAbs <= ' +
              tol +
              ' per channel) -> ' +
              (att.n + att.cropBoxes.length + textShapes) +
              ' shape(s)',
          );
          notes.push('fidelity below 1: merged flat colours are approximate, crops stay exact');
          chosen = att;
          chosenGrow = bestGrow;
          raisedTolerance = true;
          break;
        }
      }
    }
  }
  if (!chosen) {
    notes.push('budget ' + budget + ': tiling cannot fit the budget even at max tolerance -> whole image as one picture');
    return scenePictureResult(comp);
  }
  if (chosenGrow > 0 && !raisedTolerance) {
    notes.push(
      'budget ' +
        budget +
        ': detail regions grown by ' +
        chosenGrow +
        ' tile(s) -> ' +
        (chosen.n + chosen.cropBoxes.length + textShapes) +
        ' shape(s)',
    );
  }
  return sceneAssemble(comp, chosen.t, chosen.n, chosen.cropBoxes, textBoxes);
}

/**
 * Materialise the final shape list (paint order: flat rects, crops, erase rect + text box)
 * and compute the palette of rect colours by descending area.
 */
function sceneAssemble(comp, t, count, cropBoxes, textBoxes) {
  const hexCache = new Map();
  const hexOf = (p) => {
    let h = hexCache.get(p);
    if (h === undefined) {
      h = sceneHexFromPack(p);
      hexCache.set(p, h);
    }
    return h;
  };
  const areaMap = new Map();
  const shapes = [];
  if (t) {
    const n = Math.min(count, t.n);
    for (let i = 0; i < n; i++) {
      const p = t.Pp[i];
      shapes.push({ k: 'r', x: t.X[i], y: t.Y[i], w: t.Wd[i], h: t.Hh[i], c: hexOf(p), role: 'flat' });
      areaMap.set(p, (areaMap.get(p) || 0) + t.Wd[i] * t.Hh[i]);
    }
  }
  for (let i = 0; i < cropBoxes.length; i++) {
    const b = cropBoxes[i];
    shapes.push(sceneCropShape(comp, b.x, b.y, b.w, b.h, 'detail'));
  }
  for (let i = 0; i < textBoxes.length; i++) {
    const b = textBoxes[i];
    shapes.push({ k: 'r', x: b.x, y: b.y, w: b.w, h: b.h, c: b.c, role: 'bg' });
    areaMap.set(b.c, (areaMap.get(b.c) || 0) + b.w * b.h); // palette is rect-colour based
    shapes.push({
      k: 't',
      x: b.x,
      y: b.y,
      w: b.w,
      h: b.h,
      s: '',
      sz: b.sz,
      c: b.c,
      al: 'l',
      b: false,
      f: 'Arial',
      role: 'text',
    });
  }
  const entries = [];
  areaMap.forEach((area, p) => entries.push([p, area]));
  entries.sort((a, b) => b[1] - a[1] || a[0] - b[0]);
  const palette = entries.map((e) => hexOf(e[0]));
  return { shapes, palette };
}

/* ------------------------------------------------------------------ *
 * sceneRasterize / sceneDiff                                         *
 * ------------------------------------------------------------------ */

/** Painter's-algorithm rasteriser of a scene back to straight-alpha RGBA. */
function sceneRasterize(scene) {
  const sc = scene && typeof scene === 'object' ? scene : {};
  const width = Math.max(1, sceneInt(sc.width, 0, 200000, 1));
  const height = Math.max(1, sceneInt(sc.height, 0, 200000, 1));
  const n = width * height;
  const data = new Uint8ClampedArray(n * 4);
  const bg = sceneHexToRgb(sceneHexOpt(sc.background, '#ffffff'), '#ffffff');
  for (let i = 0, p = 0; i < n; i++, p += 4) {
    data[p] = bg[0];
    data[p + 1] = bg[1];
    data[p + 2] = bg[2];
    data[p + 3] = 255;
  }
  const shapes = Array.isArray(sc.shapes) ? sc.shapes : [];
  for (let si = 0; si < shapes.length; si++) {
    const s = shapes[si];
    if (!s || typeof s !== 'object') continue;
    const sx = Math.max(0, Math.floor(sceneNum(s.x, -1e6, 1e6, 0)));
    const sy = Math.max(0, Math.floor(sceneNum(s.y, -1e6, 1e6, 0)));
    const sw = Math.floor(sceneNum(s.w, 0, 1e6, 0));
    const sh = Math.floor(sceneNum(s.h, 0, 1e6, 0));
    if (sw <= 0 || sh <= 0) continue;
    const x0 = sx;
    const y0 = sy;
    const x1 = Math.min(width, sx + sw);
    const y1 = Math.min(height, sy + sh);
    if (x1 <= x0 || y1 <= y0) continue;
    if (s.k === 'i') {
      const px = s.px;
      if (!px || !px.data || typeof px.data.length !== 'number') continue;
      const pw = Math.max(0, Math.floor(sceneNum(px.width, 0, 1e6, sw)));
      const ph = Math.max(0, Math.floor(sceneNum(px.height, 0, 1e6, sh)));
      const pd = px.data;
      for (let y = y0; y < y1; y++) {
        const ry = y - sy;
        if (ry >= ph) break;
        const srow = ry * pw * 4;
        let drow = (y * width + x0) * 4;
        for (let x = x0; x < x1; x++, drow += 4) {
          const rx = x - sx;
          if (rx >= pw) break;
          const sp = srow + rx * 4;
          if (sp + 3 >= pd.length) break;
          data[drow] = pd[sp];
          data[drow + 1] = pd[sp + 1];
          data[drow + 2] = pd[sp + 2];
          data[drow + 3] = pd[sp + 3];
        }
      }
    } else {
      const col = sceneHexToRgb(sceneHexOpt(s.c, '#000000'), '#000000');
      const r = col[0];
      const g = col[1];
      const b = col[2];
      for (let y = y0; y < y1; y++) {
        let p = (y * width + x0) * 4;
        for (let x = x0; x < x1; x++, p += 4) {
          data[p] = r;
          data[p + 1] = g;
          data[p + 2] = b;
          data[p + 3] = 255;
        }
      }
    }
  }
  return { width, height, data };
}

/**
 * Compare two RGBA buffers. mismatchPixels counts pixels with any channel delta;
 * ignoreRects are excluded from the `exactRatioIgnoring` denominator only.
 */
function sceneDiff(a, b, opts) {
  const A = a && typeof a === 'object' ? a : {};
  const B = b && typeof b === 'object' ? b : {};
  const width = Math.max(1, sceneInt(A.width, 0, 200000, 1));
  const height = Math.max(1, sceneInt(A.height, 0, 200000, 1));
  const aw = width;
  const ah = height;
  const bw = Math.max(1, sceneInt(B.width, 0, 200000, 1));
  const bh = Math.max(1, sceneInt(B.height, 0, 200000, 1));
  const adata = A.data && typeof A.data === 'object' && typeof A.data.length === 'number' ? A.data : null;
  const bdata = B.data && typeof B.data === 'object' && typeof B.data.length === 'number' ? B.data : null;
  const pixels = width * height;
  const heatmap = new Uint8ClampedArray(pixels * 4);

  const rects = opts && typeof opts === 'object' && Array.isArray(opts.ignoreRects) ? opts.ignoreRects : null;
  let ignore = null;
  if (rects) {
    ignore = new Uint8Array(pixels);
    for (let i = 0; i < rects.length; i++) {
      const r = rects[i];
      if (!r || typeof r !== 'object') continue;
      const rx = Math.floor(sceneNum(r.x, -1e6, 1e6, 0));
      const ry = Math.floor(sceneNum(r.y, -1e6, 1e6, 0));
      const rw = Math.floor(sceneNum(r.w, 0, 1e6, 0));
      const rh = Math.floor(sceneNum(r.h, 0, 1e6, 0));
      if (rw <= 0 || rh <= 0) continue;
      const x0 = Math.max(0, rx);
      const y0 = Math.max(0, ry);
      const x1 = Math.min(width, rx + rw);
      const y1 = Math.min(height, ry + rh);
      if (x1 <= x0 || y1 <= y0) continue;
      for (let y = y0; y < y1; y++) ignore.fill(1, y * width + x0, y * width + x1);
    }
  }

  let mismatch = 0;
  let ignoredMismatch = 0;
  let ignoredCount = 0;
  let sum = 0;
  let maxAbs = 0;
  for (let y = 0; y < ah; y++) {
    const arow = y * aw;
    const brow2 = y * bw;
    const bRowOk = bdata && y < bh;
    for (let x = 0; x < aw; x++) {
      const i = arow + x;
      const p = i * 4;
      let d0 = 0;
      let d1 = 0;
      let d2 = 0;
      let d3 = 0;
      if (adata && p + 3 < adata.length) {
        d0 = adata[p];
        d1 = adata[p + 1];
        d2 = adata[p + 2];
        d3 = adata[p + 3];
      }
      let e0 = 0;
      let e1 = 0;
      let e2 = 0;
      let e3 = 0;
      if (bRowOk && x < bw) {
        const q = (brow2 + x) * 4;
        if (q + 3 < bdata.length) {
          e0 = bdata[q];
          e1 = bdata[q + 1];
          e2 = bdata[q + 2];
          e3 = bdata[q + 3];
        }
      }
      let dr = d0 - e0;
      if (dr < 0) dr = -dr;
      let dg = d1 - e1;
      if (dg < 0) dg = -dg;
      let db = d2 - e2;
      if (db < 0) db = -db;
      let da = d3 - e3;
      if (da < 0) da = -da;
      let di = dr;
      if (dg > di) di = dg;
      if (db > di) di = db;
      if (da > di) di = da;
      sum += di;
      if (di > maxAbs) maxAbs = di;
      if (di > 0) {
        mismatch++;
        heatmap[i * 4] = 255;
        heatmap[i * 4 + 1] = 0;
        heatmap[i * 4 + 2] = 0;
        heatmap[i * 4 + 3] = di * 2 + 48 > 255 ? 255 : di * 2 + 48;
      }
      // ignoreRects never change the raw error (mismatchPixels / exactRatio / meanAbs / maxAbs
      // / heatmap); they only take pixels out of the exactRatioIgnoring denominator.
      if (ignore && ignore[i]) {
        ignoredCount++;
        if (di > 0) ignoredMismatch++;
      }
    }
  }
  const denom = pixels - ignoredCount;
  let exactRatioIgnoring = null;
  if (ignore) {
    if (denom > 0) {
      const nonIgnoredMismatch = mismatch - ignoredMismatch;
      exactRatioIgnoring = (denom - nonIgnoredMismatch) / denom;
    }
  }
  return {
    pixels,
    mismatchPixels: mismatch,
    exactRatio: pixels > 0 ? (pixels - mismatch) / pixels : 1,
    meanAbs: pixels > 0 ? sum / pixels : 0,
    maxAbs,
    exactRatioIgnoring,
    heatmap,
  };
}

/* ------------------------------------------------------------------ *
 * stats + public entry point                                          *
 * ------------------------------------------------------------------ */

function sceneFinalize(o, comp, shapes, palette, notes, startedAt) {
  const scene = {
    width: comp.width,
    height: comp.height,
    background: o.alphaBackground,
    shapes,
    palette,
    stats: {
      shapes: 0,
      rects: 0,
      images: 0,
      texts: 0,
      paletteSize: 0,
      exactRatio: 1,
      mismatchPixels: 0,
      meanAbs: 0,
      maxAbs: 0,
      ms: 0,
      notes,
    },
  };
  const raster = sceneRasterize(scene);
  const diff = sceneDiff({ width: comp.width, height: comp.height, data: comp.rgba }, raster, null);
  const st = scene.stats;
  st.shapes = shapes.length;
  for (let i = 0; i < shapes.length; i++) {
    const k = shapes[i].k;
    if (k === 'r') st.rects++;
    else if (k === 'i') st.images++;
    else st.texts++;
  }
  st.paletteSize = palette.length;
  st.exactRatio = diff.exactRatio;
  st.mismatchPixels = diff.mismatchPixels;
  st.meanAbs = Math.round(diff.meanAbs * 10000) / 10000;
  st.maxAbs = diff.maxAbs;
  st.ms = Math.max(0, Math.round(sceneNow() - startedAt));
  return scene;
}

/**
 * Build a scene from an RGBA raster.
 * @param {{width:number,height:number,data:Uint8ClampedArray}} image
 * @param {object} [options] see sceneDefaults
 * @returns {{scale:number, scene:object}}
 */
function sceneBuild(image, options) {
  const startedAt = sceneNow();
  const o = sceneParseOptions(options);
  const notes = [];
  const src = sceneNormalizeImage(image);
  if (src.shrunk) notes.push('source downscaled to ' + src.width + 'x' + src.height + ' for memory safety');
  const comp = sceneCompose(src, o.alphaBackground);
  if (comp.transparent > 0) {
    notes.push('composited ' + comp.transparent + ' transparent pixel(s) over ' + o.alphaBackground);
  }
  const edge = Math.max(comp.width, comp.height);
  const scale = edge > o.maxSlidePx ? o.maxSlidePx / edge : 1;
  if (scale < 1) {
    notes.push('slide scale ' + Math.round(scale * 1000) / 10 + '% (' + edge + 'px > ' + o.maxSlidePx + 'px cap)');
  }

  // ---- analysis first: region classification is quantisation-independent ----
  // Text detection always reads the ORIGINAL composited raster, never the reduced palette.
  const textBoxes = o.textRegions && o.mode !== 'image' ? sceneTextDetect(comp, o) : [];
  let metrics = null;
  if (o.mode === 'smart' && o.detailToImage) {
    metrics = sceneTileMetrics(comp.pack, comp.width, comp.height, sceneTileSize);
    if (textBoxes.length) sceneForceTextTiles(metrics, textBoxes);
  }
  if (o.textRegions && o.mode !== 'image' && textBoxes.length === 0) {
    notes.push('textRegions on: no text block detected');
  }
  if (textBoxes.length) {
    notes.push(
      'text: ' + textBoxes.length + ' block(s) -> ' + textBoxes.length * 2 + ' shape(s); ink pixels excluded from fidelity',
    );
  }

  // ---- optional palette quantisation: only flat pixels may change colour ----
  let pack = comp.pack;
  if (o.maxColors > 0 && o.mode !== 'image') {
    const exclude = sceneAnalysisMask(comp.width, comp.height, metrics, textBoxes);
    const q = sceneQuantizePacked(pack, o.maxColors, exclude);
    if (q.changed) {
      pack = q.pack;
      notes.push('quantised to ' + q.colors + ' colour(s) (maxColors ' + o.maxColors + '); crops stay exact');
    }
  }

  let result;
  if (o.mode === 'image') {
    result = scenePictureResult(comp);
  } else if (o.mode === 'vector') {
    result = sceneFlatResult(comp, pack, o, notes, textBoxes);
  } else {
    result = sceneSmartResult(comp, pack, o, notes, metrics, textBoxes);
  }

  const scene = sceneFinalize(o, comp, result.shapes, result.palette, notes, startedAt);
  return { scale, scene };
}

/* ==================== src/pptx.frag.js ==================== */
/* =============================================================================
 * dsh-pixeltrace - PowerPoint (OOXML PresentationML) writer fragment.
 *
 * Plain script fragment: no import/export statements, no top-level await, no DOM.
 * All top-level identifiers are prefixed `pptx`.  Concatenated into the client
 * bundle by the Lead and evaluated with `new Function(...)` in Node tests.
 *
 * Public API (CONTRACT.md section 5):
 *   async pptxEncodePng(rgba, width, height) -> Promise<Uint8Array>
 *   async pptxDecodePng(bytes)               -> Promise<{width,height,data}>
 *   async pptxBuild(slides, options)         -> Promise<Uint8Array>
 *   async pptxZip(files)                     -> Promise<Uint8Array>
 *   async pptxUnzip(bytes)                   -> Promise<Map<string,Uint8Array>>
 *
 * The package reproduces the source scene exactly: every object is an editable
 * PowerPoint object placed at round(px * emuPerPx * scale) EMU (1 px = 9525 EMU
 * at scale 1), so a rasteriser reproduces the source image pixel for pixel.
 * ========================================================================== */

/* ---------------------------------------------------------------- constants */

var pptxXmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

var pptxNsDrawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
var pptxNsRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var pptxNsPresentation = 'http://schemas.openxmlformats.org/presentationml/2006/main';
var pptxNsPackageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
var pptxNsContentTypes = 'http://schemas.openxmlformats.org/package/2006/content-types';
var pptxNsCoreProperties = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
var pptxNsDc = 'http://purl.org/dc/elements/1.1/';
var pptxNsDcTerms = 'http://purl.org/dc/terms/';
var pptxNsDcMiType = 'http://purl.org/dc/dcmitype/';
var pptxNsXsi = 'http://www.w3.org/2001/XMLSchema-instance';
var pptxNsExtendedProperties = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
var pptxNsVt = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

var pptxRelType = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
};

var pptxContentType = {
  presentation: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  coreProperties: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProperties: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  png: 'image/png'
};

/* Fixed timestamp so identical input gives identical bytes (CONTRACT 2.4). */
var pptxFixedTimestamp = '2020-01-01T00:00:00Z';
var pptxFixedDosTime = 0;
var pptxFixedDosDate = 0x0021; /* 1980-01-01 */

var pptxDefaultEmuPerPx = 9525; /* 96 dpi: 1 px = 9525 EMU */
var pptxDefaultMaxSlideEmu = 51206400; /* 56 in */
var pptxEmuPerPxCap = 9525;

/* --------------------------------------------------------------- utilities */

var pptxEncoderCache = null;
var pptxDecoderCache = null;

/** UTF-8 encode a string. */
function pptxUtf8(text) {
  if (pptxEncoderCache === null) pptxEncoderCache = new TextEncoder();
  return pptxEncoderCache.encode(text);
}

/** UTF-8 decode bytes. */
function pptxUtf8Decode(bytes) {
  if (pptxDecoderCache === null) pptxDecoderCache = new TextDecoder('utf-8');
  return pptxDecoderCache.decode(bytes);
}

/** Coerce binary-ish input to a Uint8Array view (no copy). */
function pptxToBytes(value) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof Uint8ClampedArray) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (value && typeof value === 'object' && typeof value.byteLength === 'number' && value.buffer) {
    return new Uint8Array(value.buffer, value.byteOffset || 0, value.byteLength);
  }
  if (Array.isArray(value)) return Uint8Array.from(value);
  throw new TypeError('pptx: expected binary data (Uint8Array)');
}

/** Finite number or fallback. */
function pptxNumber(value, fallback) {
  var n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

/** Concatenate Uint8Array chunks. */
function pptxConcat(chunks) {
  var total = 0;
  var i;
  for (i = 0; i < chunks.length; i++) total += chunks[i].length;
  var out = new Uint8Array(total);
  var offset = 0;
  for (i = 0; i < chunks.length; i++) {
    out.set(chunks[i], offset);
    offset += chunks[i].length;
  }
  return out;
}

/** Byte equality (used for media de-duplication, CONTRACT 5 "byte equality"). */
function pptxBytesEqual(a, b) {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/** FNV-1a 32 bit hash, only used as a bucket key for byte comparison. */
function pptxHashBytes(bytes) {
  var hash = 0x811c9dc5;
  for (var i = 0; i < bytes.length; i++) {
    hash ^= bytes[i];
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash >>> 0;
}

/** Escape a string for XML text/attribute content and drop illegal controls. */
var pptxXmlEscapeMap = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' };

function pptxXmlEscape(value) {
  var text = value === null || value === undefined ? '' : String(value);
  text = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  return text.replace(/[&<>"']/g, function (ch) {
    return pptxXmlEscapeMap[ch];
  });
}

/** Normalise a colour to lowercase 6 digit hex (no leading '#'), CONTRACT 2.7. */
function pptxColorHex(value, fallback) {
  var fallbackHex = fallback || 'ffffff';
  if (typeof value !== 'string') return fallbackHex;
  var text = value.trim();
  if (text.charAt(0) === '#') text = text.slice(1);
  if (!/^[0-9a-fA-F]{6}$/.test(text)) return fallbackHex;
  return text.toLowerCase();
}

/** Readable text colour for a solid fill: dark ink on light fills, light on dark. */
function pptxContrastTextColor(fillHex) {
  var r = parseInt(fillHex.slice(0, 2), 16) / 255;
  var g = parseInt(fillHex.slice(2, 4), 16) / 255;
  var b = parseInt(fillHex.slice(4, 6), 16) / 255;
  var luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return luminance > 0.55 ? '000000' : 'ffffff';
}

/** Sanitise a deck option string. */
function pptxOptionText(value, fallback) {
  if (typeof value === 'string') {
    var text = value.trim();
    if (text !== '') return text.slice(0, 500);
    return fallback;
  }
  if (typeof value === 'number' && Number.isFinite(value)) return String(value);
  return fallback;
}

/* ------------------------------------------------------------------ CRC-32 */

var pptxCrcTableCache = null;

function pptxCrcTable() {
  if (pptxCrcTableCache !== null) return pptxCrcTableCache;
  var table = new Int32Array(256);
  for (var n = 0; n < 256; n++) {
    var value = n;
    for (var bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[n] = value;
  }
  pptxCrcTableCache = table;
  return table;
}

/** CRC-32 of bytes [start, end). */
function pptxCrc32(bytes, start, end) {
  var table = pptxCrcTable();
  var crc = -1;
  var from = start === undefined || start === null ? 0 : start;
  var to = end === undefined || end === null ? bytes.length : end;
  for (var i = from; i < to; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

/* ------------------------------------------------- deflate / inflate (zlib) */

var pptxHasRawDeflateCache = null;
var pptxNodeZlibCache = null;

/** True when the platform has CompressionStream('deflate-raw'). */
function pptxSupportsRawDeflate() {
  if (pptxHasRawDeflateCache !== null) return pptxHasRawDeflateCache;
  pptxHasRawDeflateCache = false;
  if (typeof CompressionStream === 'function') {
    try {
      /* Constructing validates the format token without consuming data. */
      void new CompressionStream('deflate-raw');
      pptxHasRawDeflateCache = true;
    } catch (error) {
      pptxHasRawDeflateCache = false;
    }
  }
  return pptxHasRawDeflateCache;
}

/** Read a whole ReadableStream into one Uint8Array. */
async function pptxReadStream(readable) {
  var reader = readable.getReader();
  var chunks = [];
  var total = 0;
  for (;;) {
    var step = await reader.read();
    if (step.done) break;
    var chunk = step.value;
    chunks.push(chunk);
    total += chunk.length;
  }
  var out = new Uint8Array(total);
  var offset = 0;
  for (var i = 0; i < chunks.length; i++) {
    out.set(chunks[i], offset);
    offset += chunks[i].length;
  }
  return out;
}

/** Start draining a transform stream before writing, to avoid back-pressure deadlock. */
async function pptxTransformStream(format, direction, bytes) {
  var stream = direction === 'compress' ? new CompressionStream(format) : new DecompressionStream(format);
  var collecting = pptxReadStream(stream.readable);
  var writer = stream.writable.getWriter();
  await writer.write(bytes);
  await writer.close();
  var out = await collecting;
  return out;
}

async function pptxNodeZlib() {
  if (pptxNodeZlibCache === null) {
    pptxNodeZlibCache = await import('node:zlib');
  }
  return pptxNodeZlibCache;
}

var pptxHasRawInflateCache = null;

/** True when the platform has DecompressionStream('deflate-raw'). */
function pptxSupportsRawInflate() {
  if (pptxHasRawInflateCache !== null) return pptxHasRawInflateCache;
  pptxHasRawInflateCache = false;
  if (typeof DecompressionStream === 'function') {
    try {
      void new DecompressionStream('deflate-raw');
      pptxHasRawInflateCache = true;
    } catch (error) {
      pptxHasRawInflateCache = false;
    }
  }
  return pptxHasRawInflateCache;
}

/** raw DEFLATE (ZIP method 8). */
async function pptxDeflateRaw(bytes) {
  if (pptxSupportsRawDeflate()) {
    return await pptxTransformStream('deflate-raw', 'compress', bytes);
  }
  var zlib = await pptxNodeZlib();
  return new Uint8Array(zlib.deflateRawSync(bytes));
}

/** zlib wrapped DEFLATE stream (PNG IDAT). */
async function pptxDeflateZlib(bytes) {
  if (typeof CompressionStream === 'function') {
    try {
      return await pptxTransformStream('deflate', 'compress', bytes);
    } catch (error) {
      /* Fall through to node:zlib below. */
    }
  }
  var zlib = await pptxNodeZlib();
  return new Uint8Array(zlib.deflateSync(bytes));
}

/** raw DEFLATE input (ZIP method 8). */
async function pptxInflateRaw(bytes) {
  if (pptxSupportsRawInflate()) {
    return await pptxTransformStream('deflate-raw', 'decompress', bytes);
  }
  var zlib = await pptxNodeZlib();
  return new Uint8Array(zlib.inflateRawSync(bytes));
}

/** zlib wrapped DEFLATE stream (PNG IDAT). */
async function pptxInflateZlib(bytes) {
  if (typeof DecompressionStream === 'function') {
    return await pptxTransformStream('deflate', 'decompress', bytes);
  }
  var zlib = await pptxNodeZlib();
  return new Uint8Array(zlib.inflateSync(bytes));
}

/* --------------------------------------------------------------------- PNG */

var pptxPngSignature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
var pptxPngIdatChunkSize = 512 * 1024;

function pptxPngChunk(type, data) {
  var typeBytes = pptxUtf8(type);
  var out = new Uint8Array(12 + data.length);
  var view = new DataView(out.buffer);
  view.setUint32(0, data.length, false);
  out.set(typeBytes, 4);
  out.set(data, 8);
  view.setUint32(8 + data.length, pptxCrc32(out, 4, 8 + data.length), false);
  return out;
}

/** Cost of one filtered row: sum of signed byte magnitudes (PNG heuristics). */
function pptxPngRowCost(row) {
  var cost = 0;
  for (var i = 0; i < row.length; i++) {
    var byteValue = row[i];
    cost += byteValue < 128 ? byteValue : 256 - byteValue;
  }
  return cost;
}

var pptxPngPaeth = function (a, b, c) {
  var p = a + b - c;
  var pa = p > a ? p - a : a - p;
  var pb = p > b ? p - b : b - p;
  var pc = p > c ? p - c : c - p;
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
};

/**
 * Adaptive per-row filtering (None/Sub/Up/Avg/Paeth, lowest cost wins, ties to
 * the lower filter type). Deterministic.
 */
function pptxPngFilterRows(rgba, width, height) {
  var stride = width * 4;
  var out = new Uint8Array(height * (stride + 1));
  var prev = new Uint8Array(stride);
  var none = new Uint8Array(stride);
  var sub = new Uint8Array(stride);
  var up = new Uint8Array(stride);
  var average = new Uint8Array(stride);
  var paeth = new Uint8Array(stride);
  var rowIndex;
  for (rowIndex = 0; rowIndex < height; rowIndex++) {
    var rowStart = rowIndex * stride;
    var row = rgba.subarray(rowStart, rowStart + stride);
    var byteIndex;
    for (byteIndex = 0; byteIndex < stride; byteIndex++) {
      var current = row[byteIndex];
      var left = byteIndex >= 4 ? row[byteIndex - 4] : 0;
      var above = prev[byteIndex];
      var aboveLeft = byteIndex >= 4 ? prev[byteIndex - 4] : 0;
      none[byteIndex] = current;
      sub[byteIndex] = (current - left) & 0xff;
      up[byteIndex] = (current - above) & 0xff;
      average[byteIndex] = (current - ((left + above) >>> 1)) & 0xff;
      paeth[byteIndex] = (current - pptxPngPaeth(left, above, aboveLeft)) & 0xff;
    }
    var costs = [
      pptxPngRowCost(none),
      pptxPngRowCost(sub),
      pptxPngRowCost(up),
      pptxPngRowCost(average),
      pptxPngRowCost(paeth)
    ];
    var best = 0;
    for (var candidate = 1; candidate < costs.length; candidate++) {
      if (costs[candidate] < costs[best]) best = candidate;
    }
    var outputStart = rowIndex * (stride + 1);
    out[outputStart] = best;
    var chosen = best === 0 ? none : best === 1 ? sub : best === 2 ? up : best === 3 ? average : paeth;
    out.set(chosen, outputStart + 1);
    prev.set(row);
  }
  return out;
}

/** Encode straight RGBA (row major, 4 bytes per pixel) as an 8 bit RGBA PNG. */
async function pptxEncodePng(rgba, width, height) {
  var w = Math.round(pptxNumber(width, 0));
  var h = Math.round(pptxNumber(height, 0));
  if (w <= 0 || h <= 0) {
    throw new TypeError('pptxEncodePng: width and height must be positive integers');
  }
  var data = pptxToBytes(rgba);
  var needed = w * h * 4;
  if (data.length < needed) {
    throw new TypeError('pptxEncodePng: rgba needs ' + needed + ' bytes for ' + w + 'x' + h + ', got ' + data.length);
  }
  var pixels = data.length === needed ? data : data.subarray(0, needed);
  var filtered = pptxPngFilterRows(pixels, w, h);
  var compressed = await pptxDeflateZlib(filtered);

  var ihdr = new Uint8Array(13);
  var ihdrView = new DataView(ihdr.buffer);
  ihdrView.setUint32(0, w, false);
  ihdrView.setUint32(4, h, false);
  ihdr[8] = 8; /* bit depth */
  ihdr[9] = 6; /* colour type: truecolour + alpha */
  ihdr[10] = 0; /* compression: deflate */
  ihdr[11] = 0; /* filter method: adaptive */
  ihdr[12] = 0; /* interlace: none */

  var parts = [new Uint8Array(pptxPngSignature)];
  parts.push(pptxPngChunk('IHDR', ihdr));
  for (var offset = 0; offset < compressed.length; offset += pptxPngIdatChunkSize) {
    var end = Math.min(compressed.length, offset + pptxPngIdatChunkSize);
    parts.push(pptxPngChunk('IDAT', compressed.subarray(offset, end)));
  }
  parts.push(pptxPngChunk('IEND', new Uint8Array(0)));
  return pptxConcat(parts);
}

/** Map colour type to channel count for 8 bit non-interlaced PNGs. */
function pptxPngChannels(colourType) {
  if (colourType === 0) return 1;
  if (colourType === 2) return 3;
  if (colourType === 3) return 1;
  if (colourType === 4) return 2;
  if (colourType === 6) return 4;
  return 0;
}

/**
 * Decode PNG bytes to straight RGBA. Supports 8 bit colour types 0/2/3/4/6,
 * non-interlaced (our own encoder is type 6). Verifies every chunk CRC.
 */
async function pptxDecodePng(bytes) {
  var buf = pptxToBytes(bytes);
  if (buf.length < 8) throw new Error('pptxDecodePng: file is too short to be a PNG');
  for (var s = 0; s < 8; s++) {
    if (buf[s] !== pptxPngSignature[s]) throw new Error('pptxDecodePng: bad PNG signature');
  }
  var view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  var position = 8;
  var width = 0;
  var height = 0;
  var bitDepth = 0;
  var colourType = 0;
  var compression = 0;
  var filterMethod = 0;
  var interlace = 0;
  var palette = null;
  var transparency = null;
  var idatChunks = [];
  var sawHeader = false;
  var sawEnd = false;

  while (position + 12 <= buf.length) {
    var length = view.getUint32(position, false);
    var type = pptxUtf8Decode(buf.subarray(position + 4, position + 8));
    var dataStart = position + 8;
    if (dataStart + length + 4 > buf.length) {
      throw new Error('pptxDecodePng: truncated ' + type + ' chunk');
    }
    var storedCrc = view.getUint32(dataStart + length, false);
    if (pptxCrc32(buf, position + 4, dataStart + length) !== storedCrc) {
      throw new Error('pptxDecodePng: CRC-32 mismatch in ' + type + ' chunk');
    }
    var chunk = buf.subarray(dataStart, dataStart + length);
    if (type === 'IHDR') {
      if (length !== 13) throw new Error('pptxDecodePng: IHDR must be 13 bytes');
      width = view.getUint32(dataStart, false);
      height = view.getUint32(dataStart + 4, false);
      bitDepth = chunk[8];
      colourType = chunk[9];
      compression = chunk[10];
      filterMethod = chunk[11];
      interlace = chunk[12];
      sawHeader = true;
    } else if (type === 'PLTE') {
      palette = chunk.slice();
    } else if (type === 'tRNS') {
      transparency = chunk.slice();
    } else if (type === 'IDAT') {
      idatChunks.push(chunk);
    } else if (type === 'IEND') {
      sawEnd = true;
      break;
    }
    position = dataStart + length + 4;
  }

  if (!sawHeader) throw new Error('pptxDecodePng: missing IHDR');
  if (!sawEnd) throw new Error('pptxDecodePng: missing IEND');
  if (width <= 0 || height <= 0) throw new Error('pptxDecodePng: invalid image size');
  if (bitDepth !== 8) throw new Error('pptxDecodePng: only 8 bit PNGs are supported (got ' + bitDepth + ' bit)');
  if (compression !== 0) throw new Error('pptxDecodePng: unsupported compression method ' + compression);
  if (filterMethod !== 0) throw new Error('pptxDecodePng: unsupported filter method ' + filterMethod);
  if (interlace !== 0) throw new Error('pptxDecodePng: interlaced PNGs are not supported');
  var channels = pptxPngChannels(colourType);
  if (channels === 0) throw new Error('pptxDecodePng: unsupported colour type ' + colourType);
  if (colourType === 3 && palette === null) throw new Error('pptxDecodePng: palette image without PLTE');
  if (idatChunks.length === 0) throw new Error('pptxDecodePng: no IDAT data');

  var raw = await pptxInflateZlib(pptxConcat(idatChunks));
  var stride = width * channels;
  if (raw.length < height * (stride + 1)) {
    throw new Error('pptxDecodePng: compressed data is shorter than the image (' + raw.length + ' bytes)');
  }

  var filtered = new Uint8Array(stride * height);
  var previous = new Uint8Array(stride);
  var cursor = 0;
  var y;
  for (y = 0; y < height; y++) {
    var filterType = raw[cursor++];
    if (cursor + stride > raw.length) throw new Error('pptxDecodePng: truncated filtered rows');
    var source = raw.subarray(cursor, cursor + stride);
    cursor += stride;
    var row = filtered.subarray(y * stride, (y + 1) * stride);
    var i;
    if (filterType === 0) {
      row.set(source);
    } else if (filterType === 1) {
      for (i = 0; i < stride; i++) {
        row[i] = (source[i] + (i >= channels ? row[i - channels] : 0)) & 0xff;
      }
    } else if (filterType === 2) {
      for (i = 0; i < stride; i++) row[i] = (source[i] + previous[i]) & 0xff;
    } else if (filterType === 3) {
      for (i = 0; i < stride; i++) {
        var left = i >= channels ? row[i - channels] : 0;
        row[i] = (source[i] + ((left + previous[i]) >>> 1)) & 0xff;
      }
    } else if (filterType === 4) {
      for (i = 0; i < stride; i++) {
        var a = i >= channels ? row[i - channels] : 0;
        var b = previous[i];
        var c = i >= channels ? previous[i - channels] : 0;
        row[i] = (source[i] + pptxPngPaeth(a, b, c)) & 0xff;
      }
    } else {
      throw new Error('pptxDecodePng: bad row filter type ' + filterType);
    }
    previous = row;
  }

  var rgba = new Uint8ClampedArray(width * height * 4);
  var pixelCount = width * height;
  var p;
  if (colourType === 6) {
    rgba.set(filtered);
  } else if (colourType === 2) {
    for (p = 0; p < pixelCount; p++) {
      rgba[p * 4] = filtered[p * 3];
      rgba[p * 4 + 1] = filtered[p * 3 + 1];
      rgba[p * 4 + 2] = filtered[p * 3 + 2];
      rgba[p * 4 + 3] = 255;
    }
  } else if (colourType === 0) {
    for (p = 0; p < pixelCount; p++) {
      var grey = filtered[p];
      rgba[p * 4] = grey;
      rgba[p * 4 + 1] = grey;
      rgba[p * 4 + 2] = grey;
      rgba[p * 4 + 3] = 255;
    }
  } else if (colourType === 4) {
    for (p = 0; p < pixelCount; p++) {
      var greyAlpha = filtered[p * 2];
      rgba[p * 4] = greyAlpha;
      rgba[p * 4 + 1] = greyAlpha;
      rgba[p * 4 + 2] = greyAlpha;
      rgba[p * 4 + 3] = filtered[p * 2 + 1];
    }
  } else {
    /* colourType === 3, palette; tRNS gives per-entry alpha */
    for (p = 0; p < pixelCount; p++) {
      var index = filtered[p];
      rgba[p * 4] = palette[index * 3];
      rgba[p * 4 + 1] = palette[index * 3 + 1];
      rgba[p * 4 + 2] = palette[index * 3 + 2];
      rgba[p * 4 + 3] = transparency !== null && index < transparency.length ? transparency[index] : 255;
    }
  }
  return { width: width, height: height, data: rgba };
}

/* --------------------------------------------------------------------- ZIP */

function pptxZipEntries(files) {
  var entries = [];
  if (Array.isArray(files)) {
    for (var i = 0; i < files.length; i++) entries.push(files[i]);
  } else if (files instanceof Map) {
    files.forEach(function (data, name) {
      entries.push({ name: name, data: data });
    });
  } else if (files && typeof files === 'object') {
    Object.keys(files).forEach(function (name) {
      entries.push({ name: name, data: files[name] });
    });
  } else {
    throw new TypeError('pptxZip: expected an array of {name, data}, a Map or a plain object');
  }
  if (entries.length === 0) throw new TypeError('pptxZip: nothing to pack');
  if (entries.length > 0xffff) throw new RangeError('pptxZip: too many entries for a non-ZIP64 archive');
  var out = [];
  for (var j = 0; j < entries.length; j++) {
    var entry = entries[j];
    if (!entry || typeof entry.name !== 'string' || entry.name === '') {
      throw new TypeError('pptxZip: entry ' + j + ' needs a non-empty name');
    }
    var nameBytes = pptxUtf8(entry.name);
    if (nameBytes.length > 0xffff) throw new RangeError('pptxZip: entry name is too long: ' + entry.name);
    var data;
    if (typeof entry.data === 'string') {
      data = pptxUtf8(entry.data);
    } else {
      data = pptxToBytes(entry.data);
    }
    var flags = 0;
    for (var b = 0; b < nameBytes.length; b++) {
      if (nameBytes[b] > 0x7f) {
        flags = 0x0800; /* general purpose bit 11: UTF-8 names */
        break;
      }
    }
    out.push({ name: entry.name, nameBytes: nameBytes, data: data, flags: flags });
  }
  return out;
}

/**
 * Build a ZIP archive (local headers + central directory) with fixed DOS
 * timestamps, CRC-32 and no ZIP64. Deflate-raw when available, otherwise store;
 * the smaller of the two is always used, consistently in both headers.
 */
async function pptxZip(files) {
  var entries = pptxZipEntries(files);
  var prepared = [];
  var localSize = 0;
  var centralSize = 0;
  var index;
  for (index = 0; index < entries.length; index++) {
    var entry = entries[index];
    var data = entry.data;
    var crc = pptxCrc32(data, 0, data.length);
    var method = 0;
    var payload = data;
    if (data.length > 0) {
      var compressed = null;
      try {
        compressed = await pptxDeflateRaw(data);
      } catch (error) {
        compressed = null;
      }
      if (compressed !== null && compressed.length < data.length) {
        method = 8;
        payload = compressed;
      }
    }
    prepared.push({
      name: entry.name,
      nameBytes: entry.nameBytes,
      flags: entry.flags,
      crc: crc,
      method: method,
      compressedSize: payload.length,
      size: data.length,
      payload: payload,
      offset: localSize
    });
    localSize += 30 + entry.nameBytes.length + payload.length;
    centralSize += 46 + entry.nameBytes.length;
  }
  var totalSize = localSize + centralSize + 22;
  var buffer = new ArrayBuffer(totalSize);
  var out = new Uint8Array(buffer);
  var view = new DataView(buffer);
  var position = 0;

  for (index = 0; index < prepared.length; index++) {
    var local = prepared[index];
    view.setUint32(position, 0x04034b50, true); /* local file header signature */
    view.setUint16(position + 4, 20, true); /* version needed: 2.0 */
    view.setUint16(position + 6, local.flags, true); /* flags: UTF-8 off for ASCII names */
    view.setUint16(position + 8, local.method, true);
    view.setUint16(position + 10, pptxFixedDosTime, true);
    view.setUint16(position + 12, pptxFixedDosDate, true);
    view.setUint32(position + 14, local.crc, true);
    view.setUint32(position + 18, local.compressedSize, true);
    view.setUint32(position + 22, local.size, true);
    view.setUint16(position + 26, local.nameBytes.length, true);
    view.setUint16(position + 28, 0, true); /* extra field length */
    out.set(local.nameBytes, position + 30);
    out.set(local.payload, position + 30 + local.nameBytes.length);
    position += 30 + local.nameBytes.length + local.payload.length;
  }

  var centralStart = position;
  for (index = 0; index < prepared.length; index++) {
    var central = prepared[index];
    view.setUint32(position, 0x02014b50, true); /* central directory signature */
    view.setUint16(position + 4, 0x0014, true); /* version made by: 2.0, MS-DOS */
    view.setUint16(position + 6, 20, true); /* version needed: 2.0 */
    view.setUint16(position + 8, central.flags, true); /* flags: must match local header */
    view.setUint16(position + 10, central.method, true);
    view.setUint16(position + 12, pptxFixedDosTime, true);
    view.setUint16(position + 14, pptxFixedDosDate, true);
    view.setUint32(position + 16, central.crc, true);
    view.setUint32(position + 20, central.compressedSize, true);
    view.setUint32(position + 24, central.size, true);
    view.setUint16(position + 28, central.nameBytes.length, true);
    view.setUint16(position + 30, 0, true); /* extra length */
    view.setUint16(position + 32, 0, true); /* comment length */
    view.setUint16(position + 34, 0, true); /* disk number start */
    view.setUint16(position + 36, 0, true); /* internal attributes */
    view.setUint32(position + 38, 0, true); /* external attributes */
    view.setUint32(position + 42, central.offset, true);
    out.set(central.nameBytes, position + 46);
    position += 46 + central.nameBytes.length;
  }

  view.setUint32(position, 0x06054b50, true); /* end of central directory */
  view.setUint16(position + 4, 0, true);
  view.setUint16(position + 6, 0, true);
  view.setUint16(position + 8, prepared.length, true);
  view.setUint16(position + 10, prepared.length, true);
  view.setUint32(position + 12, centralSize, true); /* central directory size */
  view.setUint32(position + 16, centralStart, true); /* central directory offset */
  view.setUint16(position + 20, 0, true); /* comment length */
  return out;
}

/** Parse a ZIP archive into an ordered Map of name -> Uint8Array. */
async function pptxUnzip(bytes) {
  var buf = pptxToBytes(bytes);
  if (buf.length < 22) throw new Error('pptxUnzip: file is too short to be a ZIP archive');
  var view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  var eocd = -1;
  var lowest = Math.max(0, buf.length - 22 - 0xffff);
  for (var scan = buf.length - 22; scan >= lowest; scan--) {
    if (view.getUint32(scan, true) === 0x06054b50) {
      eocd = scan;
      break;
    }
  }
  if (eocd < 0) throw new Error('pptxUnzip: end of central directory record not found');
  var entryCount = view.getUint16(eocd + 10, true);
  if (entryCount === 0xffff) throw new Error('pptxUnzip: ZIP64 archives are not supported');
  var centralOffset = view.getUint32(eocd + 16, true);
  if (centralOffset > buf.length) throw new Error('pptxUnzip: central directory offset is out of range');

  var result = new Map();
  var position = centralOffset;
  for (var index = 0; index < entryCount; index++) {
    if (position + 46 > buf.length) throw new Error('pptxUnzip: truncated central directory');
    if (view.getUint32(position, true) !== 0x02014b50) {
      throw new Error('pptxUnzip: bad central directory signature at ' + position);
    }
    var flags = view.getUint16(position + 8, true);
    var method = view.getUint16(position + 10, true);
    var crc = view.getUint32(position + 16, true);
    var compressedSize = view.getUint32(position + 20, true);
    var uncompressedSize = view.getUint32(position + 24, true);
    var nameLength = view.getUint16(position + 28, true);
    var extraLength = view.getUint16(position + 30, true);
    var commentLength = view.getUint16(position + 32, true);
    var localOffset = view.getUint32(position + 42, true);
    var name = pptxUtf8Decode(buf.subarray(position + 46, position + 46 + nameLength));
    position += 46 + nameLength + extraLength + commentLength;
    if (flags & 0x0001) throw new Error('pptxUnzip: encrypted entries are not supported (' + name + ')');
    if (name.charAt(name.length - 1) === '/') continue; /* directory entry */
    if (localOffset + 30 > buf.length) throw new Error('pptxUnzip: local header out of range for ' + name);
    if (view.getUint32(localOffset, true) !== 0x04034b50) {
      throw new Error('pptxUnzip: bad local header signature for ' + name);
    }
    var localNameLength = view.getUint16(localOffset + 26, true);
    var localExtraLength = view.getUint16(localOffset + 28, true);
    var dataStart = localOffset + 30 + localNameLength + localExtraLength;
    if (dataStart + compressedSize > buf.length) {
      throw new Error('pptxUnzip: compressed data out of range for ' + name);
    }
    var raw = buf.subarray(dataStart, dataStart + compressedSize);
    var data;
    if (method === 0) {
      data = raw.slice();
    } else if (method === 8) {
      data = await pptxInflateRaw(raw);
    } else {
      throw new Error('pptxUnzip: unsupported compression method ' + method + ' for ' + name);
    }
    if (data.length !== uncompressedSize) {
      throw new Error('pptxUnzip: size mismatch for ' + name + ' (' + data.length + ' != ' + uncompressedSize + ')');
    }
    if (pptxCrc32(data, 0, data.length) !== crc) {
      throw new Error('pptxUnzip: CRC-32 mismatch for ' + name);
    }
    result.set(name, data);
  }
  return result;
}

/* ------------------------------------------------- PresentationML builders */

/** Shared <p:spTree> opening: group shape + group transform. */
function pptxSpTreeHead() {
  return (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  );
}

function pptxXfrm(rect) {
  return (
    '<a:xfrm><a:off x="' + rect.x + '" y="' + rect.y + '"/>' +
    '<a:ext cx="' + rect.w + '" cy="' + rect.h + '"/></a:xfrm>'
  );
}

/** `k:'r'`: filled rectangle, no outline (CONTRACT 5 mapping). */
function pptxRectShapeXml(id, name, colourHex, rect) {
  return (
    '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + pptxXmlEscape(name) + '"/>' +
    '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr>' + pptxXfrm(rect) +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="' + colourHex + '"/></a:solidFill>' +
    '<a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
  );
}

/** `k:'i'`: picture exactly filling the shape box (no crop, stretch fill). */
function pptxPictureShapeXml(id, name, relId, rect) {
  return (
    '<p:pic><p:nvPicPr><p:cNvPr id="' + id + '" name="' + pptxXmlEscape(name) + '"/>' +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="' + relId + '"/>' +
    '<a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr>' + pptxXfrm(rect) +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
  );
}

/** Font size in hundredths of a point, clamped to a sane PresentationML range. */
function pptxTextSize(sz) {
  var hundredths = Math.round(pptxNumber(sz, 12) * 100);
  if (!(hundredths > 0)) hundredths = 1200;
  if (hundredths < 100) hundredths = 100;
  if (hundredths > 400000) hundredths = 400000;
  return hundredths;
}

function pptxTextFont(font) {
  if (typeof font !== 'string') return 'Arial';
  var text = font.trim();
  return text === '' ? 'Arial' : text.slice(0, 64);
}

function pptxTextLines(text) {
  if (typeof text !== 'string' || text === '') return [''];
  var lines = text.replace(/\r\n?/g, '\n').split('\n');
  if (lines.length > 32) lines = lines.slice(0, 32);
  return lines;
}

/**
 * `k:'t'`: editable text box with zero insets, no wrap, top anchored; the box is
 * filled with the erase colour `c` so it reproduces the flat region behind the
 * text. Text ink gets a contrasting colour to stay readable.
 */
function pptxTextShapeXml(id, name, shape, rect) {
  var fill = pptxColorHex(shape.c, 'ffffff');
  var ink = pptxContrastTextColor(fill);
  var size = pptxTextSize(shape.sz);
  var bold = shape.b ? '1' : '0';
  var font = pptxXmlEscape(pptxTextFont(shape.f));
  var alignment = shape.al === 'ctr' ? 'ctr' : shape.al === 'r' ? 'r' : 'l';
  var character =
    '<a:solidFill><a:srgbClr val="' + ink + '"/></a:solidFill>' +
    '<a:latin typeface="' + font + '"/>';
  var runProperties =
    '<a:rPr lang="zh-CN" sz="' + size + '" b="' + bold + '" dirty="0">' + character + '</a:rPr>';
  var definitionProperties =
    '<a:defRPr lang="zh-CN" sz="' + size + '" b="' + bold + '" dirty="0">' + character + '</a:defRPr>';

  var paragraphs = '';
  var lines = pptxTextLines(shape.s);
  for (var i = 0; i < lines.length; i++) {
    paragraphs +=
      '<a:p><a:pPr algn="' + alignment + '">' + definitionProperties + '</a:pPr>';
    if (lines[i] !== '') {
      paragraphs += '<a:r>' + runProperties + '<a:t>' + pptxXmlEscape(lines[i]) + '</a:t></a:r>';
    }
    paragraphs += '<a:endParaRPr lang="zh-CN" sz="' + size + '" dirty="0"/></a:p>';
  }

  return (
    '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + pptxXmlEscape(name) + '"/>' +
    '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr>' + pptxXfrm(rect) +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="' + fill + '"/></a:solidFill>' +
    '<a:ln><a:noFill/></a:ln></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">' +
    '<a:noAutofit/></a:bodyPr><a:lstStyle/>' + paragraphs + '</p:txBody></p:sp>'
  );
}

function pptxShapeName(kind, id) {
  if (kind === 'i') return '图片 ' + id;
  if (kind === 't') return '文本 ' + id;
  return '矩形 ' + id;
}

/** Relationships part XML from an ordered [{id, type, target}] list. */
function pptxRelsXml(rels) {
  var parts = [pptxXmlDeclaration, '<Relationships xmlns="' + pptxNsPackageRelationships + '">'];
  for (var i = 0; i < rels.length; i++) {
    parts.push(
      '<Relationship Id="' + rels[i].id + '" Type="' + rels[i].type + '" Target="' + pptxXmlEscape(rels[i].target) + '"/>'
    );
  }
  parts.push('</Relationships>');
  return parts.join('');
}

function pptxContentTypesXml(plan) {
  var parts = [pptxXmlDeclaration, '<Types xmlns="' + pptxNsContentTypes + '">'];
  parts.push('<Default Extension="rels" ContentType="' + pptxContentType.relationships + '"/>');
  parts.push('<Default Extension="xml" ContentType="' + pptxContentType.xml + '"/>');
  parts.push('<Default Extension="png" ContentType="' + pptxContentType.png + '"/>');
  parts.push('<Override PartName="/ppt/presentation.xml" ContentType="' + pptxContentType.presentation + '"/>');
  parts.push('<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="' + pptxContentType.slideMaster + '"/>');
  parts.push('<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="' + pptxContentType.slideLayout + '"/>');
  parts.push('<Override PartName="/ppt/theme/theme1.xml" ContentType="' + pptxContentType.theme + '"/>');
  parts.push('<Override PartName="/docProps/core.xml" ContentType="' + pptxContentType.coreProperties + '"/>');
  parts.push('<Override PartName="/docProps/app.xml" ContentType="' + pptxContentType.extendedProperties + '"/>');
  for (var i = 0; i < plan.slides.length; i++) {
    parts.push(
      '<Override PartName="/ppt/slides/slide' + (i + 1) + '.xml" ContentType="' + pptxContentType.slide + '"/>'
    );
  }
  parts.push('</Types>');
  return parts.join('');
}

function pptxRootRelsXml() {
  return pptxRelsXml([
    { id: 'rId1', type: pptxRelType.officeDocument, target: 'ppt/presentation.xml' },
    { id: 'rId2', type: pptxRelType.coreProperties, target: 'docProps/core.xml' },
    { id: 'rId3', type: pptxRelType.extendedProperties, target: 'docProps/app.xml' }
  ]);
}

function pptxPresentationXml(plan) {
  var parts = [pptxXmlDeclaration];
  parts.push(
    '<p:presentation xmlns:a="' + pptxNsDrawing + '" xmlns:r="' + pptxNsRelationships +
    '" xmlns:p="' + pptxNsPresentation + '" saveSubsetFonts="1">'
  );
  parts.push('<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>');
  parts.push('<p:sldIdLst>');
  for (var i = 0; i < plan.slides.length; i++) {
    parts.push('<p:sldId id="' + (256 + i) + '" r:id="rId' + (3 + i) + '"/>');
  }
  parts.push('</p:sldIdLst>');
  parts.push('<p:sldSz cx="' + plan.cx + '" cy="' + plan.cy + '" type="custom"/>');
  parts.push('<p:notesSz cx="6858000" cy="9144000"/>');
  parts.push(
    '<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle>'
  );
  parts.push('</p:presentation>');
  return parts.join('');
}

function pptxPresentationRelsXml(plan) {
  var rels = [
    { id: 'rId1', type: pptxRelType.slideMaster, target: 'slideMasters/slideMaster1.xml' },
    { id: 'rId2', type: pptxRelType.theme, target: 'theme/theme1.xml' }
  ];
  for (var i = 0; i < plan.slides.length; i++) {
    rels.push({ id: 'rId' + (3 + i), type: pptxRelType.slide, target: 'slides/slide' + (i + 1) + '.xml' });
  }
  return pptxRelsXml(rels);
}

function pptxMasterXml() {
  var parts = [pptxXmlDeclaration];
  parts.push(
    '<p:sldMaster xmlns:a="' + pptxNsDrawing + '" xmlns:r="' + pptxNsRelationships +
    '" xmlns:p="' + pptxNsPresentation + '">'
  );
  parts.push(
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
    '<a:effectLst/></p:bgPr></p:bg>' + pptxSpTreeHead() + '</p:spTree></p:cSld>'
  );
  parts.push(
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
    ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
    ' hlink="hlink" folHlink="folHlink"/>'
  );
  parts.push('<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>');
  parts.push(
    '<p:txStyles>' +
    '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle>' +
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle>' +
    '<p:otherStyle><a:defPPr><a:defRPr lang="zh-CN" sz="1800"/></a:defPPr>' +
    '<a:lvl1pPr><a:defRPr lang="zh-CN" sz="1800"/></a:lvl1pPr></p:otherStyle>' +
    '</p:txStyles>'
  );
  parts.push('</p:sldMaster>');
  return parts.join('');
}

function pptxMasterRelsXml() {
  return pptxRelsXml([
    { id: 'rId1', type: pptxRelType.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
    { id: 'rId2', type: pptxRelType.theme, target: '../theme/theme1.xml' }
  ]);
}

function pptxLayoutXml() {
  var parts = [pptxXmlDeclaration];
  parts.push(
    '<p:sldLayout xmlns:a="' + pptxNsDrawing + '" xmlns:r="' + pptxNsRelationships +
    '" xmlns:p="' + pptxNsPresentation + '" type="blank" preserve="1">'
  );
  parts.push('<p:cSld name="Blank">' + pptxSpTreeHead() + '</p:spTree></p:cSld>');
  parts.push('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>');
  parts.push('</p:sldLayout>');
  return parts.join('');
}

function pptxLayoutRelsXml() {
  return pptxRelsXml([
    { id: 'rId1', type: pptxRelType.slideMaster, target: '../slideMasters/slideMaster1.xml' }
  ]);
}

function pptxThemeXml() {
  var fills = '';
  var lines = '';
  var effects = '';
  var backgroundFills = '';
  for (var i = 0; i < 3; i++) {
    fills += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    lines +=
      '<a:ln w="' + (9525 * (i + 1)) + '" cap="flat" cmpd="sng" algn="ctr">' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      '<a:prstDash val="solid"/></a:ln>';
    effects += '<a:effectStyle><a:effectLst/></a:effectStyle>';
    backgroundFills += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  }
  var parts = [pptxXmlDeclaration];
  parts.push('<a:theme xmlns:a="' + pptxNsDrawing + '" name="dsh-pixeltrace">');
  parts.push('<a:themeElements>');
  parts.push(
    '<a:clrScheme name="dsh-pixeltrace">' +
    '<a:dk1><a:srgbClr val="000000"/></a:dk1>' +
    '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="1F2933"/></a:dk2>' +
    '<a:lt2><a:srgbClr val="F3F5F7"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="1F6FEB"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="E3622F"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="2E9E5B"/></a:accent3>' +
    '<a:accent4><a:srgbClr val="8A5CF6"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="0E9AA7"/></a:accent5>' +
    '<a:accent6><a:srgbClr val="D9A521"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    '</a:clrScheme>'
  );
  parts.push(
    '<a:fontScheme name="dsh-pixeltrace">' +
    '<a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>'
  );
  parts.push(
    '<a:fmtScheme name="dsh-pixeltrace">' +
    '<a:fillStyleLst>' + fills + '</a:fillStyleLst>' +
    '<a:lnStyleLst>' + lines + '</a:lnStyleLst>' +
    '<a:effectStyleLst>' + effects + '</a:effectStyleLst>' +
    '<a:bgFillStyleLst>' + backgroundFills + '</a:bgFillStyleLst>' +
    '</a:fmtScheme>'
  );
  parts.push('</a:themeElements></a:theme>');
  return parts.join('');
}

function pptxCorePropertiesXml(options) {
  return (
    pptxXmlDeclaration +
    '<cp:coreProperties xmlns:cp="' + pptxNsCoreProperties + '" xmlns:dc="' + pptxNsDc +
    '" xmlns:dcterms="' + pptxNsDcTerms + '" xmlns:dcmitype="' + pptxNsDcMiType +
    '" xmlns:xsi="' + pptxNsXsi + '">' +
    '<dc:title>' + pptxXmlEscape(options.title) + '</dc:title>' +
    '<dc:subject>' + pptxXmlEscape(options.subject) + '</dc:subject>' +
    '<dc:creator>' + pptxXmlEscape(options.creator) + '</dc:creator>' +
    '<cp:lastModifiedBy>' + pptxXmlEscape(options.creator) + '</cp:lastModifiedBy>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">' + pptxFixedTimestamp + '</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">' + pptxFixedTimestamp + '</dcterms:modified>' +
    '</cp:coreProperties>'
  );
}

function pptxAppPropertiesXml(plan) {
  return (
    pptxXmlDeclaration +
    '<Properties xmlns="' + pptxNsExtendedProperties + '" xmlns:vt="' + pptxNsVt + '">' +
    '<Application>dsh-pixeltrace</Application>' +
    '<PresentationFormat>Custom</PresentationFormat>' +
    '<Slides>' + plan.slides.length + '</Slides>' +
    '<Notes>0</Notes>' +
    '<HiddenSlides>0</HiddenSlides>' +
    '<ScaleCrop>false</ScaleCrop>' +
    '<AppVersion>16.0000</AppVersion>' +
    '</Properties>'
  );
}

function pptxSlideXml(slide) {
  var parts = [pptxXmlDeclaration];
  parts.push(
    '<p:sld xmlns:a="' + pptxNsDrawing + '" xmlns:r="' + pptxNsRelationships +
    '" xmlns:p="' + pptxNsPresentation + '">'
  );
  parts.push('<p:cSld name="' + pptxXmlEscape(slide.name) + '">');
  parts.push(
    '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="' + slide.background + '"/></a:solidFill>' +
    '<a:effectLst/></p:bgPr></p:bg>'
  );
  parts.push(pptxSpTreeHead());
  parts.push(slide.shapesXml);
  parts.push('</p:spTree></p:cSld>');
  parts.push('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>');
  parts.push('</p:sld>');
  return parts.join('');
}

/* ------------------------------------------------------------- deck builder */

function pptxBuildOptions(raw) {
  var input = raw && typeof raw === 'object' ? raw : {};
  var emuPerPx = pptxNumber(input.emuPerPx, pptxDefaultEmuPerPx);
  if (!(emuPerPx > 0)) emuPerPx = pptxDefaultEmuPerPx;
  var maxSlideEmu = pptxNumber(input.maxSlideEmu, pptxDefaultMaxSlideEmu);
  if (!(maxSlideEmu > 0)) maxSlideEmu = pptxDefaultMaxSlideEmu;
  return {
    title: pptxOptionText(input.title, 'PixelTrace'),
    subject: pptxOptionText(input.subject, ''),
    creator: pptxOptionText(input.creator, 'PixelTrace'),
    emuPerPx: emuPerPx,
    maxSlideEmu: maxSlideEmu,
    layout16x9: input.layout16x9 === true
  };
}

/** Validate/normalise the slide list; empty input becomes one blank white slide. */
function pptxNormaliseSlides(slides) {
  if (slides !== undefined && slides !== null && !Array.isArray(slides)) {
    throw new TypeError('pptxBuild: slides must be an array of {scene}');
  }
  var input = Array.isArray(slides) ? slides : [];
  if (input.length === 0) {
    input = [{ scene: { width: 1280, height: 720, background: '#ffffff', shapes: [] } }];
  }
  var list = [];
  for (var i = 0; i < input.length; i++) {
    var item = input[i] && typeof input[i] === 'object' ? input[i] : {};
    var scene = item.scene;
    if (!scene || typeof scene !== 'object') {
      throw new TypeError('pptxBuild: slide ' + (i + 1) + ' has no scene');
    }
    var width = Math.round(pptxNumber(scene.width, 0));
    var height = Math.round(pptxNumber(scene.height, 0));
    if (!(width > 0) || !(height > 0)) {
      throw new TypeError('pptxBuild: slide ' + (i + 1) + ' needs a positive integer scene width/height');
    }
    var name = item.name;
    if (typeof name !== 'string' || name === '') name = 'slide-' + (i + 1);
    list.push({
      scene: scene,
      name: name.slice(0, 200),
      background: pptxColorHex(item.background !== undefined && item.background !== null ? item.background : scene.background, 'ffffff'),
      width: width,
      height: height
    });
  }
  return list;
}

/** Shape box in source pixels -> EMU box; null for degenerate shapes. */
function pptxShapeRect(shape, toEmu) {
  var rect = {
    x: toEmu(pptxNumber(shape.x, 0)),
    y: toEmu(pptxNumber(shape.y, 0)),
    w: toEmu(pptxNumber(shape.w, 0)),
    h: toEmu(pptxNumber(shape.h, 0))
  };
  if (rect.w <= 0 || rect.h <= 0) return null;
  return rect;
}

/**
 * Return the index of the media part for this crop, encoding it on first use and
 * de-duplicating identical crops by byte equality (CONTRACT 5).
 */
async function pptxMediaIndex(store, px) {
  if (!px || typeof px !== 'object' || !px.data || typeof px.data.length !== 'number') {
    throw new TypeError('pptxBuild: image shape needs px {data, width, height}');
  }
  var width = Math.round(pptxNumber(px.width, 0));
  var height = Math.round(pptxNumber(px.height, 0));
  if (!(width > 0) || !(height > 0)) {
    throw new TypeError('pptxBuild: image shape px.width/px.height must be positive');
  }
  var needed = width * height * 4;
  if (px.data.length < needed) {
    throw new TypeError(
      'pptxBuild: image shape px.data is too short (' + px.data.length + ' < ' + needed + ' bytes)'
    );
  }
  var pixels;
  if (typeof px.data.byteLength === 'number' && px.data.buffer) {
    pixels = new Uint8Array(px.data.buffer, px.data.byteOffset || 0, needed);
  } else {
    /* tolerate a plain array of byte values */
    pixels = pptxToBytes(px.data).subarray(0, needed);
  }
  var key = width + 'x' + height + ':' + pptxHashBytes(pixels);
  var bucket = store.buckets.get(key);
  if (bucket) {
    for (var candidate = 0; candidate < bucket.length; candidate++) {
      if (pptxBytesEqual(store.items[bucket[candidate]].pixels, pixels)) return bucket[candidate];
    }
  }
  var png = await pptxEncodePng(pixels, width, height);
  var index = store.items.length;
  store.items.push({
    name: 'image' + (index + 1) + '.png',
    pixels: pixels,
    png: png,
    width: width,
    height: height
  });
  if (!bucket) {
    bucket = [];
    store.buckets.set(key, bucket);
  }
  bucket.push(index);
  return index;
}

/**
 * Build the layout plan: slide geometry, shape XML and relationships. Media is
 * encoded here, before XML, so relationship ids can reference image parts.
 */
async function pptxPlanDeck(options) {
  var store = { items: [], buckets: new Map() };
  var slides = [];
  var nextShapeId = 2;
  var prepared = options.slides;
  var maxWidth = 1;
  var maxHeight = 1;
  var index;

  for (index = 0; index < prepared.length; index++) {
    if (prepared[index].width > maxWidth) maxWidth = prepared[index].width;
    if (prepared[index].height > maxHeight) maxHeight = prepared[index].height;
  }
  var edge = Math.max(maxWidth, maxHeight);
  var scale = Math.min(1, options.maxSlideEmu / edge / pptxEmuPerPxCap);
  var emuPerPx = options.emuPerPx;
  function toEmu(pixelValue) {
    return Math.round(pixelValue * emuPerPx * scale);
  }
  var cx = Math.max(1, toEmu(maxWidth));
  var cy = Math.max(1, toEmu(maxHeight));
  if (options.layout16x9) {
    if (cx * 9 < cy * 16) cx = Math.max(1, Math.round((cy * 16) / 9));
    else if (cx * 9 > cy * 16) cy = Math.max(1, Math.round((cx * 9) / 16));
  }

  for (index = 0; index < prepared.length; index++) {
    var slide = prepared[index];
    var scene = slide.scene;
    var shapes = Array.isArray(scene.shapes) ? scene.shapes : [];
    var rels = [{ id: 'rId1', type: pptxRelType.slideLayout, target: '../slideLayouts/slideLayout1.xml' }];
    var relIdByMedia = new Map();
    var shapeXml = '';
    for (var shapeIndex = 0; shapeIndex < shapes.length; shapeIndex++) {
      var shape = shapes[shapeIndex];
      if (!shape || typeof shape !== 'object') continue;
      var kind = shape.k;
      if (kind !== 'r' && kind !== 'i' && kind !== 't') continue;
      var rect = pptxShapeRect(shape, toEmu);
      if (rect === null) continue;
      var id = nextShapeId++;
      if (kind === 'r') {
        shapeXml += pptxRectShapeXml(id, pptxShapeName('r', id), pptxColorHex(shape.c, 'ffffff'), rect);
      } else if (kind === 't') {
        shapeXml += pptxTextShapeXml(id, pptxShapeName('t', id), shape, rect);
      } else {
        var mediaIndex = await pptxMediaIndex(store, shape.px);
        var relId = relIdByMedia.get(mediaIndex);
        if (relId === undefined) {
          relId = 'rId' + (rels.length + 1);
          rels.push({
            id: relId,
            type: pptxRelType.image,
            target: '../media/' + store.items[mediaIndex].name
          });
          relIdByMedia.set(mediaIndex, relId);
        }
        shapeXml += pptxPictureShapeXml(id, pptxShapeName('i', id), relId, rect);
      }
    }
    slides.push({
      name: slide.name,
      background: slide.background,
      width: slide.width,
      height: slide.height,
      shapesXml: shapeXml,
      rels: rels
    });
  }

  return {
    slides: slides,
    media: store.items,
    cx: cx,
    cy: cy,
    scale: scale,
    emuPerPx: emuPerPx
  };
}

/**
 * Build a complete .pptx package from scenes.
 * @param {Array<{scene:object, name?:string, background?:string}>} slides
 * @param {{title?:string, subject?:string, creator?:string, emuPerPx?:number,
 *          maxSlideEmu?:number, layout16x9?:boolean}} [options]
 * @returns {Promise<Uint8Array>}
 */
async function pptxBuild(slides, options) {
  var settings = pptxBuildOptions(options);
  var normalised = pptxNormaliseSlides(slides);
  var plan = await pptxPlanDeck({
    slides: normalised,
    emuPerPx: settings.emuPerPx,
    maxSlideEmu: settings.maxSlideEmu,
    layout16x9: settings.layout16x9
  });
  var files = [];
  files.push({ name: '[Content_Types].xml', data: pptxUtf8(pptxContentTypesXml(plan)) });
  files.push({ name: '_rels/.rels', data: pptxUtf8(pptxRootRelsXml()) });
  files.push({ name: 'docProps/core.xml', data: pptxUtf8(pptxCorePropertiesXml(settings)) });
  files.push({ name: 'docProps/app.xml', data: pptxUtf8(pptxAppPropertiesXml(plan)) });
  files.push({ name: 'ppt/presentation.xml', data: pptxUtf8(pptxPresentationXml(plan)) });
  files.push({ name: 'ppt/_rels/presentation.xml.rels', data: pptxUtf8(pptxPresentationRelsXml(plan)) });
  files.push({ name: 'ppt/theme/theme1.xml', data: pptxUtf8(pptxThemeXml()) });
  files.push({ name: 'ppt/slideMasters/slideMaster1.xml', data: pptxUtf8(pptxMasterXml()) });
  files.push({ name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: pptxUtf8(pptxMasterRelsXml()) });
  files.push({ name: 'ppt/slideLayouts/slideLayout1.xml', data: pptxUtf8(pptxLayoutXml()) });
  files.push({ name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: pptxUtf8(pptxLayoutRelsXml()) });
  var index;
  for (index = 0; index < plan.slides.length; index++) {
    files.push({
      name: 'ppt/slides/slide' + (index + 1) + '.xml',
      data: pptxUtf8(pptxSlideXml(plan.slides[index]))
    });
    files.push({
      name: 'ppt/slides/_rels/slide' + (index + 1) + '.xml.rels',
      data: pptxUtf8(pptxRelsXml(plan.slides[index].rels))
    });
  }
  for (index = 0; index < plan.media.length; index++) {
    files.push({ name: 'ppt/media/' + plan.media[index].name, data: plan.media[index].png });
  }
  return await pptxZip(files);
}

/* ==================== src/ui.frag.js ==================== */
/* PixelTrace panel: drop an image, rebuild it with PowerPoint objects, export .pptx.
 *
 * Plain fragment: no import/export, no top-level await. React arrives from the
 * browser module table as `React`/`h` in the surrounding factory closure.
 * Every top-level name starts with `ui`.
 */

const uiPanelId = 'pixeltrace';
const uiSidebarOrder = 32;
const uiStyleId = 'dsh-pixeltrace-style';

/** zh -> en. Untranslated keys fall back to the Chinese key itself. */
const uiStrings = {
  '像素复刻 PixelTrace': 'PixelTrace',
  '把一张图还原成 PPT 里可编辑的形状 · 像素级一致':
    'Rebuild any image as editable PowerPoint objects, pixel for pixel',
  '拖入图片，或点击选择': 'Drop an image here, or click to choose',
  'PNG · JPEG · WebP · GIF · 可多选（每张一页）· 也可以直接 Ctrl+V 粘贴':
    'PNG · JPEG · WebP · GIF · multiple files (one slide each) · or Ctrl+V to paste',
  '选择图片': 'Choose images',
  '清空': 'Clear',
  '输出参数': 'Output',
  '输出模式': 'Mode',
  '智能分层': 'Smart layers',
  '纯矢量': 'Vector only',
  '整图保真': 'Whole picture',
  '平面区域变成可编辑色块，细节区域变成精确裁剪的图片，整页仍逐像素一致':
    'Flat areas become editable rectangles, detail areas become exactly cropped pictures; the slide stays pixel-identical',
  '全部变成可编辑矩形：每个色块都能单独改，但形状数量最多':
    'Everything becomes editable rectangles: every block is editable, but the shape count is highest',
  '每页只放一张图片对象：100% 一致，但不能分层编辑':
    'One picture object per slide: 100% identical, but nothing to edit layer by layer',
  '颜色量化': 'Colour quantisation',
  '0 = 保留原始颜色': '0 = keep original colours',
  '颜色容差': 'Colour tolerance',
  '合并相邻矩形时允许的单通道色差，0 = 完全精确': 'Per-channel delta allowed when merging rectangles; 0 = exact',
  '形状预算/页': 'Shape budget / slide',
  '0 = 不限制；超出预算时自动降级并在统计里说明': '0 = unlimited; over budget it degrades and says so in the notes',
  '细节区域转为精确裁剪图': 'Turn detail areas into exact crops',
  '文字识别为可编辑文本框': 'Detect text as editable text boxes',
  '实验性：文字墨迹被擦成背景色，留下一个空文本框，方便你自己输入原文':
    'Experimental: ink is erased to the background colour and an empty text box is left for you to type',
  '透明背景色': 'Transparent background',
  '重新分析': 'Re-analyse',
  '分析中…': 'Analysing…',
  '图片': 'Images',
  '还没有图片': 'No image yet',
  '移除': 'Remove',
  '预览': 'Preview',
  '原图': 'Source',
  '重建': 'Rebuilt',
  '差异': 'Diff',
  '适应': 'Fit',
  '保真统计': 'Fidelity',
  '对象': 'objects',
  '矩形': 'rects',
  '裁剪图': 'crops',
  '文本框': 'texts',
  '颜色': 'colours',
  '完全一致像素': 'Exact pixels',
  '不一致像素': 'Mismatched px',
  '平均误差': 'Mean error',
  '最大误差': 'Max error',
  '耗时': 'Time',
  '备注': 'Notes',
  '导出 PPTX': 'Export PPTX',
  '正在生成…': 'Building…',
  '文件名': 'File name',
  '参数已修改，点「重新分析」刷新预览': 'Settings changed — press Re-analyse to refresh the preview',
  '先添加图片才能导出': 'Add an image first',
  '已导出': 'Exported',
  '已取消': 'Cancelled',
  '出错': 'Error',
  '拖入图片后在这里查看逐像素预览': 'Drop an image, then inspect the pixel-exact preview here',
  '提示：导出的每个色块、裁剪图、文本框都是独立对象，可在 PowerPoint / WPS / Keynote 里单独编辑。':
    'Tip: every rectangle, crop and text box is a separate object you can edit in PowerPoint / WPS / Keynote.',
  '正在读取图片…': 'Reading image…',
  '正在分析…': 'Analysing…',
  '正在生成 PPTX…': 'Building the .pptx…',
  '图片超过 56 英寸幻灯片上限，已等比缩放': 'Image exceeds the 56-inch slide cap and was scaled down',
  '图片过大，已降采样以获得可用速度': 'Very large image downsampled for usable speed',
  '全部精确重建': 'Reconstructed exactly',
  '页': 'slides',
  '张': '',
  '无法读取该图片': 'Could not read this image',
  '该浏览器不支持画布读取像素': 'This browser cannot read canvas pixels',
};

let uiLocaleService = null;

function uiSetLocaleService(service) {
  uiLocaleService = service || null;
}

function uiLanguageOf(locale) {
  const source = locale || uiLocaleService;
  try {
    const snapshot = source && typeof source.getSnapshot === 'function' ? source.getSnapshot() : null;
    const active = snapshot && typeof snapshot === 'object' ? snapshot.active || snapshot.locale || '' : '';
    return String(active || '').toLowerCase().startsWith('en') ? 'en' : 'zh';
  } catch {
    return 'zh';
  }
}

function uiT(zh) {
  if (uiLanguageOf() !== 'en') {
    return zh;
  }
  return Object.prototype.hasOwnProperty.call(uiStrings, zh) ? uiStrings[zh] : zh;
}

function uiUseLanguage(locale) {
  const [, bump] = useState(0);
  useEffect(() => {
    if (!locale || typeof locale.subscribe !== 'function') {
      return undefined;
    }
    let alive = true;
    const update = () => {
      if (alive) {
        bump((n) => n + 1);
      }
    };
    const off = locale.subscribe(update);
    return () => {
      alive = false;
      if (typeof off === 'function') {
        off();
      }
    };
  }, [locale]);
}

const uiCss = `
.pt-root{display:flex;flex-direction:column;gap:14px;height:100%;min-height:0;box-sizing:border-box;padding:18px 22px 22px;overflow:auto;background:var(--dsw-alias-bg-base,transparent);color:var(--dsw-alias-label-primary,#1f2328);font:13px/1.55 "Segoe UI","PingFang SC","Microsoft YaHei",system-ui,sans-serif;}
.pt-head{display:flex;align-items:center;gap:12px;flex:0 0 auto;}
.pt-logo{width:36px;height:36px;border-radius:10px;display:grid;place-items:center;flex:0 0 auto;color:var(--dsw-alias-brand-primary,#3b82f6);background:linear-gradient(135deg,rgba(59,130,246,.22),rgba(168,85,247,.22));border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.12));}
.pt-h1{margin:0;font-size:15px;font-weight:650;letter-spacing:.2px;}
.pt-sub{margin:2px 0 0;font-size:12px;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-drop{flex:0 0 auto;display:flex;align-items:center;gap:14px;padding:14px 16px;border-radius:12px;border:1.5px dashed var(--dsw-alias-border-l2,rgba(0,0,0,.2));background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.06));cursor:pointer;transition:border-color .15s ease,background .15s ease;outline:none;}
.pt-drop:hover,.pt-drop.is-over{border-color:var(--dsw-alias-brand-primary,#3b82f6);background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 9%,transparent);}
.pt-drop-icon{width:34px;height:34px;flex:0 0 auto;display:grid;place-items:center;border-radius:9px;color:var(--dsw-alias-brand-primary,#3b82f6);background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 14%,transparent);}
.pt-drop-title{font-weight:600;font-size:13px;}
.pt-drop-hint{font-size:11.5px;color:var(--dsw-alias-label-secondary,#6b7280);margin-top:2px;}
.pt-drop-spacer{flex:1 1 auto;}
.pt-form{flex:1 1 auto;min-height:0;display:grid;grid-template-columns:290px minmax(0,1fr);gap:16px;align-items:start;overflow:auto;padding-right:2px;}
@media (max-width:880px){.pt-form{grid-template-columns:minmax(0,1fr);}}
.pt-side{display:flex;flex-direction:column;gap:12px;min-width:0;}
.pt-stage{display:flex;flex-direction:column;gap:12px;min-width:0;min-height:0;overflow:auto;}
.pt-card{border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.12));background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.05));border-radius:12px;padding:12px;}
.pt-card-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:10px;}
.pt-card-title{font-size:11.5px;font-weight:650;letter-spacing:.6px;text-transform:uppercase;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-field{display:flex;flex-direction:column;gap:5px;margin-top:10px;}
.pt-field:first-child{margin-top:0;}
.pt-label{font-size:12px;font-weight:550;display:flex;justify-content:space-between;align-items:center;gap:8px;}
.pt-value{font-variant-numeric:tabular-nums;font-size:11.5px;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-hint{font-size:11px;color:var(--dsw-alias-label-secondary,#6b7280);line-height:1.45;}
.pt-seg{display:flex;gap:4px;padding:3px;border-radius:9px;background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.1));}
.pt-seg button{flex:1 1 0;appearance:none;border:0;background:transparent;color:inherit;font:inherit;font-size:12px;padding:6px 4px;border-radius:6px;cursor:pointer;transition:background .12s ease,color .12s ease;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.pt-seg button.is-on{background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-brand-primary,#3b82f6);font-weight:600;box-shadow:0 1px 2px rgba(0,0,0,.12);}
.pt-seg button:hover:not(.is-on){background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 10%,transparent);}
.pt-stage-bar .pt-seg{flex:0 0 auto;}
.pt-stage-bar .pt-seg button{flex:0 0 auto;padding:6px 10px;}
.pt-input,.pt-select{width:100%;box-sizing:border-box;padding:6px 9px;border-radius:8px;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.16));background:var(--dsw-alias-bg-base,#fff);color:inherit;font:inherit;font-size:12.5px;outline:none;}
.pt-input:focus,.pt-select:focus{border-color:var(--dsw-alias-brand-primary,#3b82f6);box-shadow:0 0 0 3px color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 18%,transparent);}
.pt-row{display:flex;align-items:center;gap:8px;}
.pt-check{display:flex;align-items:flex-start;gap:8px;margin-top:10px;cursor:pointer;}
.pt-check input{margin:2px 0 0;accent-color:var(--dsw-alias-brand-primary,#3b82f6);}
.pt-check-text{display:flex;flex-direction:column;gap:2px;}
.pt-check-title{font-size:12px;font-weight:550;}
.pt-swatch{width:34px;height:26px;padding:0;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.16));border-radius:7px;background:transparent;cursor:pointer;}
.pt-btn{appearance:none;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.16));background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.08));color:inherit;font:inherit;font-size:12.5px;font-weight:550;padding:7px 13px;border-radius:9px;cursor:pointer;transition:transform .06s ease,background .12s ease,border-color .12s ease;white-space:nowrap;}
.pt-btn:hover:not(:disabled){background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 12%,transparent);border-color:var(--dsw-alias-brand-primary,#3b82f6);}
.pt-btn:active:not(:disabled){transform:translateY(1px);}
.pt-btn:disabled{opacity:.45;cursor:not-allowed;}
.pt-btn.pt-primary{background:var(--dsw-alias-brand-primary,#3b82f6);border-color:transparent;color:#fff;}
.pt-btn.pt-primary:hover:not(:disabled){background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 86%,#000);}
.pt-btn.pt-ghost{border-color:transparent;background:transparent;color:var(--dsw-alias-label-secondary,#6b7280);padding:5px 8px;}
.pt-btn.pt-ghost:hover:not(:disabled){color:inherit;background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.12));}
.pt-btn.pt-small{font-size:11.5px;padding:4px 9px;border-radius:7px;}
.pt-files{display:flex;flex-direction:column;gap:5px;max-height:200px;overflow:auto;}
.pt-file{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:8px;cursor:pointer;border:1px solid transparent;background:var(--dsw-alias-bg-base,transparent);}
.pt-file.is-on{border-color:var(--dsw-alias-brand-primary,#3b82f6);background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 9%,transparent);}
.pt-file-body{min-width:0;flex:1 1 auto;}
.pt-file-name{display:block;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pt-file-meta{display:block;font-size:10.5px;color:var(--dsw-alias-label-secondary,#6b7280);font-variant-numeric:tabular-nums;}
.pt-x{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary,#6b7280);cursor:pointer;font-size:15px;line-height:1;padding:3px 5px;border-radius:6px;}
.pt-x:hover{color:var(--dsw-alias-state-error-primary,#dc2626);background:color-mix(in srgb,var(--dsw-alias-state-error-primary,#dc2626) 12%,transparent);}
.pt-empty{padding:22px 12px;text-align:center;font-size:12px;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-stage-bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
.pt-stage-title{font-size:12px;font-weight:600;margin-right:auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pt-canvas-wrap{position:relative;border-radius:12px;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.12));background-color:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.08));background-image:linear-gradient(45deg,rgba(127,127,127,.13) 25%,transparent 25%),linear-gradient(-45deg,rgba(127,127,127,.13) 25%,transparent 25%),linear-gradient(45deg,transparent 75%,rgba(127,127,127,.13) 75%),linear-gradient(-45deg,transparent 75%,rgba(127,127,127,.13) 75%);background-size:16px 16px;background-position:0 0,0 8px,8px -8px,-8px 0;overflow:auto;padding:10px;max-height:46vh;}
.pt-canvas{display:block;margin:0 auto;border-radius:4px;box-shadow:0 2px 12px rgba(0,0,0,.18);}
.pt-stats{display:flex;flex-wrap:wrap;gap:6px;}
.pt-stat{display:flex;flex-direction:column;gap:1px;padding:6px 9px;border-radius:10px;background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.07));border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.1));min-width:74px;}
.pt-stat-k{font-size:10.5px;letter-spacing:.4px;text-transform:uppercase;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-stat-v{font-size:14px;font-weight:650;font-variant-numeric:tabular-nums;}
.pt-stat.is-good .pt-stat-v{color:var(--dsw-alias-state-success-primary,#16a34a);}
.pt-stat.is-bad .pt-stat-v{color:var(--dsw-alias-state-error-primary,#dc2626);}
.pt-notes{font-size:11.5px;color:var(--dsw-alias-label-secondary,#6b7280);display:flex;flex-direction:column;gap:2px;margin-top:2px;}
.pt-notes code{font-size:11px;background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.14));border-radius:4px;padding:1px 4px;}
.pt-foot{flex:0 0 auto;display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding-top:4px;border-top:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.1));}
.pt-foot-name{display:flex;align-items:center;gap:7px;}
.pt-foot-name .pt-input{width:190px;}
.pt-msg{font-size:12px;border-radius:8px;padding:7px 11px;display:flex;align-items:center;gap:8px;}
.pt-msg.ok{color:var(--dsw-alias-state-success-primary,#16a34a);background:color-mix(in srgb,var(--dsw-alias-state-success-primary,#16a34a) 12%,transparent);}
.pt-msg.error{color:var(--dsw-alias-state-error-primary,#dc2626);background:color-mix(in srgb,var(--dsw-alias-state-error-primary,#dc2626) 12%,transparent);}
.pt-msg.info{color:var(--dsw-alias-label-secondary,#6b7280);background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.1));}
.pt-warn{font-size:11.5px;color:var(--dsw-alias-state-warn-primary,#b45309);}
.pt-spin{width:13px;height:13px;border-radius:50%;border:2px solid color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 30%,transparent);border-top-color:var(--dsw-alias-brand-primary,#3b82f6);animation:pt-rot .8s linear infinite;flex:0 0 auto;}
@keyframes pt-rot{to{transform:rotate(360deg);}}
.pt-sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;}
`;

function uiMountStyle() {
  if (typeof document === 'undefined') {
    return () => {};
  }
  const existing = document.getElementById(uiStyleId);
  if (existing) {
    return () => {
      /* owned by an earlier mount */
    };
  }
  const style = document.createElement('style');
  style.id = uiStyleId;
  style.textContent = uiCss;
  document.head.appendChild(style);
  return () => {
    try {
      style.remove();
    } catch {
      /* ignore */
    }
  };
}

/* ---------------------------------------------------------------- image input */

const uiMaxDecodeSide = 6000;

function uiDecodeViaImageTag(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error(uiT('无法读取该图片')));
    };
    img.src = url;
  });
}

/** Decode any image-ish source into { width, height, data, downscaled }. */
async function uiRgbaFromSource(source) {
  const width = source.width || source.naturalWidth || 0;
  const height = source.height || source.naturalHeight || 0;
  if (!width || !height) {
    throw new Error(uiT('无法读取该图片'));
  }
  const ratio = Math.min(1, uiMaxDecodeSide / Math.max(width, height));
  const w = Math.max(1, Math.round(width * ratio));
  const h = Math.max(1, Math.round(height * ratio));
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const context2d = canvas.getContext('2d', { willReadFrequently: true });
  if (!context2d) {
    throw new Error(uiT('该浏览器不支持画布读取像素'));
  }
  context2d.imageSmoothingEnabled = ratio < 1;
  context2d.imageSmoothingQuality = 'high';
  if (ratio < 1) {
    context2d.fillStyle = '#ffffff';
    context2d.fillRect(0, 0, w, h);
  }
  context2d.drawImage(source, 0, 0, w, h);
  const imageData = context2d.getImageData(0, 0, w, h);
  return { width: w, height: h, data: imageData.data, downscaled: ratio < 1 };
}

async function uiDecodeFile(file) {
  let source = null;
  if (typeof createImageBitmap === 'function') {
    try {
      source = await createImageBitmap(file);
    } catch {
      source = null;
    }
  }
  let rgba;
  if (source) {
    try {
      rgba = await uiRgbaFromSource(source);
    } finally {
      if (source && typeof source.close === 'function') {
        source.close();
      }
    }
  } else {
    const img = await uiDecodeViaImageTag(file);
    rgba = await uiRgbaFromSource(img);
  }
  return rgba;
}

/* --------------------------------------------------------------- processing */

function uiNow() {
  try {
    return typeof performance !== 'undefined' && performance && typeof performance.now === 'function'
      ? performance.now()
      : Date.now();
  } catch {
    return Date.now();
  }
}

/** Any pixel that is not fully opaque? Cheap scan, used to pick the reference. */
function uiImageHasAlpha(image) {
  const data = image && image.data;
  if (!data || !data.length) return false;
  for (let index = 3; index < data.length; index += 4) {
    if (data[index] < 255) return true;
  }
  return false;
}

/**
 * Diff two RGBA buffers while ignoring the given rectangles completely, so the
 * report stays self-consistent when text boxes deliberately drop their ink.
 */
function uiDiffIgnoring(reference, candidate, rects) {
  const width = reference.width;
  const height = reference.height;
  const mask = new Uint8Array(width * height);
  for (const rect of rects || []) {
    const x0 = Math.max(0, Math.round(rect.x));
    const y0 = Math.max(0, Math.round(rect.y));
    const x1 = Math.min(width, Math.round(rect.x + rect.w));
    const y1 = Math.min(height, Math.round(rect.y + rect.h));
    for (let y = y0; y < y1; y++) mask.fill(1, y * width + x0, y * width + x1);
  }
  let pixels = 0;
  let mismatchPixels = 0;
  let sum = 0;
  let maxAbs = 0;
  const a = reference.data;
  const b = candidate.data;
  for (let index = 0; index < mask.length; index++) {
    if (mask[index]) continue;
    const p = index * 4;
    let differs = false;
    for (let channel = 0; channel < 4; channel++) {
      const delta = a[p + channel] > b[p + channel] ? a[p + channel] - b[p + channel] : b[p + channel] - a[p + channel];
      if (delta > 0) differs = true;
      sum += delta;
      if (delta > maxAbs) maxAbs = delta;
    }
    pixels += 1;
    if (differs) mismatchPixels += 1;
  }
  return {
    pixels,
    mismatchPixels,
    exactRatio: pixels ? (pixels - mismatchPixels) / pixels : 1,
    meanAbs: pixels ? sum / (pixels * 4) : 0,
    maxAbs,
  };
}

/**
 * Analyse every item with the same normalised options.
 * Pure enough for Node tests: it only calls the scene engine.
 */
async function uiAnalyze(items, options, onProgress) {
  const out = [];
  const list = items || [];
  for (let index = 0; index < list.length; index += 1) {
    const item = list[index];
    if (typeof onProgress === 'function') {
      onProgress(uiT('正在分析…') + ' ' + (index + 1) + '/' + list.length);
    }
    const started = uiNow();
    try {
      const built = sceneBuild(item.image, options);
      const raster = sceneRasterize(built.scene);
      const textShapes = built.scene.shapes.filter((shape) => shape.k === 't');
      // Diff against what the slide really shows, not against raw pixels:
      // transparent areas become the chosen background colour, and a text box is an
      // empty editable box over its ink, so that ink is not a fidelity failure.
      let reference = item.image;
      if (textShapes.length || uiImageHasAlpha(item.image)) {
        reference = sceneRasterize(sceneBuild(item.image, { ...options, mode: 'image' }).scene);
      }
      const diff = sceneDiff(reference, raster, textShapes.length ? { ignoreRects: textShapes } : undefined);
      let stats = built.scene.stats;
      if (textShapes.length) {
        const exact = uiDiffIgnoring(reference, raster, textShapes);
        stats = {
          ...stats,
          exactRatio: exact.exactRatio,
          mismatchPixels: exact.mismatchPixels,
          meanAbs: exact.meanAbs,
          maxAbs: exact.maxAbs,
          notes: [
            ...(stats.notes || []),
            'fidelity ignores ' + textShapes.length + ' text box(es): their ink is stored as an empty editable box',
          ],
        };
      }
      out.push({
        ...item,
        scene: built.scene,
        scale: built.scale,
        stats,
        raster,
        diff,
        ms: Math.round(uiNow() - started),
        error: null,
      });
    } catch (error) {
      out.push({
        ...item,
        scene: null,
        scale: 1,
        stats: null,
        raster: null,
        diff: null,
        ms: Math.round(uiNow() - started),
        error: (error && error.message) || String(error),
      });
    }
  }
  return out;
}

function uiNewItemId() {
  uiItemSeq += 1;
  return 'pt-' + Date.now().toString(36) + '-' + uiItemSeq;
}

let uiItemSeq = 0;

function uiPreviewFrame(item, view) {
  if (!item) {
    return null;
  }
  if (view === 'original') {
    return { data: item.image.data, width: item.image.width, height: item.image.height };
  }
  if (view === 'diff' && item.diff && item.diff.heatmap) {
    return { data: item.diff.heatmap, width: item.image.width, height: item.image.height };
  }
  if (item.raster) {
    return { data: item.raster.data, width: item.raster.width, height: item.raster.height };
  }
  return { data: item.image.data, width: item.image.width, height: item.image.height };
}

/* ------------------------------------------------------------------- format */

function uiPct(value) {
  const ratio = typeof value === 'number' && Number.isFinite(value) ? value : 1;
  const pct = Math.max(0, Math.min(1, ratio)) * 100;
  if (pct >= 99.99995) {
    return '100%';
  }
  if (pct >= 99) {
    return pct.toFixed(2) + '%';
  }
  return pct.toFixed(1) + '%';
}

function uiNum(value, digits) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return '—';
  }
  return digits == null ? String(Math.round(number)) : number.toFixed(digits);
}

function uiFileSize(bytes) {
  const size = Number(bytes) || 0;
  if (size >= 1024 * 1024) {
    return (size / (1024 * 1024)).toFixed(1) + ' MB';
  }
  if (size >= 1024) {
    return Math.round(size / 1024) + ' KB';
  }
  return size + ' B';
}

/* ---------------------------------------------------------------- small UI */

function uiSection(props) {
  return h(
    'section',
    { className: 'pt-card' },
    h(
      'div',
      { className: 'pt-card-head' },
      h('span', { className: 'pt-card-title' }, props.title),
      props.action || null,
    ),
    props.children,
  );
}

function uiField(props) {
  return h(
    'div',
    { className: 'pt-field' },
    h(
      'label',
      { className: 'pt-label' },
      h('span', null, props.label),
      props.value == null ? null : h('span', { className: 'pt-value' }, props.value),
    ),
    props.children,
    props.hint ? h('span', { className: 'pt-hint' }, props.hint) : null,
  );
}

function uiSegmented(props) {
  return h(
    'div',
    { className: 'pt-seg', role: 'group' },
    (props.options || []).map((option) =>
      h(
        'button',
        {
          key: option.value,
          type: 'button',
          className: option.value === props.value ? 'is-on' : '',
          title: option.title || '',
          onClick: () => props.onChange(option.value),
        },
        option.label,
      ),
    ),
  );
}

function uiToggle(props) {
  return h(
    'label',
    { className: 'pt-check' },
    h('input', {
      type: 'checkbox',
      checked: props.checked === true,
      onChange: (event) => props.onChange(event.target.checked),
    }),
    h(
      'span',
      { className: 'pt-check-text' },
      h('span', { className: 'pt-check-title' }, props.label),
      props.hint ? h('span', { className: 'pt-hint' }, props.hint) : null,
    ),
  );
}

function uiPreviewCanvas(props) {
  const frame = props.frame;
  const canvasRef = useRef(null);
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !frame || !frame.data) {
      return;
    }
    if (canvas.width !== frame.width || canvas.height !== frame.height) {
      canvas.width = frame.width;
      canvas.height = frame.height;
    }
    const context2d = canvas.getContext('2d');
    if (!context2d) {
      return;
    }
    try {
      const pixels = frame.data instanceof Uint8ClampedArray ? frame.data : new Uint8ClampedArray(frame.data);
      if (pixels.length !== frame.width * frame.height * 4) {
        return;
      }
      context2d.putImageData(new ImageData(pixels, frame.width, frame.height), 0, 0);
    } catch {
      /* keep the previous frame */
    }
  }, [frame]);
  if (!frame) {
    return null;
  }
  const style =
    props.zoom === 'actual'
      ? { width: frame.width + 'px', height: frame.height + 'px', maxWidth: 'none', imageRendering: 'pixelated' }
      : { width: '100%', height: 'auto', maxWidth: '100%' };
  return h('canvas', { ref: canvasRef, className: 'pt-canvas', width: frame.width, height: frame.height, style });
}

function uiStat(props) {
  const tone = props.tone === 'good' ? ' is-good' : props.tone === 'bad' ? ' is-bad' : '';
  return h(
    'div',
    { className: 'pt-stat' + tone },
    h('span', { className: 'pt-stat-k' }, props.label),
    h('span', { className: 'pt-stat-v' }, props.value),
  );
}

/* --------------------------------------------------------------- main panel */

function uiInitialState() {
  return {
    items: [],
    selected: 0,
    mode: coreDefaults.mode,
    maxColors: coreDefaults.maxColors,
    mergeTolerance: coreDefaults.mergeTolerance,
    shapeBudget: coreDefaults.shapeBudget,
    detailToImage: coreDefaults.detailToImage,
    textRegions: coreDefaults.textRegions,
    alphaBackground: coreDefaults.alphaBackground,
    fileName: '',
    view: 'rebuilt',
    zoom: 'fit',
    busy: false,
    progress: '',
    dirty: false,
    dragOver: false,
    message: null,
  };
}

function uiPanel(props) {
  const locale = props && props.locale ? props.locale : null;
  const [state, setState] = useState(uiInitialState);
  const stateRef = useRef(state);
  stateRef.current = state;
  const fileInputRef = useRef(null);
  const addFilesRef = useRef(null);
  const messageTimerRef = useRef(null);
  uiUseLanguage(locale);

  const patch = (delta) => setState((current) => ({ ...current, ...delta }));

  const flash = (kind, text) => {
    setState((current) => ({ ...current, message: { kind, text } }));
    if (messageTimerRef.current) {
      clearTimeout(messageTimerRef.current);
    }
    messageTimerRef.current = setTimeout(() => {
      messageTimerRef.current = null;
      setState((current) => (current.message && current.message.text === text ? { ...current, message: null } : current));
    }, 6000);
  };
  useEffect(
    () => () => {
      if (messageTimerRef.current) {
        clearTimeout(messageTimerRef.current);
      }
    },
    [],
  );

  const currentOptions = () =>
    coreNormalizeOptions({
      mode: state.mode,
      maxColors: state.maxColors,
      mergeTolerance: state.mergeTolerance,
      shapeBudget: state.shapeBudget,
      detailToImage: state.detailToImage,
      textRegions: state.textRegions,
      alphaBackground: state.alphaBackground,
    });

  const runAnalysis = async (items, options) => {
    if (!items.length) {
      return items;
    }
    setState((current) => ({ ...current, busy: true, progress: uiT('正在分析…'), message: null }));
    const analysed = await uiAnalyze(items, options, (text) =>
      setState((current) => ({ ...current, progress: text })),
    );
    setState((current) => ({ ...current, items: analysed, busy: false, progress: '', dirty: false }));
    return analysed;
  };

  const addFiles = async (fileLike) => {
    const list = Array.from(fileLike || []).filter(
      (file) => file && (/^image\//.test(file.type || '') || /\.(png|jpe?g|webp|gif|bmp|avif)$/i.test(file.name || '')),
    );
    if (!list.length) {
      return;
    }
    setState((current) => ({ ...current, busy: true, progress: uiT('正在读取图片…'), message: null }));
    const added = [];
    for (const file of list) {
      try {
        const rgba = await uiDecodeFile(file);
        added.push({
          id: uiNewItemId(),
          name: coreSafeName(file.name, 'image'),
          width: rgba.width,
          height: rgba.height,
          image: { width: rgba.width, height: rgba.height, data: rgba.data },
          downscaled: rgba.downscaled === true,
          bytes: file.size || 0,
          scene: null,
          scale: 1,
          raster: null,
          diff: null,
          ms: 0,
          error: null,
        });
      } catch (error) {
        flash('error', (error && error.message) || uiT('无法读取该图片'));
      }
    }
    if (!added.length) {
      setState((current) => ({ ...current, busy: false, progress: '' }));
      return;
    }
    const merged = stateRef.current.items.concat(added);
    const options = currentOptions();
    const analysed = await runAnalysis(merged, options);
    const firstNew = analysed.findIndex((item) => item.id === added[0].id);
    setState((current) => ({
      ...current,
      selected: firstNew >= 0 ? firstNew : current.selected,
      fileName: current.fileName || coreSafeName(added[0].name, 'pixeltrace'),
    }));
  };
  addFilesRef.current = addFiles;

  const removeAt = (index) => {
    setState((current) => {
      const items = current.items.filter((_, itemIndex) => itemIndex !== index);
      let selected = current.selected;
      if (index < current.selected) {
        selected -= 1;
      } else if (index === current.selected) {
        selected = Math.min(index, items.length - 1);
      }
      return { ...current, items, selected: Math.max(0, Math.min(selected, Math.max(0, items.length - 1))) };
    });
  };

  const clearAll = () => {
    setState((current) => ({
      ...current,
      items: [],
      selected: 0,
      fileName: '',
      dirty: false,
      message: null,
      progress: '',
      busy: false,
    }));
  };

  const reanalyse = async () => {
    const items = stateRef.current.items;
    if (!items.length) {
      return items;
    }
    return runAnalysis(items, currentOptions());
  };

  const exportDeck = async () => {
    const snapshot = stateRef.current;
    if (!snapshot.items.length) {
      flash('info', uiT('先添加图片才能导出'));
      return;
    }
    const options = coreNormalizeOptions({
      mode: snapshot.mode,
      maxColors: snapshot.maxColors,
      mergeTolerance: snapshot.mergeTolerance,
      shapeBudget: snapshot.shapeBudget,
      detailToImage: snapshot.detailToImage,
      textRegions: snapshot.textRegions,
      alphaBackground: snapshot.alphaBackground,
    });
    const stale = snapshot.items.some((item) => !item.scene || snapshot.dirty);
    let items = snapshot.items;
    if (stale) {
      items = await runAnalysis(snapshot.items, options);
    }
    const usable = items.filter((item) => item.scene && !item.error);
    if (!usable.length) {
      flash('error', uiT('出错'));
      return;
    }
    setState((current) => ({ ...current, busy: true, progress: uiT('正在生成 PPTX…'), message: null }));
    try {
      const result = await coreBuildDeck(usable, {
        title: snapshot.fileName || usable[0].name,
      });
      const baseName = coreSafeName(snapshot.fileName || usable[0].name, 'pixeltrace');
      const fileName = coreDeckFileName(baseName, usable.length);
      const saved = await coreSaveBytes(
        result.bytes,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      );
      setState((current) => ({ ...current, busy: false, progress: '' }));
      if (saved && saved.cancelled) {
        flash('info', uiT('已取消'));
      } else {
        flash('ok', uiT('已导出') + ' · ' + fileName + ' · ' + uiFileSize(result.bytes.length));
      }
    } catch (error) {
      setState((current) => ({ ...current, busy: false, progress: '' }));
      flash('error', (error && error.message) || uiT('出错'));
    }
  };

  useEffect(() => {
    const onPaste = (event) => {
      const data = event.clipboardData;
      if (!data) {
        return;
      }
      const files = [];
      const items = data.items || [];
      for (let index = 0; index < items.length; index += 1) {
        const entry = items[index];
        if (entry && entry.kind === 'file' && /^image\//.test(entry.type || '')) {
          const file = entry.getAsFile();
          if (file) {
            files.push(file);
          }
        }
      }
      if (!files.length) {
        for (const file of data.files || []) {
          if (/^image\//.test(file.type || '')) {
            files.push(file);
          }
        }
      }
      if (files.length) {
        event.preventDefault();
        if (addFilesRef.current) {
          addFilesRef.current(files);
        }
      }
    };
    document.addEventListener('paste', onPaste);
    return () => document.removeEventListener('paste', onPaste);
  }, []);

  const options = currentOptions();
  const selectedItem = state.items[state.selected] || null;
  const frame = uiPreviewFrame(selectedItem, state.view);
  const stats = selectedItem && (selectedItem.stats || selectedItem.scene) ? selectedItem.stats || selectedItem.scene.stats : null;
  const summary = coreSummary(state.items);
  const canExport = state.items.some((item) => item.scene && !item.error) && !state.busy;
  const modeHint =
    state.mode === 'vector'
      ? uiT('全部变成可编辑矩形：每个色块都能单独改，但形状数量最多')
      : state.mode === 'image'
        ? uiT('每页只放一张图片对象：100% 一致，但不能分层编辑')
        : uiT('平面区域变成可编辑色块，细节区域变成精确裁剪的图片，整页仍逐像素一致');

  return h(
    'div',
    { className: 'pt-root' },
    h(
      'header',
      { className: 'pt-head' },
      h(
        'span',
        { className: 'pt-logo', 'aria-hidden': true },
        h(
          'svg',
          { viewBox: '0 0 24 24', width: 21, height: 21, fill: 'currentColor' },
          h('path', { d: 'M3 3h5v5H3zM3 10h5v5H3zM3 17h5v5H3zM10 3h5v5h-5zM10 10h5v5h-5zM10 17h5v5h-5zM17 3h5v5h-5zM17 10h5v5h-5z', opacity: 0.45 }),
          h('path', { d: 'M17 17h5v5h-5z' }),
        ),
      ),
      h(
        'div',
        null,
        h('h1', { className: 'pt-h1' }, uiT('像素复刻 PixelTrace')),
        h(
          'p',
          { className: 'pt-sub' },
          uiT('把一张图还原成 PPT 里可编辑的形状 · 像素级一致'),
        ),
      ),
    ),
    h('input', {
      ref: fileInputRef,
      type: 'file',
      accept: 'image/*',
      multiple: true,
      className: 'pt-sr',
      onChange: (event) => {
        const files = event.target.files;
        addFiles(files);
        event.target.value = '';
      },
    }),
    h(
      'div',
      {
        className: 'pt-drop' + (state.dragOver ? ' is-over' : ''),
        role: 'button',
        tabIndex: 0,
        onClick: () => {
          if (fileInputRef.current) {
            fileInputRef.current.click();
          }
        },
        onKeyDown: (event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            if (fileInputRef.current) {
              fileInputRef.current.click();
            }
          }
        },
        onDragOver: (event) => {
          event.preventDefault();
          if (!state.dragOver) {
            patch({ dragOver: true });
          }
        },
        onDragLeave: () => {
          if (state.dragOver) {
            patch({ dragOver: false });
          }
        },
        onDrop: (event) => {
          event.preventDefault();
          patch({ dragOver: false });
          if (event.dataTransfer && event.dataTransfer.files) {
            addFiles(event.dataTransfer.files);
          }
        },
      },
      h(
        'span',
        { className: 'pt-drop-icon', 'aria-hidden': true },
        h(
          'svg',
          { viewBox: '0 0 24 24', width: 19, height: 19, fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round' },
          h('path', { d: 'M12 16V4' }),
          h('path', { d: 'm7 9 5-5 5 5' }),
          h('path', { d: 'M4 17v2a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-2' }),
        ),
      ),
      h(
        'div',
        { className: 'pt-drop-spacer' },
        h('div', { className: 'pt-drop-title' }, uiT('拖入图片，或点击选择')),
        h(
          'div',
          { className: 'pt-drop-hint' },
          uiT('PNG · JPEG · WebP · GIF · 可多选（每张一页）· 也可以直接 Ctrl+V 粘贴'),
        ),
      ),
      h(
        'button',
        {
          type: 'button',
          className: 'pt-btn pt-primary',
          disabled: state.busy,
          onClick: (event) => {
            event.stopPropagation();
            if (fileInputRef.current) {
              fileInputRef.current.click();
            }
          },
        },
        uiT('选择图片'),
      ),
    ),
    h(
      'div',
      { className: 'pt-form' },
      h(
        'aside',
        { className: 'pt-side' },
        h(
          uiSection,
          {
            title: uiT('图片'),
            action: state.items.length
              ? h(
                  'button',
                  { type: 'button', className: 'pt-btn pt-ghost pt-small', onClick: clearAll, disabled: state.busy },
                  uiT('清空'),
                )
              : null,
          },
          state.items.length
            ? h(
                'div',
                { className: 'pt-files' },
                state.items.map((item, index) =>
                  h(
                    'div',
                    {
                      key: item.id,
                      className: 'pt-file' + (index === state.selected ? ' is-on' : ''),
                      role: 'button',
                      tabIndex: 0,
                      onClick: () => patch({ selected: index }),
                      onKeyDown: (event) => {
                        if (event.key === 'Enter') {
                          patch({ selected: index });
                        }
                      },
                    },
                    h(
                      'span',
                      { className: 'pt-file-body' },
                      h('span', { className: 'pt-file-name', title: item.name }, item.name),
                      h(
                        'span',
                        { className: 'pt-file-meta' },
                        item.width +
                          '×' +
                          item.height +
                          (item.scene ? ' · ' + item.scene.stats.shapes + ' ' + uiT('对象') : '') +
                          (item.error ? ' · ' + uiT('出错') : ''),
                      ),
                    ),
                    h(
                      'button',
                      {
                        type: 'button',
                        className: 'pt-x',
                        title: uiT('移除'),
                        onClick: (event) => {
                          event.stopPropagation();
                          removeAt(index);
                        },
                      },
                      '×',
                    ),
                  ),
                ),
              )
            : h('div', { className: 'pt-empty' }, uiT('还没有图片')),
        ),
        h(
          uiSection,
          { title: uiT('输出参数') },
          h(
            uiField,
            { label: uiT('输出模式'), hint: modeHint },
            h(uiSegmented, {
              value: state.mode,
              options: [
                { value: 'smart', label: uiT('智能分层'), title: uiT('智能分层') },
                { value: 'vector', label: uiT('纯矢量'), title: uiT('纯矢量') },
                { value: 'image', label: uiT('整图保真'), title: uiT('整图保真') },
              ],
              onChange: (value) => patch({ mode: value, dirty: !!state.items.length }),
            }),
          ),
          h(
            uiField,
            { label: uiT('形状预算/页'), value: String(state.shapeBudget), hint: uiT('0 = 不限制；超出预算时自动降级并在统计里说明') },
            h('input', {
              type: 'number',
              className: 'pt-input',
              min: 0,
              max: 200000,
              step: 500,
              value: state.shapeBudget,
              onChange: (event) => patch({ shapeBudget: event.target.value, dirty: !!state.items.length }),
            }),
          ),
          h(
            uiField,
            { label: uiT('颜色容差'), value: String(state.mergeTolerance), hint: uiT('合并相邻矩形时允许的单通道色差，0 = 完全精确') },
            h('input', {
              type: 'range',
              min: 0,
              max: 64,
              step: 1,
              value: state.mergeTolerance,
              style: { width: '100%', accentColor: 'var(--dsw-alias-brand-primary,#3b82f6)' },
              onChange: (event) => patch({ mergeTolerance: event.target.value, dirty: !!state.items.length }),
            }),
          ),
          h(
            uiField,
            { label: uiT('颜色量化'), hint: uiT('0 = 保留原始颜色') },
            h(
              'select',
              {
                className: 'pt-select',
                value: String(state.maxColors),
                onChange: (event) => patch({ maxColors: Number(event.target.value), dirty: !!state.items.length }),
              },
              [0, 16, 32, 64, 128, 256].map((count) =>
                h('option', { key: count, value: String(count) }, count === 0 ? uiT('0 = 保留原始颜色') : String(count)),
              ),
            ),
          ),
          h(uiToggle, {
            label: uiT('细节区域转为精确裁剪图'),
            checked: state.detailToImage,
            onChange: (checked) => patch({ detailToImage: checked, dirty: !!state.items.length }),
          }),
          h(uiToggle, {
            label: uiT('文字识别为可编辑文本框'),
            hint: uiT('实验性：文字墨迹被擦成背景色，留下一个空文本框，方便你自己输入原文'),
            checked: state.textRegions,
            onChange: (checked) => patch({ textRegions: checked, dirty: !!state.items.length }),
          }),
          h(
            uiField,
            { label: uiT('透明背景色') },
            h(
              'div',
              { className: 'pt-row' },
              h('input', {
                type: 'color',
                className: 'pt-swatch',
                value: options.alphaBackground,
                onChange: (event) => patch({ alphaBackground: event.target.value, dirty: !!state.items.length }),
              }),
              h('input', {
                type: 'text',
                className: 'pt-input',
                value: state.alphaBackground,
                spellCheck: false,
                onChange: (event) => patch({ alphaBackground: event.target.value, dirty: !!state.items.length }),
              }),
            ),
          ),
        ),
      ),
      h(
        'main',
        { className: 'pt-stage' },
        h(
          'div',
          { className: 'pt-stage-bar' },
          h(
            'span',
            { className: 'pt-stage-title' },
            selectedItem
              ? selectedItem.name + ' · ' + selectedItem.width + '×' + selectedItem.height
              : uiT('预览'),
          ),
          h(uiSegmented, {
            value: state.view,
            options: [
              { value: 'original', label: uiT('原图') },
              { value: 'rebuilt', label: uiT('重建') },
              { value: 'diff', label: uiT('差异') },
            ],
            onChange: (value) => patch({ view: value }),
          }),
          h(uiSegmented, {
            value: state.zoom,
            options: [
              { value: 'fit', label: uiT('适应') },
              { value: 'actual', label: '1:1' },
            ],
            onChange: (value) => patch({ zoom: value }),
          }),
          h(
            'button',
            { type: 'button', className: 'pt-btn pt-small', onClick: reanalyse, disabled: state.busy || !state.items.length },
            uiT('重新分析'),
          ),
        ),
        h(
          'div',
          { className: 'pt-canvas-wrap' },
          frame
            ? h(uiPreviewCanvas, { frame, zoom: state.zoom })
            : h('div', { className: 'pt-empty' }, uiT('拖入图片后在这里查看逐像素预览')),
        ),
        state.dirty
          ? h('div', { className: 'pt-warn' }, uiT('参数已修改，点「重新分析」刷新预览'))
          : null,
        selectedItem && selectedItem.error
          ? h('div', { className: 'pt-msg error' }, selectedItem.error)
          : null,
        selectedItem && selectedItem.scene
          ? h(
              uiSection,
              { title: uiT('保真统计') },
              h(
                'div',
                { className: 'pt-stats' },
                h(uiStat, { label: uiT('完全一致像素'), value: uiPct(stats.exactRatio), tone: stats.exactRatio >= 1 ? 'good' : 'bad' }),
                h(uiStat, { label: uiT('不一致像素'), value: uiNum(stats.mismatchPixels) }),
                h(uiStat, { label: uiT('对象'), value: uiNum(stats.shapes) }),
                h(uiStat, { label: uiT('矩形'), value: uiNum(stats.rects) }),
                h(uiStat, { label: uiT('裁剪图'), value: uiNum(stats.images) }),
                h(uiStat, { label: uiT('文本框'), value: uiNum(stats.texts) }),
                h(uiStat, { label: uiT('颜色'), value: uiNum(stats.paletteSize) }),
                h(uiStat, { label: uiT('平均误差'), value: uiNum(stats.meanAbs, 2) }),
                h(uiStat, { label: uiT('最大误差'), value: uiNum(stats.maxAbs) }),
                h(uiStat, { label: uiT('耗时'), value: uiNum(selectedItem.ms) + ' ms' }),
              ),
              selectedItem.downscaled
                ? h('div', { className: 'pt-notes' }, h('span', null, uiT('图片过大，已降采样以获得可用速度')))
                : null,
              stats.notes && stats.notes.length
                ? h(
                    'div',
                    { className: 'pt-notes' },
                    h('span', null, uiT('备注') + ':'),
                    stats.notes.map((note, index) => h('code', { key: index }, note)),
                  )
                : null,
            )
          : null,
        state.items.length > 1
          ? h(
              'div',
              { className: 'pt-hint' },
              summary.slides + ' ' + uiT('页') + ' · ' + summary.shapes + ' ' + uiT('对象') + ' · ' + uiPct(summary.exactRatio) + ' · ' + uiNum(summary.ms) + ' ms',
            )
          : null,
      ),
    ),
    state.busy && state.progress
      ? h('div', { className: 'pt-msg info' }, h('span', { className: 'pt-spin' }), state.progress)
      : null,
    state.message
      ? h('div', { className: 'pt-msg ' + (state.message.kind || 'info') }, state.message.text)
      : null,
    h(
      'footer',
      { className: 'pt-foot' },
      h(
        'div',
        { className: 'pt-foot-name' },
        h('span', { className: 'pt-hint' }, uiT('文件名')),
        h('input', {
          type: 'text',
          className: 'pt-input',
          value: state.fileName,
          spellCheck: false,
          onChange: (event) => patch({ fileName: event.target.value }),
        }),
      ),
      h(
        'button',
        { type: 'button', className: 'pt-btn', onClick: reanalyse, disabled: state.busy || !state.items.length },
        uiT('重新分析'),
      ),
      h(
        'button',
        { type: 'button', className: 'pt-btn pt-primary', onClick: exportDeck, disabled: !canExport },
        state.busy && state.progress === uiT('正在生成 PPTX…') ? uiT('正在生成…') : uiT('导出 PPTX'),
      ),
      state.items.length
        ? h('span', { className: 'pt-hint' }, state.items.filter((item) => item.scene).length + ' ' + uiT('页'))
        : null,
      h('span', { className: 'pt-hint', style: { flex: '1 1 auto' } }, ''),
      h('span', { className: 'pt-hint' }, uiT('提示：导出的每个色块、裁剪图、文本框都是独立对象，可在 PowerPoint / WPS / Keynote 里单独编辑。')),
    ),
  );
}

function uiSidebarIcon(props) {
  const raw = props && props.size != null ? props.size : 20;
  const size = typeof raw === 'number' ? raw : parseFloat(raw) || 20;
  return h(
    'svg',
    {
      viewBox: '0 0 24 24',
      width: size,
      height: size,
      fill: 'currentColor',
      'aria-hidden': true,
      style: { display: 'block' },
    },
    h('path', {
      d: 'M3 3h5v5H3zM3 10h5v5H3zM3 17h5v5H3zM10 3h5v5h-5zM10 10h5v5h-5zM10 17h5v5h-5zM17 3h5v5h-5zM17 10h5v5h-5z',
      opacity: 0.45,
    }),
    h('path', { d: 'M17 17h5v5h-5z' }),
  );
}

/** Register the sidebar entry and the main panel. */
function uiApply(ctx) {
  let localeService = null;
  try {
    localeService = (ctx && typeof ctx.get === 'function' && ctx.get('locale')) || (ctx && ctx.locale) || null;
  } catch {
    localeService = (ctx && ctx.locale) || null;
  }
  uiSetLocaleService(localeService);

  if (typeof ctx.effect === 'function') {
    ctx.effect(() => uiMountStyle());
  } else {
    uiMountStyle();
  }

  const openPanel = () => {
    try {
      ctx.layout.selectPanel(uiPanelId);
    } catch {
      /* layout service may not be ready yet */
    }
  };

  ctx.slots.inject('main', () =>
    ctx.slots.inject('sidebar.panellist', () => {
      const stopMain = ctx.slots.register(
        {
          name: 'main',
          key: uiPanelId,
          inject: () => ({ locale: localeService, openPanel }),
        },
        uiPanel,
      );
      const stopIcon = ctx.slots.register(
        {
          name: 'sidebar.panellist',
          id: uiPanelId,
          order: uiSidebarOrder,
          label: () => (uiLanguageOf() === 'en' ? 'PixelTrace' : '像素复刻'),
        },
        uiSidebarIcon,
      );
      return () => {
        stopMain();
        stopIcon();
      };
    }),
  );
}

    function apply(ctx) {
      return uiApply(ctx);
    }

    return { inject: ['slots', 'locale'], apply };
  },
});
