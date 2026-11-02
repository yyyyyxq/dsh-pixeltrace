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
