/* dsh-pixeltrace — scene engine tests (CONTRACT §7).
 *
 * Every fixture is generated procedurally in this file (no binary fixtures, no duplicated sources).
 * The fragment is loaded through test/harness.mjs, i.e. evaluated exactly like the panel does:
 * concatenated source inside `new Function` with no import/export and no DOM.
 *
 * Run from D:\pixeltrace:  node --test test/scene.test.mjs
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { loadScene } from './harness.mjs';

const engine = loadScene();
const { sceneBuild, sceneRasterize, sceneDiff, sceneDefaults } = engine;

/* ------------------------------------------------------------------ *
 * procedural fixtures                                                  *
 * ------------------------------------------------------------------ */

/** Build an RGBA raster from a per-pixel callback returning [r,g,b,a]. */
function paint(w, h, fn) {
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const v = fn(x, y);
      const i = (y * w + x) * 4;
      data[i] = v[0];
      data[i + 1] = v[1];
      data[i + 2] = v[2];
      data[i + 3] = v.length > 3 ? v[3] : 255;
    }
  }
  return { width: w, height: h, data };
}

const checker = (w, h, cell, a = [200, 30, 40], b = [250, 250, 250]) =>
  paint(w, h, (x, y) => (((x / cell) | 0) + ((y / cell) | 0)) % 2 ? a : b);

const solid = (w, h, c) => paint(w, h, () => c);

const gradient = (w, h) =>
  paint(w, h, (x, y) => [Math.round((x * 255) / (w - 1 || 1)), Math.round((y * 255) / (h - 1 || 1)), 128]);

/** Horizontal colour bands: each bar must come out as exactly one rect in vector mode. */
function bands(w, h, colours) {
  const bh = Math.max(1, Math.round(h / colours.length));
  return paint(w, h, (x, y) => colours[Math.min(colours.length - 1, (y / bh) | 0)]);
}

/**
 * Deterministic "screenshot": header, sidebar with menu bars, a gradient photo inset, cards with
 * glyph bars, a table with hairlines, a sine polyline and a footer bar. Proportional, so it can be
 * rendered at any size.
 */
function uiShot(w, h) {
  const img = paint(w, h, () => [246, 247, 249]);
  const put = (x, y, c) => {
    if (x < 0 || y < 0 || x >= w || y >= h) return;
    const i = (y * w + x) * 4;
    img.data[i] = c[0];
    img.data[i + 1] = c[1];
    img.data[i + 2] = c[2];
    img.data[i + 3] = 255;
  };
  const rect = (x, y, rw, rh, c) => {
    for (let yy = Math.round(y); yy < Math.round(y + rh); yy++) {
      for (let xx = Math.round(x); xx < Math.round(x + rw); xx++) put(xx, yy, c);
    }
  };
  const R = (v) => Math.round(v);
  rect(0, 0, w, R(h * 0.06), [31, 45, 61]);
  rect(0, R(h * 0.06), R(w * 0.145), h, [255, 255, 255]);
  for (let m = 0; m < 12; m++) {
    for (let g = 0; g < 7; g++) {
      rect(
        R(w * 0.012) + g * R(w * 0.01),
        R(h * 0.09) + m * R(h * 0.045),
        Math.max(2, R(w * 0.0045)),
        Math.max(2, R(h * 0.012)),
        [90, 100, 115],
      );
    }
  }
  // gradient photo inset (the region smart mode is expected to turn into one exact crop)
  const px = R(w * 0.3);
  const py = R(h * 0.12);
  const pw = R(w * 0.28);
  const ph = R(h * 0.22);
  for (let y = py; y < py + ph; y++) {
    for (let x = px; x < px + pw; x++) {
      const t = ((x - px) / pw) * 0.6 + ((y - py) / ph) * 0.4;
      put(x, y, [R(255 - 200 * t), R(122 + 60 * t), R(24 + 180 * t)]);
    }
  }
  for (let c = 0; c < 3; c++) {
    const cx = R(w * 0.18) + c * R(w * 0.26);
    rect(cx, R(h * 0.42), R(w * 0.24), R(h * 0.14), [255, 255, 255]);
    for (let gl = 0; gl < 8; gl++) {
      rect(cx + R(w * 0.012) + gl * R(w * 0.008), R(h * 0.445), Math.max(2, R(w * 0.005)), Math.max(2, R(h * 0.011)), [17, 24, 39]);
    }
    for (let gl = 0; gl < 5; gl++) {
      rect(cx + R(w * 0.012) + gl * R(w * 0.009), R(h * 0.47), Math.max(2, R(w * 0.006)), Math.max(2, R(h * 0.016)), [17, 24, 39]);
    }
  }
  for (let r = 0; r < 9; r++) {
    rect(R(w * 0.18), R(h * 0.62) + r * R(h * 0.03), R(w * 0.78), 1, [229, 231, 235]);
    for (let gl = 0; gl < 10; gl++) {
      rect(
        R(w * 0.19) + gl * R(w * 0.008),
        R(h * 0.63) + r * R(h * 0.03),
        Math.max(2, R(w * 0.005)),
        Math.max(2, R(h * 0.009)),
        gl % 3 === 0 ? [17, 24, 39] : [90, 100, 115],
      );
    }
  }
  for (let t = 0; t < R(h * 0.22); t++) {
    const x = R(w * 0.62) + R(Math.sin(t / 20) * w * 0.05);
    put(x, R(h * 0.68) + t, [244, 114, 182]);
    put(x + 1, R(h * 0.68) + t, [244, 114, 182]);
  }
  rect(0, h - R(h * 0.03), w, R(h * 0.03), [238, 240, 243]);
  return img;
}

/** Deterministic value-noise photo (3 octaves) — a "must become a handful of crops" source. */
function hash2(x, y) {
  let h = (x * 374761393 + y * 668265263) | 0;
  h = Math.imul(h ^ (h >>> 13), 1274126177);
  return ((h ^ (h >>> 16)) >>> 0) / 4294967296;
}
function photo(w, h) {
  return paint(w, h, (x, y) => {
    const v = hash2(x >> 2, y >> 2) * 0.55 + hash2(x >> 4, y >> 4) * 0.3 + hash2(x >> 6, y >> 6) * 0.15;
    return [Math.round(30 + v * 200), Math.round(50 + v * 180), Math.round(80 + v * 140)];
  });
}

/**
 * Text-like raster: 4 lines of 9 glyphs, 7x12 ink boxes on white. `aa` adds the lighter edge column
 * that a real screenshot has after antialiasing.
 */
function textImage(w, h, aa) {
  const img = paint(w, h, () => [255, 255, 255]);
  const put = (x, y, c) => {
    if (x < 0 || y < 0 || x >= w || y >= h) return;
    const i = (y * w + x) * 4;
    img.data[i] = c[0];
    img.data[i + 1] = c[1];
    img.data[i + 2] = c[2];
    img.data[i + 3] = 255;
  };
  const ink = [32, 32, 36];
  const edge = [150, 150, 158];
  for (let line = 0; line < 4; line++) {
    const y = 20 + line * 22;
    for (let g = 0; g < 9; g++) {
      const x = 20 + g * 11;
      for (let yy = 0; yy < 12; yy++) {
        put(x + 1, y + yy, ink);
        if (aa) put(x, y + yy, edge);
        if (yy === 0 || yy === 11) {
          for (let xx = 0; xx < 7; xx++) put(x + xx, y + yy, xx === 0 || xx === 6 ? (aa ? edge : ink) : ink);
        }
      }
    }
  }
  return img;
}

/** A text block on the left plus a gradient "photo" on the right: text and crops must not overlap. */
function textWithPhoto(w, h) {
  const img = textImage(w, h, true);
  const put = (x, y, c) => {
    if (x < 0 || y < 0 || x >= w || y >= h) return;
    const i = (y * w + x) * 4;
    img.data[i] = c[0];
    img.data[i + 1] = c[1];
    img.data[i + 2] = c[2];
    img.data[i + 3] = 255;
  };
  const px = Math.round(w * 0.62);
  const py = Math.round(h * 0.15);
  const pw = Math.round(w * 0.3);
  const ph = Math.round(h * 0.5);
  for (let y = py; y < py + ph; y++) {
    for (let x = px; x < px + pw; x++) {
      const t = ((x - px) / pw) * 0.6 + ((y - py) / ph) * 0.4;
      put(x, y, [Math.round(255 - 200 * t), Math.round(122 + 60 * t), Math.round(24 + 180 * t)]);
    }
  }
  return img;
}

/* ------------------------------------------------------------------ *
 * test helpers                                                         *
 * ------------------------------------------------------------------ */

const hexRgb = (hex) => [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];

/**
 * Composite the source over the scene background exactly as the contract defines alpha-over
 * (a=0 → background, 0<a<255 → rounded mix, a=255 → source). The scene engine is supposed to
 * rebuild this raster, so fidelity assertions compare against it, never against raw alpha.
 */
function composited(image, bgHex) {
  const [br, bg, bb] = hexRgb(bgHex);
  const data = new Uint8ClampedArray(image.width * image.height * 4);
  for (let i = 0; i < image.width * image.height; i++) {
    const p = i * 4;
    const a = image.data[p + 3];
    let r = image.data[p];
    let g = image.data[p + 1];
    let b = image.data[p + 2];
    if (a === 0) {
      r = br;
      g = bg;
      b = bb;
    } else if (a !== 255) {
      const inv = 255 - a;
      r = ((r * a + br * inv + 127) / 255) | 0;
      g = ((g * a + bg * inv + 127) / 255) | 0;
      b = ((b * a + bb * inv + 127) / 255) | 0;
    }
    data[p] = r;
    data[p + 1] = g;
    data[p + 2] = b;
    data[p + 3] = 255;
  }
  return { width: image.width, height: image.height, data };
}

const rectsOf = (scene) => scene.shapes.filter((s) => s.k === 'r');
const cropsOf = (scene) => scene.shapes.filter((s) => s.k === 'i');
const textsOf = (scene) => scene.shapes.filter((s) => s.k === 't');
const ignoreRectsOf = (scene) => textsOf(scene).map((t) => ({ x: t.x, y: t.y, w: t.w, h: t.h }));
const cropBoxes = (scene) => cropsOf(scene).map((c) => [c.x, c.y, c.w, c.h]);

/** Raster of the scene, as an image the public sceneDiff can consume. */
function rasterOf(scene) {
  const r = sceneRasterize(scene);
  return { width: r.width, height: r.height, data: r.data };
}

/** sceneDiff of the rebuilt scene against the composited source. */
function diffVsSource(image, scene, opts) {
  return sceneDiff(composited(image, scene.background), rasterOf(scene), opts);
}

/** Assert the rebuilt scene is pixel-identical to the composited source. */
function assertPixelExact(image, scene) {
  const d = diffVsSource(image, scene);
  assert.equal(d.mismatchPixels, 0, 'expected zero mismatching pixels');
  assert.equal(d.exactRatio, 1, 'expected exactRatio 1');
  assert.equal(d.maxAbs, 0, 'expected maxAbs 0');
  assert.equal(d.meanAbs, 0, 'expected meanAbs 0');
  assert.equal(scene.stats.exactRatio, 1, 'stats.exactRatio must agree');
  assert.equal(scene.stats.mismatchPixels, 0, 'stats.mismatchPixels must agree');
}

/** Every shape is integral, inside the scene and carries contract-valid colours/roles. */
function assertShapeGeometry(scene) {
  assert.ok(scene.width >= 1 && scene.height >= 1, 'scene size at least 1x1');
  assert.ok(Number.isInteger(scene.width) && Number.isInteger(scene.height));
  assert.match(scene.background, /^#[0-9a-f]{6}$/);
  for (const s of scene.shapes) {
    assert.ok(s.k === 'r' || s.k === 'i' || s.k === 't', 'unknown shape kind ' + s.k);
    for (const f of ['x', 'y', 'w', 'h']) {
      assert.ok(Number.isInteger(s[f]), s.k + '.' + f + ' must be integer, got ' + s[f]);
    }
    assert.ok(s.w >= 1 && s.h >= 1, 'shape must have positive size');
    assert.ok(s.x >= 0 && s.y >= 0, 'shape origin inside the scene');
    assert.ok(s.x + s.w <= scene.width && s.y + s.h <= scene.height, 'shape must not stick out of the scene');
    if (s.k === 'r') {
      assert.match(s.c, /^#[0-9a-f]{6}$/, 'rect colour must be lowercase hex');
      assert.ok(s.role === 'bg' || s.role === 'flat' || s.role === 'fix', 'rect role');
    } else if (s.k === 'i') {
      assert.deepEqual(s.px.width, s.w);
      assert.deepEqual(s.px.height, s.h);
      assert.equal(s.px.data.length, s.w * s.h * 4, 'crop pixel count');
      assert.ok(s.role === 'photo' || s.role === 'text' || s.role === 'detail', 'crop role');
    } else {
      assert.equal(typeof s.s, 'string');
      assert.ok(Number.isInteger(s.sz) && s.sz >= 6 && s.sz <= 96, 'text size clamped to 6..96');
      assert.match(s.c, /^#[0-9a-f]{6}$/);
      assert.ok(s.al === 'l' || s.al === 'ctr' || s.al === 'r');
      assert.equal(typeof s.b, 'boolean');
      assert.equal(s.role, 'text');
    }
  }
}

/** Paint-order palette check: rect colours by descending area. */
function assertPalette(scene) {
  const area = new Map();
  for (const s of rectsOf(scene)) area.set(s.c, (area.get(s.c) || 0) + s.w * s.h);
  const expected = [...area.entries()].sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0)).map((e) => e[0]);
  assert.deepEqual(scene.palette, expected, 'palette is the rect colours by descending area');
  assert.equal(scene.stats.paletteSize, scene.palette.length);
  assert.equal(new Set(scene.palette).size, scene.palette.length, 'palette has no duplicates');
}

/**
 * Coverage invariant: every pixel is painted by exactly one shape. Text mode intentionally paints
 * the erase rect and then the empty text box on top of it, so overlaps are only allowed inside
 * text boxes (which also carry the excluded ink pixels).
 */
function coverage(scene) {
  const w = scene.width;
  const h = scene.height;
  const count = new Uint8Array(w * h);
  for (const s of scene.shapes) {
    for (let y = s.y; y < s.y + s.h; y++) {
      const row = y * w;
      for (let x = s.x; x < s.x + s.w; x++) count[row + x]++;
    }
  }
  const ignore = new Uint8Array(w * h);
  for (const r of ignoreRectsOf(scene)) {
    for (let y = r.y; y < r.y + r.h; y++) {
      const row = y * w;
      ignore.fill(1, row + r.x, row + r.x + r.w);
    }
  }
  let overlap = 0;
  let gap = 0;
  let overlapOutsideIgnore = 0;
  let ignored = 0;
  for (let i = 0; i < count.length; i++) {
    if (ignore[i]) ignored++;
    if (count[i] > 1) {
      overlap++;
      if (!ignore[i]) overlapOutsideIgnore++;
    } else if (count[i] === 0) gap++;
  }
  return { overlap, gap, overlapOutsideIgnore, ignored, overlapPixels: overlap };
}

/** No seams: shapes tile the scene with no overlap (outside text boxes) and no hole. */
function assertNoSeam(scene, label) {
  const cov = coverage(scene);
  assert.equal(cov.overlapOutsideIgnore, 0, label + ': shapes must not overlap outside text boxes');
  assert.equal(cov.gap, 0, label + ': every pixel must be covered (no seams), gap=' + cov.gap);
  return cov;
}

/** Each crop must be an exact straight-alpha copy of the composited source. */
function assertCropsExact(image, scene) {
  const comp = composited(image, scene.background);
  for (const c of cropsOf(scene)) {
    assert.equal(c.px.width, c.w);
    assert.equal(c.px.height, c.h);
    for (let y = 0; y < c.h; y++) {
      for (let x = 0; x < c.w; x++) {
        const si = ((c.y + y) * comp.width + (c.x + x)) * 4;
        const pi = (y * c.w + x) * 4;
        for (let k = 0; k < 4; k++) {
          assert.equal(
            c.px.data[pi + k],
            comp.data[si + k],
            'crop ' + [c.x, c.y, c.w, c.h] + ' byte mismatch at ' + x + ',' + y + ' channel ' + k,
          );
        }
      }
    }
  }
}

/** FNV-1a over a byte array, used for deterministic-bytes checks of crop payloads. */
function hashBytes(bytes) {
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) {
    h ^= bytes[i];
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16);
}

function fnv1a(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16);
}

/** Stable hash of the whole shape list (geometry, colours, text fields, crop payload bytes). */
function hashScene(scene) {
  const parts = [scene.width, scene.height, scene.background];
  for (const s of scene.shapes) {
    let p = s.k + ':' + s.x + ',' + s.y + ',' + s.w + ',' + s.h + ':' + s.role;
    if (s.k === 'r') p += ':' + s.c;
    if (s.k === 'i') p += ':' + s.px.width + 'x' + s.px.height + ':' + hashBytes(s.px.data);
    if (s.k === 't') p += ':' + s.sz + ':' + s.c + ':' + s.al + ':' + s.b + ':' + (s.f || '') + ':' + s.s;
    parts.push(p);
  }
  parts.push('pal=' + scene.palette.join(','));
  return fnv1a(parts.join('|'));
}

/* ------------------------------------------------------------------ *
 * fixtures used by several tests                                       *
 * ------------------------------------------------------------------ */

const FIX = {
  checker64: checker(64, 64, 4),
  checker64cell8: checker(64, 64, 8),
  gradient32: gradient(32, 32),
  solid300: solid(300, 200, [10, 20, 30]),
  ui640: uiShot(640, 400),
  ui1600: uiShot(1600, 900),
  photo320: photo(320, 200),
  photo1920: photo(1920, 1080),
  bands: bands(1200, 40, Array.from({ length: 40 }, (_, k) => [k * 6, 255 - k * 6, (k * 37) % 256])),
  bars: bands(64, 120, [[200, 0, 0], [0, 200, 0], [0, 0, 200]]),
  textAA: textImage(320, 200, true),
  textSolid: textImage(320, 200, false),
  textWithPhoto: textWithPhoto(480, 260),
  transparent: paint(8, 8, () => [0, 0, 0, 0]),
  mostlyTransparent: paint(16, 16, (x, y) => [200, 10, 20, (x + y) % 3 === 0 ? 0 : 128]),
  onePixel: paint(1, 1, () => [7, 8, 9]),
  wideSolid: solid(6000, 10, [1, 2, 3]),
};

/* ------------------------------------------------------------------ *
 * 1. public surface                                                    *
 * ------------------------------------------------------------------ */

test('public surface: four contract functions and the documented defaults', () => {
  assert.equal(typeof sceneBuild, 'function');
  assert.equal(typeof sceneRasterize, 'function');
  assert.equal(typeof sceneDiff, 'function');
  assert.equal(typeof sceneDefaults, 'object');

  assert.deepEqual(
    { ...sceneDefaults },
    {
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
    },
  );
});

test('sceneBuild returns {scale, scene} and never mutates sceneDefaults', () => {
  const before = JSON.stringify(sceneDefaults);
  const out = sceneBuild(FIX.ui640, { mode: 'vector', maxColors: 8, shapeBudget: 10, textRegions: true, maxSlidePx: 1000 });
  assert.equal(typeof out.scale, 'number');
  assert.ok(out.scale > 0 && out.scale <= 1);
  assert.equal(out.scene.width, 640);
  assert.equal(out.scene.height, 400);
  assert.equal(JSON.stringify(sceneDefaults), before, 'sceneDefaults must stay untouched');
  assert.notEqual(out.scene.background, undefined);
});

/* ------------------------------------------------------------------ *
 * 2. determinism                                                       *
 * ------------------------------------------------------------------ */

test('deterministic bytes: identical input -> identical shape list (in process and fresh load)', () => {
  const a = sceneBuild(FIX.ui640, { shapeBudget: 0 }).scene;
  const b = sceneBuild(FIX.ui640, { shapeBudget: 0 }).scene;
  assert.equal(hashScene(a), hashScene(b), 'two builds of the same input must hash identically');
  assert.deepEqual(cropBoxes(a), cropBoxes(b));
  assert.equal(a.stats.shapes, b.stats.shapes);

  // A second, independent evaluation of the fragment must produce the same bytes as well.
  const fresh = loadScene();
  const c = fresh.sceneBuild(FIX.ui640, { shapeBudget: 0 }).scene;
  assert.equal(hashScene(c), hashScene(a), 'a freshly loaded fragment must be deterministic too');

  // Different inputs must not collapse to the same hash (sanity check of the hash itself).
  const d = sceneBuild(FIX.checker64, { mode: 'vector', shapeBudget: 0 }).scene;
  assert.notEqual(hashScene(d), hashScene(a));
});

/* ------------------------------------------------------------------ *
 * 3. exact reconstruction: all three modes                             *
 * ------------------------------------------------------------------ */

test('vector mode with no tolerance / no budget is pixel-exact and uses only rects', () => {
  const cases = [
    ['checker 64/4', FIX.checker64, 256],
    ['gradient 32', FIX.gradient32, 1024],
    ['single colour 300x200', FIX.solid300, 1],
  ];
  for (const [label, image, expectedRects] of cases) {
    const { scene } = sceneBuild(image, { mode: 'vector', maxColors: 0, mergeTolerance: 0, shapeBudget: 0 });
    assertPixelExact(image, scene);
    assert.equal(cropsOf(scene).length, 0, label + ': vector mode must not create pictures');
    assert.equal(textsOf(scene).length, 0, label + ': no text boxes without textRegions');
    assert.equal(rectsOf(scene).length, scene.stats.rects);
    assert.equal(rectsOf(scene).length, expectedRects, label + ': exact maximal-rect count');
    assertNoSeam(scene, label);
    assertShapeGeometry(scene);
    assertPalette(scene);
  }
});

test('vector mode merges identical row runs vertically (one maximal rect per band)', () => {
  const { scene } = sceneBuild(FIX.bars, { mode: 'vector', shapeBudget: 0 });
  assertPixelExact(FIX.bars, scene);
  assert.equal(scene.stats.rects, 3, 'three colour bands -> three maximal rects');
  assert.deepEqual(
    rectsOf(scene).map((r) => [r.x, r.y, r.w, r.h]),
    [
      [0, 0, 64, 40],
      [0, 40, 64, 40],
      [0, 80, 64, 40],
    ],
  );
  assert.deepEqual(scene.palette, ['#0000c8', '#00c800', '#c80000'], 'palette descends by area then colour');
});

test('image mode is exactly one full-size picture and pixel-exact', () => {
  for (const image of [FIX.ui640, FIX.photo320, FIX.checker64]) {
    const { scene } = sceneBuild(image, { mode: 'image' });
    assert.equal(scene.shapes.length, 1, 'image mode = one shape');
    const [crop] = cropsOf(scene);
    assert.ok(crop, 'the single shape is a picture');
    assert.deepEqual([crop.x, crop.y, crop.w, crop.h], [0, 0, image.width, image.height]);
    assert.equal(crop.role, 'photo');
    assert.equal(scene.palette.length, 0);
    assertPixelExact(image, scene);
    assertCropsExact(image, scene);
    assertShapeGeometry(scene);
    assert.equal(scene.stats.images, 1);
  }
});

test('smart mode with an unlimited budget is pixel-exact and region-exclusive', () => {
  const cases = [
    ['ui 640x400', FIX.ui640],
    ['ui 400x250', uiShot(400, 250)],
    ['checker 400x300', checker(400, 300, 8)],
    ['photo 320x200', FIX.photo320],
  ];
  for (const [label, image] of cases) {
    const { scene } = sceneBuild(image, { shapeBudget: 0 });
    assertPixelExact(image, scene);
    const cov = assertNoSeam(scene, label);
    assert.equal(cov.overlap, 0, label + ': without text boxes every pixel is painted exactly once');
    assertCropsExact(image, scene);
    assertShapeGeometry(scene);
    assertPalette(scene);
  }

  // a photo has no flat area worth tiling: one exact crop is the right answer
  const photoScene = sceneBuild(FIX.photo320, { shapeBudget: 0 }).scene;
  assert.equal(photoScene.shapes.length, 1);
  assert.equal(photoScene.stats.images, 1);

  // a checkerboard is flat: smart must fall back to exact rects, not a crop
  const checkerScene = sceneBuild(checker(400, 300, 8), { shapeBudget: 0 }).scene;
  assert.equal(checkerScene.stats.images, 0);
  assert.ok(checkerScene.stats.rects > 100);
});

test('smart mode keeps the flat UI rects and turns only the detail regions into crops', () => {
  const { scene } = sceneBuild(FIX.ui640, { shapeBudget: 0 });
  const crops = cropBoxes(scene);
  assert.ok(crops.length >= 1, 'the gradient inset must be a crop');
  // the gradient photo inset sits at (0.30w, 0.12h) size (0.28w, 0.22h) = (192,48,184,88)
  assert.ok(
    crops.some((b) => b[0] <= 192 && b[1] <= 48 && b[0] + b[2] >= 376 && b[1] + b[3] >= 136),
    'gradient inset must be inside a crop, got ' + JSON.stringify(crops),
  );
  for (const b of crops) {
    // a crop only contains detail pixels: assert its pixels are not a single colour
    const first = scene.shapes.find((s) => s.k === 'i' && s.x === b[0] && s.y === b[1]);
    const colours = new Set();
    for (let i = 0; i < first.px.data.length && colours.size < 2; i += 4) {
      colours.add((first.px.data[i] << 16) | (first.px.data[i + 1] << 8) | first.px.data[i + 2]);
    }
    assert.ok(colours.size > 1, 'a crop must contain more than one colour');
  }
});

/* ------------------------------------------------------------------ *
 * 4. budget behaviour                                                  *
 * ------------------------------------------------------------------ */

test('detailToImage:false keeps every shape editable as a rect; regionMinPx drops tiny crops', () => {
  const flatOnly = sceneBuild(FIX.ui640, { shapeBudget: 0, detailToImage: false }).scene;
  assert.equal(cropsOf(flatOnly).length, 0, 'detailToImage:false must not create crops');
  assertPixelExact(FIX.ui640, flatOnly);
  assertNoSeam(flatOnly, 'detailToImage false');
  assertShapeGeometry(flatOnly);

  const base = sceneBuild(FIX.ui640, { shapeBudget: 0 }).scene;
  const coarse = sceneBuild(FIX.ui640, { shapeBudget: 0, regionMinPx: 64 }).scene;
  assert.ok(
    cropsOf(coarse).length <= cropsOf(base).length,
    'regionMinPx must not add crops: ' + cropsOf(coarse).length + ' vs ' + cropsOf(base).length,
  );
  assertPixelExact(FIX.ui640, coarse);
  assertNoSeam(coarse, 'regionMinPx 64');
});

test('budget is never exceeded and degrades monotonically while staying exact', () => {
  const unlimited = sceneBuild(FIX.ui640, { shapeBudget: 0 }).scene.stats.shapes;
  const budgets = [8000, 300, 60, 20, 5, 2, 1];
  let previous = Infinity;
  for (const budget of budgets) {
    const { scene } = sceneBuild(FIX.ui640, { shapeBudget: budget });
    assert.ok(
      scene.stats.shapes <= budget,
      'budget ' + budget + ' exceeded: ' + scene.stats.shapes + ' shapes',
    );
    assert.ok(scene.stats.shapes <= previous, 'tighter budget must not produce more shapes');
    assertPixelExact(FIX.ui640, scene);
    assertNoSeam(scene, 'budget ' + budget);
    assertShapeGeometry(scene);
    if (scene.stats.shapes < unlimited) {
      assert.ok(scene.stats.notes.length > 0, 'a degraded scene must explain itself in stats.notes');
      assert.ok(
        scene.stats.notes.some((n) => n.indexOf('budget ' + budget) === 0),
        'note must mention the budget, got ' + JSON.stringify(scene.stats.notes),
      );
    }
    previous = scene.stats.shapes;
  }
  // a budget of 1 can only be satisfied by the single whole-image crop
  const last = sceneBuild(FIX.ui640, { shapeBudget: 1 }).scene;
  assert.equal(last.shapes.length, 1);
  assert.equal(last.shapes[0].k, 'i');
  assert.deepEqual([last.shapes[0].w, last.shapes[0].h], [640, 400]);
});

test('vector mode under budget pressure falls back to one picture with an honest note', () => {
  const { scene } = sceneBuild(FIX.ui640, { mode: 'vector', shapeBudget: 60 });
  assert.equal(scene.shapes.length, 1);
  assert.equal(scene.shapes[0].k, 'i');
  assertPixelExact(FIX.ui640, scene);
  assert.ok(
    scene.stats.notes.some((n) => n.indexOf('budget 60') === 0 && n.indexOf('picture') >= 0),
    'notes must explain the budget fallback, got ' + JSON.stringify(scene.stats.notes),
  );
});

test('smart mode degrades through merge tolerance before falling back to a picture (step b)', () => {
  // 25 flat bands whose adjacent colours differ by 2 per channel: no detail tiles, so growth
  // cannot help and only a raised merge tolerance can meet a small budget.
  const image = paint(400, 100, (x, y) => {
    const k = (y / 4) | 0;
    return [k * 2, k * 2, k * 2];
  });
  const budget = 20;
  const { scene } = sceneBuild(image, { shapeBudget: budget });
  assert.ok(scene.stats.shapes <= budget, 'budget ' + budget + ' exceeded: ' + scene.stats.shapes);
  assert.ok(scene.stats.rects > 0, 'the tolerance step keeps editable rects instead of a full crop');
  assert.equal(scene.stats.images, 0);
  const d = diffVsSource(image, scene);
  assert.ok(d.maxAbs <= 2, 'merged colours must stay inside the raised tolerance, maxAbs=' + d.maxAbs);
  assert.ok(d.exactRatio < 1, 'a lossy tolerance merge must be reported honestly');
  assert.ok(
    scene.stats.notes.some((n) => n.indexOf('raised merge tolerance') >= 0),
    'notes must record the tolerance step, got ' + JSON.stringify(scene.stats.notes),
  );
  assert.ok(scene.stats.notes.some((n) => n.indexOf('fidelity below 1') >= 0));
  assertNoSeam(scene, 'tolerance degradation');
});

test('mergeTolerance is a real per-pixel error bound and buys fewer shapes', () => {
  const exact = sceneBuild(FIX.ui640, { shapeBudget: 0, mergeTolerance: 0 }).scene;
  const merged = sceneBuild(FIX.ui640, { shapeBudget: 0, mergeTolerance: 32 }).scene;
  assert.ok(merged.stats.shapes <= exact.stats.shapes, 'tolerance must not add shapes');
  const d = diffVsSource(FIX.ui640, merged);
  assert.ok(d.maxAbs <= 32, 'maxAbs ' + d.maxAbs + ' must respect mergeTolerance 32');
  assert.ok(d.exactRatio <= 1);
  if (merged.stats.shapes < exact.stats.shapes) {
    assert.ok(d.mismatchPixels > 0, 'a lossy merge must be reported honestly');
    assert.ok(d.exactRatio < 1);
    assertNoSeam(merged, 'tolerance 32');
  }
});

/* ------------------------------------------------------------------ *
 * 5. palette quantisation (regression: crops are quantisation-independent) *
 * ------------------------------------------------------------------ */

test('maxColors never changes crop rectangles and keeps the palette bounded', () => {
  const base = sceneBuild(FIX.ui640, { shapeBudget: 0 }).scene;
  const baseCrops = cropBoxes(base);
  assert.ok(baseCrops.length >= 1);
  for (const maxColors of [8, 16, 32, 64]) {
    const { scene } = sceneBuild(FIX.ui640, { shapeBudget: 0, maxColors });
    assert.deepEqual(cropBoxes(scene), baseCrops, 'crops must not depend on maxColors (' + maxColors + ')');
    assert.ok(scene.palette.length <= maxColors, 'palette ' + scene.palette.length + ' > maxColors ' + maxColors);
    assert.ok(scene.stats.paletteSize <= maxColors);
    // the crop bytes stay exact source crops; only flat rect colours may change
    assertCropsExact(FIX.ui640, scene);
    const d = diffVsSource(FIX.ui640, scene, { ignoreRects: cropBoxes(scene).map((b) => ({ x: b[0], y: b[1], w: b[2], h: b[3] })) });
    assert.equal(d.exactRatioIgnoring, 1, 'crops stay exact under quantisation');
  }
});

test('median-cut colour reduction is monotone in error and exact when the image has few colours', () => {
  const mean = [];
  for (const maxColors of [8, 16, 32]) {
    const { scene } = sceneBuild(FIX.bands, { mode: 'vector', shapeBudget: 0, maxColors });
    assert.equal(scene.palette.length, maxColors, 'palette should use the whole budget on a rich image');
    assert.ok(scene.palette.length <= maxColors);
    mean.push(diffVsSource(FIX.bands, scene).meanAbs);
  }
  assert.ok(mean[0] > mean[1] && mean[1] > mean[2], 'more allowed colours must reduce error, got ' + mean.join(' > '));

  // an image that already fits the budget is left untouched (still exact)
  const few = checker(64, 64, 4);
  const { scene } = sceneBuild(few, { mode: 'vector', shapeBudget: 0, maxColors: 8 });
  assertPixelExact(few, scene);
  assert.equal(scene.stats.notes.filter((n) => n.indexOf('quantised') >= 0).length, 0);
});

/* ------------------------------------------------------------------ *
 * 6. text regions                                                      *
 * ------------------------------------------------------------------ */

test('text regions become empty editable text boxes over an erase rect, rest of the image stays exact', () => {
  const image = FIX.textAA;
  const { scene } = sceneBuild(image, { shapeBudget: 0, textRegions: true });
  const texts = textsOf(scene);
  assert.ok(texts.length >= 1, 'text must be detected');

  const d = diffVsSource(image, scene);
  assert.ok(d.mismatchPixels > 0, 'ink pixels are erased, so plain exactRatio must be below 1');
  assert.ok(d.exactRatio < 1);
  const dIgnored = diffVsSource(image, scene, { ignoreRects: ignoreRectsOf(scene) });
  assert.equal(dIgnored.exactRatioIgnoring, 1, 'everything outside text boxes must be exact');

  assert.ok(
    scene.stats.notes.some((n) => n.indexOf('block(s)') >= 0),
    'notes must report detected blocks, got ' + JSON.stringify(scene.stats.notes),
  );

  const order = new Map(scene.shapes.map((s, i) => [s, i]));
  for (const t of texts) {
    // geometry: zero insets -> the box is exactly the ink bbox, integer, inside the scene
    assert.ok(t.w >= 1 && t.h >= 1);
    assert.equal(t.s, '', 'detected text boxes are empty (editable placeholder)');
    assert.equal(t.al, 'l');
    assert.equal(t.b, false);
    assert.equal(t.f, 'Arial');
    // fixture glyph height is 12px: sz = clamp(round(inkH*0.9*0.75),6,96) = clamp(round(8.1)) = 8
    assert.equal(t.sz, 8, 'font size derives from ink height');
    // exactly one erase rect, painted first, covering the same rect with the dominant background
    const erase = rectsOf(scene).filter(
      (r) => r.x === t.x && r.y === t.y && r.w === t.w && r.h === t.h && r.role === 'bg',
    );
    assert.equal(erase.length, 1, 'one erase rect per text box');
    assert.ok(order.get(erase[0]) < order.get(t), 'the erase rect is painted before the text box');
    assert.match(erase[0].c, /^#[0-9a-f]{6}$/);
  }
  assertShapeGeometry(scene);
  assertNoSeam(scene, 'text regions');
});

test('text boxes and detail crops never overlap', () => {
  const image = FIX.textWithPhoto;
  const { scene } = sceneBuild(image, { shapeBudget: 0, textRegions: true });
  const texts = textsOf(scene);
  const crops = cropsOf(scene);
  assert.ok(texts.length >= 1, 'text must be detected');
  assert.ok(crops.length >= 1, 'the gradient photo must be a crop');
  for (const t of texts) {
    for (const c of crops) {
      const ox = Math.max(t.x, c.x);
      const oy = Math.max(t.y, c.y);
      const ex = Math.min(t.x + t.w, c.x + c.w);
      const ey = Math.min(t.y + t.h, c.y + c.h);
      assert.ok(ex <= ox || ey <= oy, 'text box ' + [t.x, t.y, t.w, t.h] + ' overlaps crop ' + [c.x, c.y, c.w, c.h]);
    }
  }
  const d = diffVsSource(image, scene, { ignoreRects: ignoreRectsOf(scene) });
  assert.equal(d.exactRatioIgnoring, 1);
  assertNoSeam(scene, 'text + photo');
});

test('textRegions off leaves text as exact geometry; no text found is reported honestly', () => {
  const { scene } = sceneBuild(FIX.textAA, { shapeBudget: 0, textRegions: false });
  assert.equal(textsOf(scene).length, 0);
  assertPixelExact(FIX.textAA, scene);

  const ui = sceneBuild(FIX.ui640, { shapeBudget: 0, textRegions: true }).scene;
  assert.equal(textsOf(ui).length, 0, 'UI bars are too thin to be text blocks');
  assert.ok(
    ui.stats.notes.some((n) => n.indexOf('no text block detected') >= 0),
    'notes must say nothing was found, got ' + JSON.stringify(ui.stats.notes),
  );
  assertPixelExact(FIX.ui640, ui);

  const tall = sceneBuild(FIX.textAA, { shapeBudget: 0, textRegions: true, textMinHeightPx: 100 }).scene;
  assert.equal(textsOf(tall).length, 0, 'textMinHeightPx 100 must reject 12px glyphs');
});

/* ------------------------------------------------------------------ *
 * 7. edge cases                                                        *
 * ------------------------------------------------------------------ */

test('edge cases: 1x1, transparent, single colour, custom background, slide cap, bad input', () => {
  for (const mode of ['smart', 'vector', 'image']) {
    const one = sceneBuild(FIX.onePixel, { mode }).scene;
    assert.equal(one.width, 1);
    assert.equal(one.height, 1);
    assert.ok(one.shapes.length >= 1);
    assertPixelExact(FIX.onePixel, one);
    assertNoSeam(one, '1x1 ' + mode);
    assertShapeGeometry(one);
  }

  // fully transparent source: composites to the background colour, reported in the notes
  const clear = sceneBuild(FIX.transparent, {}).scene;
  assert.equal(clear.shapes.length, 1);
  assert.equal(clear.shapes[0].c, '#ffffff');
  assertPixelExact(FIX.transparent, clear);
  assert.ok(
    clear.stats.notes.some((n) => n.indexOf('transparent') >= 0),
    'notes must report the composited transparent pixels',
  );

  // semi-transparent pixels are alpha-composited, not dropped
  const semi = sceneBuild(FIX.mostlyTransparent, { mode: 'vector', shapeBudget: 0 }).scene;
  assertPixelExact(FIX.mostlyTransparent, semi);
  assertShapeGeometry(semi);

  // custom alphaBackground changes the composite target
  const pink = sceneBuild(FIX.transparent, { alphaBackground: '#102030' }).scene;
  assert.equal(pink.background, '#102030');
  assert.equal(pink.shapes[0].c, '#102030');
  assertPixelExact(FIX.transparent, pink);

  // maxSlidePx clamps the px -> slide scale but leaves the scene at source resolution
  const wide = sceneBuild(FIX.wideSolid, { maxSlidePx: 5376 });
  assert.equal(wide.scene.width, 6000);
  assert.equal(wide.scale, 5376 / 6000);
  assert.ok(wide.scene.stats.notes.some((n) => n.indexOf('slide scale') >= 0));

  // degenerate / hostile input must not throw
  for (const bad of [
    sceneBuild({}).scene,
    sceneBuild({ width: 0, height: 0, data: new Uint8ClampedArray(0) }).scene,
    sceneBuild({ width: 4, height: 4, data: new Uint8ClampedArray(8) }).scene,
    sceneBuild(FIX.ui640, null).scene,
    sceneBuild(FIX.ui640, { mode: 'bogus', maxColors: -3, mergeTolerance: 'x', shapeBudget: -5, textMinHeightPx: -2, regionMinPx: 0 }).scene,
  ]) {
    assert.ok(bad.width >= 1 && bad.height >= 1, 'a scene is always at least 1x1');
    assert.ok(bad.shapes.length >= 1, 'a degenerate scene still covers its pixels');
    assertShapeGeometry(bad);
    assertNoSeam(bad, 'degenerate input');
  }
});

/* ------------------------------------------------------------------ *
 * 8. stats / diff API                                                  *
 * ------------------------------------------------------------------ */

test('sceneRasterize / sceneDiff contract and stats consistency', () => {
  const { scene } = sceneBuild(FIX.ui640, { shapeBudget: 0, textRegions: true });
  const raster = sceneRasterize(scene);
  assert.equal(raster.width, scene.width);
  assert.equal(raster.height, scene.height);
  assert.equal(raster.data.length, scene.width * scene.height * 4);
  assert.ok(raster.data instanceof Uint8ClampedArray);

  const d = diffVsSource(FIX.ui640, scene);
  assert.equal(d.pixels, scene.width * scene.height);
  assert.equal(d.heatmap.length, d.pixels * 4);
  assert.equal(scene.stats.shapes, scene.shapes.length);
  assert.equal(scene.stats.rects, rectsOf(scene).length);
  assert.equal(scene.stats.images, cropsOf(scene).length);
  assert.equal(scene.stats.texts, textsOf(scene).length);
  assert.equal(scene.stats.exactRatio, d.exactRatio);
  assert.equal(scene.stats.mismatchPixels, d.mismatchPixels);
  assert.ok(scene.stats.maxAbs >= d.maxAbs - 1e-9);
  assert.ok(scene.stats.ms >= 0);
  assertPalette(scene);
  assertShapeGeometry(scene);

  // identical rasters: no mismatch, empty heatmap, no ignored denominator
  const same = sceneDiff(raster, raster);
  assert.equal(same.mismatchPixels, 0);
  assert.equal(same.exactRatio, 1);
  assert.equal(same.exactRatioIgnoring, null, 'null without ignoreRects');
  assert.ok(same.heatmap.every((v) => v === 0), 'identical rasters have a transparent heatmap');

  // a real difference is reported in the heatmap (red, alpha scaled by delta)
  const changed = { width: 2, height: 1, data: new Uint8ClampedArray([0, 0, 0, 255, 255, 255, 255, 255]) };
  const other = { width: 2, height: 1, data: new Uint8ClampedArray([0, 0, 0, 255, 10, 10, 10, 255]) };
  const delta = sceneDiff(changed, other);
  assert.equal(delta.pixels, 2);
  assert.equal(delta.mismatchPixels, 1);
  assert.equal(delta.exactRatio, 0.5);
  assert.equal(delta.maxAbs, 245);
  assert.equal(delta.heatmap[4], 255, 'mismatch pixel is red');
  assert.ok(delta.heatmap[7] > 0, 'mismatch pixel alpha encodes the delta');
  assert.equal(delta.heatmap[0], 0, 'exact pixel has no heat');

  // ignoreRects excludes pixels from exactRatioIgnoring only
  const ignored = sceneDiff(changed, other, { ignoreRects: [{ x: 1, y: 0, w: 1, h: 1 }] });
  assert.equal(ignored.exactRatio, 0.5, 'plain ratio still sees the pixel');
  assert.equal(ignored.exactRatioIgnoring, 1);
  const allIgnored = sceneDiff(changed, other, { ignoreRects: [{ x: 0, y: 0, w: 2, h: 1 }] });
  assert.equal(allIgnored.exactRatioIgnoring, null, 'null when every pixel is ignored');
});

/* ------------------------------------------------------------------ *
 * 9. performance                                                       *
 * ------------------------------------------------------------------ */

test('performance: 1600x900 screenshot under 4s, 1920x1080 photo stays a handful of shapes', () => {
  const t0 = performance.now();
  const { scene } = sceneBuild(FIX.ui1600, {});
  const ms = performance.now() - t0;
  assertPixelExact(FIX.ui1600, scene);
  assert.ok(scene.stats.shapes <= sceneDefaults.shapeBudget, 'shapes ' + scene.stats.shapes + ' must respect the 8000 budget');
  assert.ok(ms < 4000, '1600x900 screenshot took ' + ms.toFixed(0) + 'ms (target < 4000ms)');
  console.log(
    '  [perf] 1600x900 UI, smart defaults: ' +
      ms.toFixed(0) +
      'ms, ' +
      scene.stats.shapes +
      ' shapes (' +
      scene.stats.rects +
      ' rects, ' +
      scene.stats.images +
      ' crops), exactRatio=' +
      scene.stats.exactRatio,
  );

  const t1 = performance.now();
  const big = sceneBuild(FIX.photo1920, {}).scene;
  const msBig = performance.now() - t1;
  assertPixelExact(FIX.photo1920, big);
  assert.ok(big.stats.shapes <= 5, 'a 1920x1080 photo must not explode: ' + big.stats.shapes + ' shapes');
  assert.ok(msBig < 4000, '1920x1080 photo took ' + msBig.toFixed(0) + 'ms (target < 4000ms)');
  console.log(
    '  [perf] 1920x1080 photo, smart defaults: ' +
      msBig.toFixed(0) +
      'ms, ' +
      big.stats.shapes +
      ' shape(s), exactRatio=' +
      big.stats.exactRatio,
  );
});

/* ------------------------------------------------------------------ *
 * 10. cross-mode invariant                                             *
 * ------------------------------------------------------------------ */

test('every mode obeys the geometry contract on a mixed screenshot', () => {
  for (const options of [
    { mode: 'smart', shapeBudget: 0 },
    { mode: 'smart', shapeBudget: 120 },
    { mode: 'vector', shapeBudget: 0 },
    { mode: 'image' },
    { mode: 'smart', shapeBudget: 0, textRegions: true, maxColors: 16 },
  ]) {
    const { scene } = sceneBuild(FIX.ui640, options);
    assertShapeGeometry(scene);
    assertNoSeam(scene, JSON.stringify(options));
    assertPalette(scene);
    const d = diffVsSource(FIX.ui640, scene, { ignoreRects: ignoreRectsOf(scene) });
    assert.equal(d.exactRatioIgnoring, 1, 'rebuild must be exact outside ignored text: ' + JSON.stringify(options));
    console.log(
      '  [mode] ' +
        JSON.stringify(options) +
        ' -> ' +
        scene.stats.shapes +
        ' shapes, exact=' +
        scene.stats.exactRatio.toFixed(4),
    );
  }
});
