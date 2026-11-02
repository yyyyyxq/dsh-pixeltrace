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
