#!/usr/bin/env node
/**
 * End-to-end verification for the PixelTrace engine, outside the browser.
 *
 *   node tools/verify.mjs                 # full pipeline: synthesise -> build -> pptx -> validate
 *   node tools/verify.mjs compare A.png B.png   # pixel diff of two PNG files
 *
 * The pipeline deliberately feeds the engine a hand-drawn RGBA raster (no decoder
 * needed), checks that the shape list reproduces every pixel, writes the source
 * PNG next to the .pptx, and validates the package with the independent
 * @office-kit/pptx reader. out/verify/verify.pptx is the file to hand to
 * LibreOffice (ppt_render) for the real rendering comparison.
 */
import { mkdirSync, writeFileSync, readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { loadEngine, repoRoot } from '../test/harness.mjs';

const OFFICE_KIT = 'file:///C:/Users/yxq/.dsh/profiles/web/node_modules/@office-kit/pptx/dist/index.js';
const OUT = join(repoRoot, 'out', 'verify');

const engine = loadEngine();

/** Small deterministic RGBA painter, enough to fake flat UI + photo regions. */
function painter(width, height) {
  const data = new Uint8ClampedArray(width * height * 4);
  const put = (x, y, r, g, b, a = 255) => {
    if (x < 0 || y < 0 || x >= width || y >= height) return;
    const p = (y * width + x) * 4;
    data[p] = r; data[p + 1] = g; data[p + 2] = b; data[p + 3] = a;
  };
  const rect = (x0, y0, w, h, r, g, b) => {
    for (let y = y0; y < y0 + h; y++) for (let x = x0; x < x0 + w; x++) put(x, y, r, g, b);
  };
  return { data, width, height, put, rect };
}

function drawUi(width = 640, height = 400) {
  const img = painter(width, height);
  img.rect(0, 0, width, height, 255, 255, 255);
  img.rect(0, 0, width, 48, 27, 79, 156);                 // header bar
  img.rect(0, 48, 200, height - 48, 238, 242, 248);       // side panel
  for (let i = 0; i < 6; i++) img.rect(16, 70 + i * 34, 168, 18, 34, 48, 63); // sidebar text rows
  for (let y = 96; y < 240; y++) {                        // diagonal gradient = "photo"
    for (let x = 224; x < 600; x++) {
      const t = (x - 224) / 376 * 0.6 + (y - 96) / 144 * 0.4;
      img.put(x, y, Math.round(255 - 200 * t), Math.round(122 + 60 * t), Math.round(24 + 180 * t));
    }
  }
  img.rect(232, 264, 300, 60, 245, 197, 24);              // yellow text block
  for (let i = 0; i < 3; i++) img.rect(240, 272 + i * 16, 220 - i * 40, 10, 17, 17, 17); // text-like bars
  for (let i = 0; i < 10; i++) img.rect(232, 340 + i * 5, 380, 2, 91, 107, 124); // hairlines
  img.rect(560, 300, 60, 60, 43, 182, 115);               // solid accent
  img.rect(0, 392, width, 8, 20, 20, 20);                 // footer
  return img;
}

function drawFlat() {
  const img = painter(40, 24);
  img.rect(0, 0, 40, 16, 12, 34, 56);
  img.rect(0, 16, 40, 8, 240, 240, 240);
  return img;
}

/** Many distinct flat colours (24 bands) next to one gradient "photo" region. */
function drawBands() {
  const img = painter(240, 160);
  const palette = [];
  for (let i = 0; i < 24; i++) {
    palette.push([(i * 47 + 11) % 256, (i * 89 + 29) % 256, (i * 131 + 7) % 256]);
  }
  for (let i = 0; i < 24; i++) img.rect((i % 8) * 30, Math.floor(i / 8) * 20, 30, 20, ...palette[i]);
  for (let y = 60; y < 160; y++) {
    for (let x = 0; x < 240; x++) {
      img.put(x, y, Math.round(255 - x * 0.7), Math.round(40 + y * 0.9), Math.round(120 + (x + y) * 0.3));
    }
  }
  return img;
}

function imageOf(paint) {
  return { width: paint.width, height: paint.height, data: paint.data };
}

function stats(scene) {
  const s = scene && scene.stats ? scene.stats : {};
  return {
    shapes: s.shapes ?? (scene.shapes || []).length,
    rects: s.rects,
    images: s.images,
    texts: s.texts,
    paletteSize: s.paletteSize,
    exactRatio: s.exactRatio,
    mismatchPixels: s.mismatchPixels,
    meanAbs: s.meanAbs,
    maxAbs: s.maxAbs,
    ms: s.ms,
    notes: s.notes ?? [],
  };
}

/**
 * Count mismatched pixels strictly inside the given rects. This is the honest
 * way to prove that crops survive palette quantisation: the rest of the image
 * is allowed to change, the crop pixels are not.
 */
function cropMismatch(source, rebuilt, rects) {
  let bad = 0;
  let total = 0;
  for (const rect of rects) {
    const x0 = Math.max(0, Math.round(rect.x));
    const y0 = Math.max(0, Math.round(rect.y));
    const x1 = Math.min(source.width, Math.round(rect.x + rect.w));
    const y1 = Math.min(source.height, Math.round(rect.y + rect.h));
    for (let y = y0; y < y1; y++) {
      for (let x = x0; x < x1; x++) {
        const index = (y * source.width + x) * 4;
        total++;
        for (let channel = 0; channel < 4; channel++) {
          if (source.data[index + channel] !== rebuilt.data[index + channel]) {
            bad++;
            break;
          }
        }
      }
    }
  }
  return { bad, total };
}

async function compare(job, options) {
  const source = imageOf(job.paint);
  const { scene } = engine.sceneBuild(source, options);
  const rebuilt = engine.sceneRasterize(scene);
  const diff = engine.sceneDiff(source, rebuilt);
  return { job, options, scene, diff, summary: stats(scene) };
}

function line(name, ok, detail) {
  const mark = ok === true ? 'PASS' : ok === false ? 'FAIL' : 'INFO';
  console.log(`  [${mark}] ${name}${detail === undefined ? '' : ' — ' + detail}`);
  return ok;
}

async function runPipeline() {
  mkdirSync(OUT, { recursive: true });
  const ui = drawUi();
  const flat = drawFlat();
  const report = { generatedAt: new Date().toISOString(), checks: [], cases: [] };
  const record = (name, ok, detail) => {
    report.checks.push({ name, ok, detail });
    line(name, ok, detail);
  };

  console.log('PixelTrace engine verification');
  console.log(`  harness: ${repoRoot}`);

  // ---- case 1: solid flat image, vector mode must be exact with a couple of rects
  {
    const run = await compare({ name: 'flat', paint: flat }, { mode: 'vector' });
    console.log(`\n· flat 40x24 vector`);
    record('flat/vector shapes small', run.summary.shapes <= 8, `${run.summary.shapes} shape(s)`);
    record('flat/vector exact', run.diff.exactRatio === 1, `exactRatio ${run.diff.exactRatio}`);
    report.cases.push({ name: run.job.name, mode: 'vector', summary: run.summary, exactRatio: run.diff.exactRatio });
  }

  // ---- case 2: smart mode on a UI-like picture with a photo region
  {
    const started = Date.now();
    const run = await compare({ name: 'ui', paint: ui }, { mode: 'smart' });
    const wall = Date.now() - started;
    console.log(`\n· ui 640x400 smart`);
    record('ui/smart exact', run.diff.exactRatio === 1, `exactRatio ${run.diff.exactRatio}, mismatch ${run.diff.mismatchPixels}`);
    record('ui/smart has editable rects', (run.summary.rects ?? 0) > 0, `${run.summary.rects} rect(s)`);
    record('ui/smart keeps the photo as crop', (run.summary.images ?? 0) > 0, `${run.summary.images} image(s)`);
    record('ui/smart budget respected', run.summary.shapes <= 8000, `${run.summary.shapes} shape(s)`);
    record('ui/smart runtime under 4s', wall < 4000, `${wall} ms wall`);
    report.cases.push({ name: run.job.name, mode: 'smart', summary: run.summary, exactRatio: run.diff.exactRatio, wall });
  }

  // ---- case 3: whole-image mode is a single picture and exactly reproducible
  {
    const run = await compare({ name: 'ui-image', paint: ui }, { mode: 'image' });
    console.log(`\n· ui 640x400 image`);
    record('ui/image single picture', run.summary.shapes === 1 && run.summary.images === 1, `${run.summary.shapes} shape(s)`);
    record('ui/image exact', run.diff.exactRatio === 1, `exactRatio ${run.diff.exactRatio}`);
    report.cases.push({ name: run.job.name, mode: 'image', summary: run.summary, exactRatio: run.diff.exactRatio });
  }

  // ---- case 4: vector only, budget pressure and degradation notes
  {
    const run = await compare({ name: 'ui-vector', paint: ui }, { mode: 'vector' });
    console.log(`\n· ui 640x400 vector`);
    const vectorNote = run.diff.exactRatio === 1 ? 'lossless with rects only' : `lossy as expected (${run.diff.exactRatio})`;
    record('ui/vector reports its fidelity', true, vectorNote);
    const budget = await compare({ name: 'ui-budget', paint: ui }, { mode: 'vector', shapeBudget: 24 });
    console.log(`\n· ui 640x400 vector, budget 24`);
    record('budget note recorded', (budget.summary.notes || []).length > 0, JSON.stringify(budget.summary.notes).slice(0, 200));
    record('budget respected', budget.summary.shapes <= 24, `${budget.summary.shapes} shape(s)`);
    report.cases.push({ name: run.job.name, mode: 'vector', summary: run.summary, exactRatio: run.diff.exactRatio });
    report.cases.push({ name: budget.job.name, mode: 'vector', summary: budget.summary, exactRatio: budget.diff.exactRatio });
  }

  // ---- case 5: palette quantisation keeps crops exact
  {
    const run = await compare({ name: 'ui-quantised', paint: ui }, { mode: 'smart', maxColors: 16 });
    console.log(`\n· ui 640x400 smart, maxColors 16`);
    record('quantised palette size', (run.summary.paletteSize ?? 0) <= 16, `${run.summary.paletteSize} colour(s)`);
    // Flat colours are re-quantised on purpose; only the picture crops must stay exact.
    const crops = (run.scene.shapes || []).filter((shape) => shape.k === 'i');
    const rebuilt = engine.sceneRasterize(run.scene);
    const inside = cropMismatch(imageOf(ui), rebuilt, crops);
    record('quantised crops stay exact', inside.bad === 0, `${crops.length} crop(s), ${inside.bad}/${inside.total} crop pixel(s) changed`);
    // The test image happens to fit inside 16 colours, so quantisation must be a
    // no-op here; the lossy case is checked separately on a 24-colour image.
    const paletteFits = (run.summary.paletteSize ?? 0) <= 16;
    record(
      'quantised flat colours',
      paletteFits ? run.diff.exactRatio === 1 : run.diff.exactRatio < 1,
      `${run.summary.paletteSize} colour(s), whole-image exactRatio ${run.diff.exactRatio}`,
    );
    report.cases.push({ name: run.job.name, mode: 'smart', summary: run.summary, exactRatio: run.diff.exactRatio });
  }

  // ---- case 5b: quantisation really bites, crops still stay exact
  {
    const bands = drawBands();
    const run = await compare({ name: 'bands-quantised', paint: bands }, { mode: 'smart', maxColors: 8 });
    console.log(`\n· bands 240x160 smart, maxColors 8`);
    record('bands palette clamped', (run.summary.paletteSize ?? 0) <= 8, `${run.summary.paletteSize} colour(s)`);
    const crops = (run.scene.shapes || []).filter((shape) => shape.k === 'i');
    record('bands has exact crops', crops.length > 0, `${crops.length} crop(s)`);
    const rebuilt = engine.sceneRasterize(run.scene);
    const inside = cropMismatch(imageOf(bands), rebuilt, crops);
    record('bands crops stay exact', inside.bad === 0, `${crops.length} crop(s), ${inside.bad}/${inside.total} crop pixel(s) changed`);
    record('bands flats quantised', run.diff.exactRatio < 1, `whole-image exactRatio ${run.diff.exactRatio}`);
    report.cases.push({ name: run.job.name, mode: 'smart', summary: run.summary, exactRatio: run.diff.exactRatio });
  }

  // ---- case 6: PNG round trip
  {
    const source = imageOf(ui);
    const png = await engine.pptxEncodePng(source.data, source.width, source.height);
    const back = await engine.pptxDecodePng(png);
    const diff = engine.sceneDiff(source, back);
    writeFileSync(join(OUT, 'scene-source.png'), Buffer.from(png));
    console.log(`\n· png round trip (${png.byteLength} bytes)`);
    record('png round trip exact', diff.exactRatio === 1, `exactRatio ${diff.exactRatio}`);
  }

  // ---- case 7: build a two-slide deck and validate it with @office-kit/pptx
  {
    const smartUi = engine.sceneBuild(imageOf(ui), { mode: 'smart' }).scene;
    const flatScene = engine.sceneBuild(imageOf(flat), { mode: 'vector' }).scene;
    const started = Date.now();
    const bytes = await engine.pptxBuild([
      { scene: smartUi, name: 'screenshot' },
      { scene: flatScene, name: 'flat' },
    ], { title: 'PixelTrace verification', creator: 'PixelTrace (dsh-pixeltrace)' });
    const wall = Date.now() - started;
    const file = join(OUT, 'verify.pptx');
    writeFileSync(file, Buffer.from(bytes));
    console.log(`\n· deck: ${file} (${bytes.byteLength} bytes, ${wall} ms)`);
    record('deck written', existsSync(file), `${bytes.byteLength} bytes in ${wall} ms`);

    const unzipped = await engine.pptxUnzip(bytes);
    const names = Array.from(unzipped.keys());
    const required = [
      '[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml',
      'ppt/slideMasters/slideMaster1.xml', 'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/theme/theme1.xml',
    ];
    const missing = required.filter((name) => !names.includes(name));
    record('ooxml parts present', missing.length === 0, missing.length ? `missing ${missing.join(', ')}` : `${names.length} entries`);

    try {
      const office = await import(OFFICE_KIT);
      const presentation = await office.loadPresentation(new Uint8Array(readFileSync(file)));
      const validation = (office.validatePresentation(presentation) ?? []).map((issue) => String(issue.message ?? issue));
      const slides = office.getSlides(presentation);
      record('office-kit loads deck', true, `${slides.length} slide(s)`);
      record('office-kit validation clean', validation.length === 0, validation.slice(0, 6).join(' | ') || 'no issues');
      const shapeCounts = slides.map((slide) => office.getSlideShapes(slide).length);
      record('office-kit sees shapes', shapeCounts.every((count) => count > 0), JSON.stringify(shapeCounts));
      report.ooxml = { slides: slides.length, shapeCounts, issues: validation };
    } catch (error) {
      record('office-kit loads deck', false, String(error?.message ?? error));
    }
  }

  writeFileSync(join(OUT, 'report.json'), JSON.stringify(report, null, 2));
  const failed = report.checks.filter((check) => check.ok === false);
  console.log(`\n${failed.length === 0 ? 'ALL CHECKS PASSED' : `${failed.length} CHECK(S) FAILED`} — report: ${join(OUT, 'report.json')}`);
  if (failed.length > 0) process.exitCode = 1;
}

async function readAsImage(file) {
  const bytes = new Uint8Array(readFileSync(file));
  return await engine.pptxDecodePng(bytes);
}

async function runCompare(aPath, bPath) {
  const a = await readAsImage(aPath);
  const b = await readAsImage(bPath);
  const diff = engine.sceneDiff(a, b);
  console.log(JSON.stringify({
    a: { file: aPath, width: a.width, height: a.height },
    b: { file: bPath, width: b.width, height: b.height },
    exactRatio: diff.exactRatio,
    mismatchPixels: diff.mismatchPixels,
    meanAbs: diff.meanAbs,
    maxAbs: diff.maxAbs,
  }, null, 2));
  return diff;
}

/** Write an amplified difference map so a human can see where the pixels went. */
async function runDiffmap(aPath, bPath, outPath) {
  const a = await readAsImage(aPath);
  const b = await readAsImage(bPath);
  const diff = engine.sceneDiff(a, b);
  const source = diff.heatmap instanceof Uint8ClampedArray ? diff.heatmap : null;
  if (!source) throw new Error('sceneDiff did not return a heatmap');
  const png = await engine.pptxEncodePng(source, a.width, a.height);
  writeFileSync(outPath, png);
  console.log(JSON.stringify({
    a: aPath,
    b: bPath,
    out: outPath,
    exactRatio: diff.exactRatio,
    mismatchPixels: diff.mismatchPixels,
    meanAbs: diff.meanAbs,
    maxAbs: diff.maxAbs,
  }, null, 2));
}

const argv = process.argv.slice(2);
if (argv[0] === 'compare') {
  if (!argv[1] || !argv[2]) {
    console.error('usage: node tools/verify.mjs compare <a.png> <b.png>');
    process.exit(1);
  }
  await runCompare(argv[1], argv[2]);
} else if (argv[0] === 'diffmap') {
  if (!argv[1] || !argv[2] || !argv[3]) {
    console.error('usage: node tools/verify.mjs diffmap <a.png> <b.png> <out.png>');
    process.exit(1);
  }
  await runDiffmap(argv[1], argv[2], argv[3]);
} else {
  await runPipeline();
}
