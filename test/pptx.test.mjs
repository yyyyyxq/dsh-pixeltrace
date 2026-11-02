// Tests for the PresentationML writer fragment (src/pptx.frag.js).
//
// Covers CONTRACT.md section 7 acceptance criteria for pptx-dev:
//   * ZIP round-trip, header consistency, CRC-32 and fixed DOS timestamps
//   * PNG encode -> decode equality and determinism
//   * XML shape mapping: exact EMU geometry, ids, names, fills, no outlines,
//     cropped pictures, text boxes, package/relationship/content-type integrity
//   * determinism (same input -> identical bytes) and scene non-mutation
//   * external validation with @office-kit/pptx (a real OOXML library that is
//     intentionally NOT a dependency of this repo; resolved from the harness
//     profile only, and skipped when it is unavailable)
//   * writes the demo deck out/pptx-dev-demo.pptx for the Lead to render
//
// Run: node --test test/pptx.test.mjs   (from D:\pixeltrace)

import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { posix as posixPath } from 'node:path';
import { test } from 'node:test';

import { loadPptx, repoRoot } from './harness.mjs';

const { pptxBuild, pptxEncodePng, pptxDecodePng, pptxZip, pptxUnzip } = loadPptx();

const OUT_DIR = path.join(repoRoot, 'out');
const DEMO_PATH = path.join(OUT_DIR, 'pptx-dev-demo.pptx');
const OFFICE_KIT_URL =
  'file:///C:/Users/yxq/.dsh/profiles/web/node_modules/@office-kit/pptx/dist/index.js';
const EMU = 9525; // 1 px = 9525 EMU at scale 1 (96 dpi)

/* ------------------------------------------------------------------ helpers */

const decoder = new TextDecoder('utf-8');

function crc32(bytes, start = 0, end = bytes.length) {
  // Independent CRC-32 (ZIP/PNG polynomial) used to verify the fragment's own.
  let crc = -1;
  for (let i = start; i < end; i++) {
    crc ^= bytes[i];
    for (let bit = 0; bit < 8; bit++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
  }
  return (crc ^ -1) >>> 0;
}

function u8(value) {
  return value instanceof Uint8Array ? value : new Uint8Array(value);
}

function text(bytes) {
  return decoder.decode(u8(bytes));
}

function readPart(parts, name) {
  const bytes = parts.get(name);
  assert.ok(bytes, `package is missing part ${name}`);
  return text(bytes);
}

/** Parse the central directory and the matching local headers of a ZIP archive. */
function parseZipDirectory(bytes) {
  const bin = u8(bytes);
  const view = new DataView(bin.buffer, bin.byteOffset, bin.byteLength);
  let eocd = -1;
  for (let i = bin.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  assert.ok(eocd >= 0, 'end of central directory record not found');
  const count = view.getUint16(eocd + 10, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  assert.equal(view.getUint16(eocd + 8, true), count, 'entry count on disk must match total');
  const entries = [];
  let position = centralOffset;
  for (let i = 0; i < count; i++) {
    assert.equal(view.getUint32(position, true), 0x02014b50, 'central directory signature');
    const flags = view.getUint16(position + 8, true);
    const method = view.getUint16(position + 10, true);
    const time = view.getUint16(position + 12, true);
    const date = view.getUint16(position + 14, true);
    const crc = view.getUint32(position + 16, true);
    const compressedSize = view.getUint32(position + 20, true);
    const size = view.getUint32(position + 24, true);
    const nameLength = view.getUint16(position + 28, true);
    const extraLength = view.getUint16(position + 30, true);
    const commentLength = view.getUint16(position + 32, true);
    const offset = view.getUint32(position + 42, true);
    assert.equal(view.getUint32(offset, true), 0x04034b50, 'local file header signature');
    const name = text(bin.subarray(position + 46, position + 46 + nameLength));
    const localNameLength = view.getUint16(offset + 26, true);
    const localExtraLength = view.getUint16(offset + 28, true);
    const localFlags = view.getUint16(offset + 6, true);
    const localMethod = view.getUint16(offset + 8, true);
    const dataStart = offset + 30 + localNameLength + localExtraLength;
    assert.ok(dataStart + compressedSize <= bin.length, `data out of range for ${name}`);
    entries.push({
      name,
      flags,
      method,
      time,
      date,
      crc,
      compressedSize,
      size,
      offset,
      localFlags,
      localMethod,
      data: bin.subarray(dataStart, dataStart + compressedSize),
    });
    position += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** Parse PNG chunks and verify every chunk CRC (independent implementation). */
function parsePng(bytes) {
  const bin = u8(bytes);
  assert.deepEqual([...bin.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const view = new DataView(bin.buffer, bin.byteOffset, bin.byteLength);
  const chunks = [];
  let position = 8;
  while (position + 12 <= bin.length) {
    const length = view.getUint32(position, false);
    const type = text(bin.subarray(position + 4, position + 8));
    const data = bin.subarray(position + 8, position + 8 + length);
    const stored = view.getUint32(position + 8 + length, false);
    assert.equal(crc32(bin, position + 4, position + 8 + length), stored, `CRC of ${type} chunk`);
    chunks.push({ type, data });
    position += 12 + length;
    if (type === 'IEND') break;
  }
  return chunks;
}

function sceneOf(width, height, background, shapes) {
  return { width, height, background, shapes, palette: [], stats: {} };
}

function cropOf(width, height, pixel) {
  const data = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const [r, g, b, a = 255] = pixel(x, y);
      const i = (y * width + x) * 4;
      data[i] = r;
      data[i + 1] = g;
      data[i + 2] = b;
      data[i + 3] = a;
    }
  }
  return { data, width, height };
}

const checker = (x, y) => (((x >> 2) + (y >> 2)) % 2 === 0 ? [255, 140, 0, 255] : [20, 20, 20, 255]);
const gradient = (x, y) => [x * 6, y * 6, 255 - x * 6, 255 - (y % 2) * 40];

/** Shapes of one <p:spTree>, in document order (z-order). */
function parseShapes(slideXml) {
  const shapes = [];
  const block = /<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g;
  let match;
  while ((match = block.exec(slideXml))) {
    const body = match[2];
    const cNvPr = /<p:cNvPr id="(\d+)" name="([^"]*)"/.exec(body);
    assert.ok(cNvPr, 'shape without <p:cNvPr>');
    const xfrm =
      /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(body);
    assert.ok(xfrm, `shape ${cNvPr[1]} without <a:xfrm>`);
    shapes.push({
      kind: match[1] === 'pic' ? 'picture' : 'shape',
      id: Number(cNvPr[1]),
      name: cNvPr[2],
      x: Number(xfrm[1]),
      y: Number(xfrm[2]),
      w: Number(xfrm[3]),
      h: Number(xfrm[4]),
      body,
    });
  }
  return shapes;
}

function resolveRelTarget(relsName, target) {
  const dir = posixPath.dirname(relsName); // e.g. 'ppt/slides/_rels'
  const base = posixPath.dirname(dir); // e.g. 'ppt/slides'
  if (target.startsWith('/')) return target.slice(1);
  return posixPath.normalize(posixPath.join(base, target));
}

/**
 * Package integrity: content types cover every part, every relationship target
 * exists, and every r:id/r:embed used by a part is declared in its .rels.
 */
function assertPackageIntegrity(parts) {
  assert.ok(parts.has('[Content_Types].xml'), 'missing [Content_Types].xml');
  assert.ok(parts.has('_rels/.rels'), 'missing _rels/.rels');

  const typesXml = readPart(parts, '[Content_Types].xml');
  const defaults = new Map();
  const overrides = new Map();
  for (const match of typesXml.matchAll(/<Default Extension="([^"]+)" ContentType="([^"]+)"\/>/g)) {
    defaults.set(match[1].toLowerCase(), match[2]);
  }
  for (const match of typesXml.matchAll(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g)) {
    overrides.set(match[1].replace(/^\//, ''), match[2]);
    assert.ok(parts.has(match[1].replace(/^\//, '')), `dangling content type override ${match[1]}`);
  }
  for (const name of parts.keys()) {
    if (name === '[Content_Types].xml') continue;
    const direct = overrides.get(name);
    const extension = name.includes('.') ? name.slice(name.lastIndexOf('.') + 1).toLowerCase() : '';
    const byExtension = defaults.get(extension);
    assert.ok(direct || byExtension, `part ${name} has no content type`);
    if (direct) assert.ok(direct.length > 0);
  }

  // Every .rels part: unique ids, resolvable targets.
  const relsBySource = new Map();
  for (const [name, bytes] of parts) {
    if (!name.endsWith('.rels')) continue;
    const relsXml = text(bytes);
    let source;
    const nested = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(name);
    if (nested) source = `${nested[1]}/${nested[2]}`;
    else if (name === '_rels/.rels') source = '';
    else throw new Error(`unexpected rels location ${name}`);
    const declared = new Map();
    for (const match of relsXml.matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"\/>/g)) {
      const id = match[1];
      assert.ok(!declared.has(id), `duplicate relationship id ${id} in ${name}`);
      const target = resolveRelTarget(name, match[3]);
      assert.ok(parts.has(target), `relationship ${id} in ${name} points at missing part ${target}`);
      declared.set(id, target);
    }
    assert.ok(declared.size > 0, `empty relationships part ${name}`);
    if (source !== '') {
      assert.ok(parts.has(source), `relationships part ${name} has no source part ${source}`);
      relsBySource.set(source, { relsName: name, declared });
    }
  }

  // Every r:id/r:embed etc. used inside a part must be declared in its rels.
  const attr = /\sr:(id|embed|link|pict|dm|lo|qs|cs)="([^"]+)"/g;
  for (const [name, bytes] of parts) {
    if (!name.endsWith('.xml')) continue;
    const known = relsBySource.get(name);
    for (const match of text(bytes).matchAll(attr)) {
      assert.ok(known, `part ${name} uses r:${match[1]} but has no relationships part`);
      assert.ok(
        known.declared.has(match[2]),
        `part ${name} references undeclared ${match[2]} (r:${match[1]})`,
      );
    }
  }
  return { defaults, overrides };
}

/** Presentation-level ids that strict consumers (PowerPoint) check. */
function assertDeckIds(presentationXml, slideCount) {
  const master = /<p:sldMasterId id="(\d+)" r:id="([^"]+)"\/>/.exec(presentationXml);
  assert.ok(master, 'missing <p:sldMasterId>');
  const masterId = Number(master[1]);
  assert.ok(masterId >= 2147483648, `slide master id ${masterId} out of range`);
  assert.equal(master[2], 'rId1');
  const ids = [...presentationXml.matchAll(/<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g)].map((m) => Number(m[1]));
  assert.equal(ids.length, slideCount);
  for (let i = 0; i < ids.length; i++) {
    assert.ok(ids[i] >= 256 && ids[i] <= 2147483647, `slide id ${ids[i]} out of range`);
    assert.equal(ids[i], 256 + i);
  }
  const layout = /<p:sldLayoutId id="(\d+)" r:id="([^"]+)"\/>/.exec(
    presentationXml.includes('sldLayoutId') ? presentationXml : '',
  );
  // The layout id lives in the slide master, checked separately where read.
  void layout;
}

/* ------------------------------------------------------------------- tests */

test('pptxZip round-trips files, writes consistent headers and fixed timestamps', async () => {
  const files = [
    { name: '[Content_Types].xml', data: text('hello hello hello hello hello') },
    { name: 'ppt/slides/slide1.xml', data: new Uint8Array(4096).fill(7) },
    { name: 'empty.bin', data: new Uint8Array(0) },
    { name: 'unicode-名字.txt', data: text('utf-8 name') },
  ];
  const bytes = await pptxZip(files);
  const parts = await pptxUnzip(bytes);
  assert.deepEqual([...parts.keys()], files.map((f) => f.name));
  for (const file of files) {
    assert.deepEqual([...parts.get(file.name)], [...u8(typeof file.data === 'string' ? new TextEncoder().encode(file.data) : file.data)]);
  }

  const entries = parseZipDirectory(bytes);
  assert.equal(entries.length, files.length);
  for (const entry of entries) {
    const source = files.find((file) => file.name === entry.name);
    assert.ok(source, `unexpected zip entry ${entry.name}`);
    const sourceBytes =
      typeof source.data === 'string' ? new TextEncoder().encode(source.data) : u8(source.data);
    assert.equal(entry.localMethod, entry.method, `local/central method mismatch for ${entry.name}`);
    assert.equal(entry.localFlags, entry.flags, `local/central flag mismatch for ${entry.name}`);
    assert.equal(entry.time, 0, 'DOS time must be fixed at 0');
    assert.equal(entry.date, 0x0021, 'DOS date must be fixed at 1980-01-01');
    assert.equal(entry.crc, crc32(sourceBytes), `central CRC for ${entry.name}`);
    assert.equal(entry.size, sourceBytes.length);
    assert.equal(entry.compressedSize, entry.data.length);
    assert.ok(entry.method === 0 || entry.method === 8);
    assert.equal(entry.flags & 0x0800, entry.name.includes('unicode') ? 0x0800 : 0);
  }
  // Entries big enough to matter are actually deflated.
  const big = entries.find((e) => e.name === 'ppt/slides/slide1.xml');
  assert.equal(big.method, 8);
  assert.ok(big.compressedSize < big.size);

  const again = await pptxZip(files);
  assert.deepEqual([...again], [...bytes], 'pptxZip must be deterministic');

  await assert.rejects(() => pptxUnzip(new Uint8Array(64)), /ZIP|central|signature/);
  await assert.rejects(() => pptxZip([]), /nothing to pack/);
});

test('pptxEncodePng decodes back to identical RGBA and is deterministic', async () => {
  const cases = [
    { name: '1x1 opaque', width: 1, height: 1, crop: cropOf(1, 1, () => [10, 20, 30, 255]) },
    {
      name: 'gradient with alpha',
      width: 7,
      height: 5,
      crop: cropOf(7, 5, (x, y) => [x * 30, y * 40, (x + y) * 20, 255 - (x * y) % 256]),
    },
    {
      name: 'noise',
      width: 16,
      height: 16,
      crop: cropOf(16, 16, (x, y) => {
        const seed = (x * 131 + y * 197) % 251;
        return [seed, (seed * 7) % 256, (seed * 13) % 256, (seed * 3) % 256];
      }),
    },
    { name: 'transparent', width: 3, height: 2, crop: cropOf(3, 2, () => [0, 0, 0, 0]) },
  ];

  for (const item of cases) {
    const png = await pptxEncodePng(item.crop.data, item.width, item.height);
    const again = await pptxEncodePng(item.crop.data, item.width, item.height);
    assert.deepEqual([...again], [...png], `PNG encoding must be deterministic (${item.name})`);

    const chunks = parsePng(png);
    assert.equal(chunks[0].type, 'IHDR');
    assert.equal(chunks.at(-1).type, 'IEND');
    const ihdr = chunks[0].data;
    assert.equal(new DataView(ihdr.buffer, ihdr.byteOffset, ihdr.byteLength).getUint32(0, false), item.width);
    assert.equal(new DataView(ihdr.buffer, ihdr.byteOffset, ihdr.byteLength).getUint32(4, false), item.height);
    assert.equal(ihdr[8], 8, 'bit depth');
    assert.equal(ihdr[9], 6, 'colour type RGBA');
    assert.equal(ihdr[12], 0, 'non-interlaced');
    assert.ok(chunks.some((c) => c.type === 'IDAT'));

    const decoded = await pptxDecodePng(png);
    assert.equal(decoded.width, item.width);
    assert.equal(decoded.height, item.height);
    assert.deepEqual(
      [...decoded.data],
      [...item.crop.data],
      `encode -> decode must be lossless (${item.name})`,
    );
  }

  // Errors are explicit rather than silently wrong.
  await assert.rejects(() => pptxEncodePng(new Uint8Array(4), 2, 2), /needs 16 bytes/);
  await assert.rejects(() => pptxEncodePng(new Uint8Array(4), 0, 1), /positive integers/);
  await assert.rejects(() => pptxDecodePng(new Uint8Array(12)), /too short|bad PNG/);

  const png = await pptxEncodePng(cropOf(2, 2, () => [1, 2, 3, 255]).data, 2, 2);
  const corrupt = new Uint8Array(png);
  const idat = parsePng(png).find((c) => c.type === 'IDAT');
  const idatStart = png.indexOf(idat.data[0], 40);
  corrupt[idatStart + 2] ^= 0xff;
  await assert.rejects(() => pptxDecodePng(corrupt), /CRC-32 mismatch/);
});

test('pptxBuild maps shapes to exact EMU geometry (rect, cropped picture, text)', async () => {
  const image = cropOf(8, 8, checker);
  const scene = sceneOf(300, 200, '#f0f0f0', [
    { k: 'r', x: 0, y: 0, w: 300, h: 20, c: '#1f6feb', role: 'bg' },
    { k: 'r', x: 10, y: 20, w: 40, h: 30, c: '#ABCDEF', role: 'flat' },
    { k: 'i', x: 100, y: 50, w: 8, h: 8, px: image, role: 'photo' },
    { k: 't', x: 20, y: 120, w: 120, h: 20, s: 'Hello <PixelTrace> & "friends"', sz: 15, c: '#ffffff', al: 'l', b: true, f: 'Arial', role: 'text' },
  ]);
  const bytes = await pptxBuild([{ scene, name: 'page-1' }], { title: 'Map', subject: 'geometry' });
  const parts = await pptxUnzip(bytes);

  // Presentation: custom slide size, master/slide ids, layout id in the master.
  const presentation = readPart(parts, 'ppt/presentation.xml');
  const sldSz = /<p:sldSz cx="(\d+)" cy="(\d+)" type="([^"]+)"\/>/.exec(presentation);
  assert.ok(sldSz, 'missing <p:sldSz>');
  assert.deepEqual([Number(sldSz[1]), Number(sldSz[2])], [300 * EMU, 200 * EMU]);
  assert.equal(sldSz[3], 'custom');
  assertDeckIds(presentation, 1);
  assert.match(presentation, /<p:notesSz cx="6858000" cy="9144000"\/>/);
  assert.match(presentation, /xmlns:a="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/main"/);
  assert.match(presentation, /xmlns:r="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/);
  assert.match(presentation, /xmlns:p="http:\/\/schemas\.openxmlformats\.org\/presentationml\/2006\/main"/);

  const layout = readPart(parts, 'ppt/slideLayouts/slideLayout1.xml');
  assert.match(layout, /type="blank"/);
  assert.match(layout, /<p:clrMapOvr><a:masterClrMapping\/><\/p:clrMapOvr>/);
  assert.doesNotMatch(layout, /<p:ph\b/, 'blank layout must not declare placeholders');
  const master = readPart(parts, 'ppt/slideMasters/slideMaster1.xml');
  const layoutId = /<p:sldLayoutId id="(\d+)" r:id="([^"]+)"\/>/.exec(master);
  assert.ok(layoutId, 'master has no <p:sldLayoutId>');
  assert.ok(Number(layoutId[1]) >= 2147483649);
  assert.match(master, /<p:clrMap bg1="lt1" tx1="dk1"/);
  assert.match(readPart(parts, 'ppt/theme/theme1.xml'), /<a:clrScheme name="/);

  // Slide XML: background, group id, shapes.
  const slide = readPart(parts, 'ppt/slides/slide1.xml');
  assert.match(slide, /<p:cSld name="page-1">/);
  assert.match(slide, /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="f0f0f0"\/><\/a:solidFill>/);
  assert.match(slide, /<p:cNvPr id="1" name=""\/>/);
  const shapes = parseShapes(slide);
  assert.equal(shapes.length, 4);
  assert.deepEqual(
    shapes.map((s) => [s.id, s.name]),
    [
      [2, '矩形 2'],
      [3, '矩形 3'],
      [4, '图片 4'],
      [5, '文本 5'],
    ],
  );
  assert.deepEqual(
    shapes.map((s) => [s.x, s.y, s.w, s.h]),
    [
      [0, 0, 300 * EMU, 20 * EMU],
      [10 * EMU, 20 * EMU, 40 * EMU, 30 * EMU],
      [100 * EMU, 50 * EMU, 8 * EMU, 8 * EMU],
      [20 * EMU, 120 * EMU, 120 * EMU, 20 * EMU],
    ],
    'every object must sit at Math.round(px * 9525) EMU',
  );
  // Rect: preset rect, solid fill, explicitly no outline; colours lowercased.
  for (const rect of shapes.slice(0, 2)) {
    assert.match(rect.body, /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom>/);
    assert.match(rect.body, /<a:ln><a:noFill\/><\/a:ln>/);
  }
  assert.match(shapes[0].body, /<a:srgbClr val="1f6feb"\/>/);
  assert.match(shapes[1].body, /<a:srgbClr val="abcdef"\/>/, 'colour must be lowercase hex');
  // Picture: blipFill + stretch fillRect, no solid fill of its own.
  assert.match(
    shapes[2].body,
    /<p:blipFill><a:blip r:embed="(rId\d+)"\/><a:stretch><a:fillRect\/><\/a:stretch><\/p:blipFill>/,
  );
  assert.doesNotMatch(shapes[2].body, /<a:solidFill>/);
  // Text box: txBox, zero insets, no wrap, top anchor, size/bold/align mapped.
  assert.match(shapes[3].body, /<p:cNvSpPr txBox="1"\/>/);
  assert.match(
    shapes[3].body,
    /<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit\/><\/a:bodyPr>/,
  );
  assert.match(shapes[3].body, /<a:pPr algn="l">/);
  assert.match(shapes[3].body, /<a:rPr lang="zh-CN" sz="1500" b="1" dirty="0">/);
  assert.match(shapes[3].body, /<a:t>Hello &lt;PixelTrace&gt; &amp; &quot;friends&quot;<\/a:t>/);
  // Text ink must contrast with the erase fill (#ffffff -> dark ink).
  assert.match(shapes[3].body, /<a:srgbClr val="000000"\/>/);
  assert.match(shapes[3].body, /<a:latin typeface="Arial"\/>/);

  // Picture relationship resolves to a media part that decodes to the crop.
  const slideRels = readPart(parts, 'ppt/slides/_rels/slide1.xml.rels');
  const embed = /<a:blip r:embed="(rId\d+)"\/>/.exec(shapes[2].body)[1];
  const rel = new RegExp(`<Relationship Id="${embed}" Type="[^"]*/image" Target="([^"]+)"\\/>`).exec(slideRels);
  assert.ok(rel, `no image relationship ${embed}`);
  const mediaName = posixPath.normalize(posixPath.join('ppt/slides', rel[1]));
  const media = await pptxDecodePng(parts.get(mediaName));
  assert.deepEqual([media.width, media.height], [8, 8]);
  assert.deepEqual([...media.data], [...image.data], 'media must be the exact crop');

  assertPackageIntegrity(parts);
  assert.equal(parts.size, 11 + 2 * 1 + 1, 'unexpected number of package parts');
});

test('pptxBuild de-duplicates media and keeps shape ids unique across slides', async () => {
  const shared = cropOf(4, 4, checker);
  const other = cropOf(4, 4, gradient);
  const slideA = sceneOf(120, 80, '#ffffff', [
    { k: 'i', x: 0, y: 0, w: 4, h: 4, px: shared, role: 'photo' },
    { k: 'i', x: 10, y: 10, w: 4, h: 4, px: shared, role: 'photo' },
    { k: 'i', x: 20, y: 20, w: 4, h: 4, px: other, role: 'detail' },
  ]);
  const slideB = sceneOf(120, 80, '#000000', [
    { k: 'i', x: 0, y: 0, w: 4, h: 4, px: shared, role: 'photo' },
    { k: 'r', x: 40, y: 40, w: 20, h: 10, c: '#00ff00', role: 'flat' },
  ]);
  const bytes = await pptxBuild([
    { scene: slideA, name: 'a' },
    { scene: slideB, name: 'b' },
  ]);
  const parts = await pptxUnzip(bytes);
  const mediaNames = [...parts.keys()].filter((n) => n.startsWith('ppt/media/'));
  assert.deepEqual(mediaNames, ['ppt/media/image1.png', 'ppt/media/image2.png']);
  assert.equal(parts.size, 11 + 2 * 2 + 2);

  const ids = [];
  for (const index of [1, 2]) {
    const slide = readPart(parts, `ppt/slides/slide${index}.xml`);
    const shapes = parseShapes(slide);
    ids.push(...shapes.map((s) => s.id));
    const rels = readPart(parts, `ppt/slides/_rels/slide${index}.xml.rels`);
    const imageTargets = [...rels.matchAll(/Target="\.\.\/media\/([^"]+)"/g)].map((m) => m[1]);
    assert.equal(new Set(imageTargets).size, imageTargets.length, 'one relationship per image part');
    if (index === 1) {
      assert.deepEqual(imageTargets, ['image1.png', 'image2.png'], 'two identical crops share one rel');
      const embeds = [...slide.matchAll(/r:embed="(rId\d+)"/g)].map((m) => m[1]);
      assert.deepEqual(embeds, [embeds[0], embeds[0], embeds[2]], 'identical crops reuse a relationship');
      assert.notEqual(embeds[0], embeds[2], 'distinct crops use distinct relationships');
    } else {
      assert.deepEqual(imageTargets, ['image1.png'], 'media de-duplicated across slides');
    }
  }
  assert.equal(new Set(ids).size, ids.length, 'shape ids must be unique across the deck');
  assert.deepEqual([...ids].sort((a, b) => a - b), ids, 'shape ids must increase in document order');
  assert.ok(ids.every((id) => id >= 2), 'shape ids start at 2 (id 1 is the group)');
  assert.equal(ids.length, 3 + 2);
  assertPackageIntegrity(parts);

  const presentation = readPart(parts, 'ppt/presentation.xml');
  assertDeckIds(presentation, 2);
  assert.match(presentation, /<p:sldSz cx="1143000" cy="762000" type="custom"\/>/);
});

test('pptxBuild caps the slide size at maxSlideEmu and honours emuPerPx', async () => {
  const big = sceneOf(6000, 4000, '#ffffff', [
    { k: 'r', x: 0, y: 0, w: 6000, h: 4000, c: '#123456', role: 'bg' },
  ]);
  const bytes = await pptxBuild([{ scene: big }], { maxSlideEmu: 51206400 });
  const parts = await pptxUnzip(bytes);
  const presentation = readPart(parts, 'ppt/presentation.xml');
  const sldSz = /<p:sldSz cx="(\d+)" cy="(\d+)" type="custom"\/>/.exec(presentation);
  const scale = 51206400 / 6000 / EMU;
  const expectedCx = Math.round(6000 * EMU * scale);
  const expectedCy = Math.round(4000 * EMU * scale);
  assert.equal(Number(sldSz[1]), expectedCx);
  assert.equal(Number(sldSz[2]), expectedCy);
  assert.ok(Number(sldSz[1]) <= 51206400, 'slide must not exceed 56 in');
  assert.equal(Number(sldSz[1]), 51206400, '6000 px at 96 dpi is exactly the 56 in cap');
  const shapes = parseShapes(readPart(parts, 'ppt/slides/slide1.xml'));
  assert.deepEqual(
    [shapes[0].x, shapes[0].y, shapes[0].w, shapes[0].h],
    [0, 0, Math.round(6000 * EMU * scale), Math.round(4000 * EMU * scale)],
  );

  // Custom emuPerPx scales every object and the slide by the same factor.
  const small = sceneOf(100, 50, '#ffffff', [
    { k: 'r', x: 10, y: 5, w: 20, h: 10, c: '#ff0000', role: 'flat' },
  ]);
  const custom = await pptxBuild([{ scene: small }], { emuPerPx: 4762 });
  const customParts = await pptxUnzip(custom);
  const customSz = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(
    readPart(customParts, 'ppt/presentation.xml'),
  );
  assert.deepEqual([Number(customSz[1]), Number(customSz[2])], [100 * 4762, 50 * 4762]);
  const customShape = parseShapes(readPart(customParts, 'ppt/slides/slide1.xml'))[0];
  assert.deepEqual(
    [customShape.x, customShape.y, customShape.w, customShape.h],
    [10 * 4762, 5 * 4762, 20 * 4762, 10 * 4762],
  );
});

test('pptxBuild layout16x9 pads the canvas without moving the content', async () => {
  const scene = sceneOf(300, 200, '#ffffff', [
    { k: 'r', x: 0, y: 0, w: 10, h: 10, c: '#000000', role: 'flat' },
  ]);
  const bytes = await pptxBuild([{ scene }], { layout16x9: true });
  const parts = await pptxUnzip(bytes);
  const sldSz = /<p:sldSz cx="(\d+)" cy="(\d+)" type="custom"\/>/.exec(
    readPart(parts, 'ppt/presentation.xml'),
  );
  const cx = Number(sldSz[1]);
  const cy = Number(sldSz[2]);
  assert.ok(Math.abs(cx - (cy * 16) / 9) <= 1, `canvas must be 16:9, got ${cx}x${cy}`);
  assert.ok(cx > 300 * EMU, 'canvas must be widened, not cropped');
  assert.ok(cy === 200 * EMU, 'the short side stays the scene height');
  const shape = parseShapes(readPart(parts, 'ppt/slides/slide1.xml'))[0];
  assert.deepEqual([shape.x, shape.y], [0, 0], 'content stays anchored at the origin');
});

test('pptxBuild is deterministic, does not mutate scenes and skips empty shapes', async () => {
  const image = cropOf(6, 6, gradient);
  const scene = sceneOf(64, 48, '#eeeeee', [
    { k: 'r', x: 0, y: 0, w: 64, h: 8, c: '#336699', role: 'bg' },
    { k: 'i', x: 8, y: 8, w: 6, h: 6, px: image, role: 'photo' },
    { k: 't', x: 8, y: 20, w: 40, h: 12, s: 'ok', sz: 9, c: '#eeeeee', al: 'ctr', b: false, role: 'text' },
    { k: 'r', x: 0, y: 0, w: 0, h: 10, c: '#ff0000', role: 'flat' }, // degenerate
    { k: 'z', x: 1, y: 1, w: 5, h: 5 }, // unknown kind
  ]);
  const slides = [{ scene, name: 'x' }, { scene, name: 'y' }];

  const before = JSON.stringify({
    width: scene.width,
    height: scene.height,
    background: scene.background,
    shapes: scene.shapes.map((shape) => ({
      ...shape,
      px: shape.px ? { width: shape.px.width, height: shape.px.height, data: [...shape.px.data] } : undefined,
    })),
  });

  const first = await pptxBuild(slides, { title: 'T' });
  const second = await pptxBuild(slides, { title: 'T' });
  assert.deepEqual([...second], [...first], 'same input must produce identical bytes');
  const other = await pptxBuild(slides, { title: 'Different title' });
  assert.notDeepEqual([...other], [...first], 'core properties must still be honoured');

  const after = JSON.stringify({
    width: scene.width,
    height: scene.height,
    background: scene.background,
    shapes: scene.shapes.map((shape) => ({
      ...shape,
      px: shape.px ? { width: shape.px.width, height: shape.px.height, data: [...shape.px.data] } : undefined,
    })),
  });
  assert.equal(after, before, 'pptxBuild must not mutate the scene');

  const parts = await pptxUnzip(first);
  const shapes = parseShapes(readPart(parts, 'ppt/slides/slide1.xml'));
  assert.equal(shapes.length, 3, 'degenerate and unknown shapes are skipped');
  assert.deepEqual(
    shapes.map((s) => s.id),
    [2, 3, 4],
    'skipping a shape must not leave gaps or reuse ids',
  );
  assert.match(readPart(parts, 'docProps/core.xml'), /<dc:title>T<\/dc:title>/);

  // Empty slide list still yields an openable one-slide deck.
  const empty = await pptxBuild([]);
  const emptyParts = await pptxUnzip(empty);
  const emptySz = /<p:sldSz cx="(\d+)" cy="(\d+)" type="custom"\/>/.exec(
    readPart(emptyParts, 'ppt/presentation.xml'),
  );
  assert.deepEqual([Number(emptySz[1]), Number(emptySz[2])], [1280 * EMU, 720 * EMU]);
  assert.match(readPart(emptyParts, 'ppt/slides/slide1.xml'), /<p:bg>/);

  // Invalid input fails loudly.
  await assert.rejects(() => pptxBuild([{ scene: { width: 0, height: 10 } }]), /positive integer/);
  await assert.rejects(() => pptxBuild([{ name: 'no scene' }]), /no scene/);
  await assert.rejects(() => pptxBuild('nope'), /must be an array/);
  await assert.rejects(
    () =>
      pptxBuild([
        {
          scene: sceneOf(10, 10, '#ffffff', [
            { k: 'i', x: 0, y: 0, w: 2, h: 2, px: { data: new Uint8ClampedArray(4), width: 2, height: 2 } },
          ]),
        },
      ]),
    /px\.data is too short/,
  );
});

/* --------------------------------------------------------- demo deck + API */

function demoSlides() {
  const slideA = sceneOf(300, 200, '#ffffff', [
    { k: 'r', x: 0, y: 0, w: 300, h: 40, c: '#1f6feb', role: 'bg' },
    { k: 'r', x: 0, y: 40, w: 300, h: 6, c: '#0b3d91', role: 'flat' },
    { k: 'r', x: 0, y: 180, w: 300, h: 20, c: '#222222', role: 'bg' },
    { k: 'i', x: 12, y: 60, w: 24, h: 24, px: cropOf(24, 24, checker), role: 'photo' },
    {
      k: 't',
      x: 52,
      y: 64,
      w: 200,
      h: 20,
      s: 'PixelTrace 1:1',
      sz: 15,
      c: '#ffffff',
      al: 'l',
      b: true,
      f: 'Arial',
      role: 'text',
    },
  ]);
  const slideB = sceneOf(300, 200, '#f7f7f7', [
    { k: 'r', x: 20, y: 120, w: 40, h: 60, c: '#e3622f', role: 'flat' },
    { k: 'r', x: 70, y: 80, w: 40, h: 100, c: '#2e9e5b', role: 'flat' },
    { k: 'r', x: 120, y: 100, w: 40, h: 80, c: '#8a5cf6', role: 'flat' },
    { k: 'i', x: 200, y: 40, w: 40, h: 40, px: cropOf(40, 40, gradient), role: 'detail' },
    {
      k: 't',
      x: 20,
      y: 24,
      w: 160,
      h: 24,
      s: 'Two slides,\nzero pixels lost.',
      sz: 17,
      c: '#f7f7f7',
      al: 'l',
      b: false,
      role: 'text',
    },
  ]);
  return [
    { scene: slideA, name: 'overview' },
    { scene: slideB, name: 'palette' },
  ];
}

let demoCache = null;

async function demoDeck() {
  if (demoCache !== null) return demoCache;
  const bytes = await pptxBuild(demoSlides(), {
    title: 'PixelTrace demo',
    subject: 'editable objects that rebuild a raster image exactly',
    creator: 'dsh-pixeltrace',
  });
  demoCache = { bytes, parts: await pptxUnzip(bytes) };
  return demoCache;
}

test('demo deck is built, inspectable and written to out/pptx-dev-demo.pptx', async (t) => {
  const { bytes, parts } = await demoDeck();
  assert.ok(bytes.length > 2000, 'demo deck looks too small');
  assertPackageIntegrity(parts);
  assert.equal(parts.size, 11 + 2 * 2 + 2, 'two slides, two distinct crops');

  const presentation = readPart(parts, 'ppt/presentation.xml');
  assertDeckIds(presentation, 2);
  assert.match(presentation, /<p:sldSz cx="2857500" cy="1905000" type="custom"\/>/);

  const slide1 = readPart(parts, 'ppt/slides/slide1.xml');
  const slide2 = readPart(parts, 'ppt/slides/slide2.xml');
  assert.match(slide1, /<p:cSld name="overview">/);
  assert.match(slide2, /<p:cSld name="palette">/);
  assert.equal(parseShapes(slide1).length, 5);
  assert.equal(parseShapes(slide2).length, 5);
  assert.match(slide2, /<a:t>Two slides,<\/a:t>/);
  assert.match(slide2, /<a:t>zero pixels lost.<\/a:t>/);

  mkdirSync(OUT_DIR, { recursive: true });
  writeFileSync(DEMO_PATH, bytes);
  t.diagnostic(`demo written: ${DEMO_PATH} (${bytes.length} bytes)`);
  t.diagnostic(`package parts: ${parts.size}`);
  for (const [name, data] of parts) t.diagnostic(`  ${name} ${u8(data).length}`);

  // Raw XML sanity: the first slide must look like real PresentationML.
  console.log('--- ppt/slides/slide1.xml (pretty printed) ---');
  console.log(slide1.replace(/></g, '>\n<'));
  console.log('--- end slide1.xml ---');
});

test('external OOXML validation with @office-kit/pptx (zero errors, exact EMU readback)', async (t) => {
  let officeKit = null;
  try {
    officeKit = await import(OFFICE_KIT_URL);
  } catch (error) {
    officeKit = null;
  }
  if (officeKit === null) {
    t.skip(`@office-kit/pptx is not available at ${OFFICE_KIT_URL}`);
    return;
  }
  const { loadPresentation, validatePresentation, getSlideCount, getSlides, getSlideShapes, getSlideSize, getShapeBounds, getShapeKind, getShapeName, getShapeId } = officeKit;
  assert.equal(typeof loadPresentation, 'function', 'office-kit API changed');

  const { bytes } = await demoDeck();
  const pres = await loadPresentation(bytes);

  const issues = validatePresentation(pres);
  const errors = issues.filter((issue) => issue.severity === 'error');
  t.diagnostic(`validation issues: ${JSON.stringify(issues)}`);
  assert.deepEqual(errors, [], 'validatePresentation reported errors');

  assert.equal(getSlideCount(pres), 2);
  const size = getSlideSize(pres);
  assert.equal(size.width, 2857500);
  assert.equal(size.height, 1905000);
  assert.equal(size.type, 'custom');

  const expected = [
    // slide 1: rects, picture, text box (EMU, scale 1)
    [
      { kind: 'shape', name: '矩形 2', bounds: [0, 0, 300 * EMU, 40 * EMU] },
      { kind: 'shape', name: '矩形 3', bounds: [0, 40 * EMU, 300 * EMU, 6 * EMU] },
      { kind: 'shape', name: '矩形 4', bounds: [0, 180 * EMU, 300 * EMU, 20 * EMU] },
      { kind: 'picture', name: '图片 5', bounds: [12 * EMU, 60 * EMU, 24 * EMU, 24 * EMU] },
      { kind: 'shape', name: '文本 6', bounds: [52 * EMU, 64 * EMU, 200 * EMU, 20 * EMU] },
    ],
    [
      { kind: 'shape', name: '矩形 7', bounds: [20 * EMU, 120 * EMU, 40 * EMU, 60 * EMU] },
      { kind: 'shape', name: '矩形 8', bounds: [70 * EMU, 80 * EMU, 40 * EMU, 100 * EMU] },
      { kind: 'shape', name: '矩形 9', bounds: [120 * EMU, 100 * EMU, 40 * EMU, 80 * EMU] },
      { kind: 'picture', name: '图片 10', bounds: [200 * EMU, 40 * EMU, 40 * EMU, 40 * EMU] },
      { kind: 'shape', name: '文本 11', bounds: [20 * EMU, 24 * EMU, 160 * EMU, 24 * EMU] },
    ],
  ];

  const slides = getSlides(pres);
  const allIds = [];
  for (let index = 0; index < slides.length; index++) {
    const shapes = getSlideShapes(slides[index]);
    assert.equal(shapes.length, expected[index].length, `slide ${index + 1} shape count`);
    t.diagnostic(`slide ${index + 1}: ${shapes.length} shapes`);
    for (let shapeIndex = 0; shapeIndex < shapes.length; shapeIndex++) {
      const shape = shapes[shapeIndex];
      const want = expected[index][shapeIndex];
      const bounds = getShapeBounds(shape);
      assert.ok(bounds, `shape ${shapeIndex} has no bounds`);
      assert.deepEqual(
        [bounds.x, bounds.y, bounds.w, bounds.h],
        want.bounds,
        `slide ${index + 1} shape ${shapeIndex} geometry`,
      );
      assert.equal(getShapeName(shape), want.name);
      const kind = getShapeKind(shape);
      assert.equal(kind, want.kind, `slide ${index + 1} shape ${shapeIndex} kind`);
      const id = getShapeId(shape);
      assert.ok(id >= 2);
      allIds.push(id);
      t.diagnostic(`  ${kind} ${id} "${want.name}" ${JSON.stringify(bounds)}`);
    }
  }
  assert.equal(new Set(allIds).size, allIds.length, 'shape ids must be unique across the deck');

  // The external reader must resolve every part's content type the way we declared it.
  // office-kit lists the 16 OPC parts but omits the `[Content_Types].xml` map itself.
  const packageParts = officeKit.listPackageParts(pres);
  assert.equal(packageParts.length, 16);
  assert.deepEqual(
    packageParts.map((part) => part.name).sort(),
    [
      '/_rels/.rels',
      '/docProps/app.xml',
      '/docProps/core.xml',
      '/ppt/_rels/presentation.xml.rels',
      '/ppt/media/image1.png',
      '/ppt/media/image2.png',
      '/ppt/presentation.xml',
      '/ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      '/ppt/slideLayouts/slideLayout1.xml',
      '/ppt/slideMasters/_rels/slideMaster1.xml.rels',
      '/ppt/slideMasters/slideMaster1.xml',
      '/ppt/slides/_rels/slide1.xml.rels',
      '/ppt/slides/_rels/slide2.xml.rels',
      '/ppt/slides/slide1.xml',
      '/ppt/slides/slide2.xml',
      '/ppt/theme/theme1.xml',
    ],
  );
  const byType = new Map(packageParts.map((part) => [part.name, part.contentType]));
  assert.equal(
    byType.get('/ppt/presentation.xml'),
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  );
  assert.equal(
    byType.get('/ppt/slides/slide1.xml'),
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  );
  assert.equal(byType.get('/ppt/slides/slide2.xml'), 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
  assert.equal(byType.get('/ppt/media/image1.png'), 'image/png');
  assert.equal(byType.get('/ppt/theme/theme1.xml'), 'application/vnd.openxmlformats-officedocument.theme+xml');
  for (const part of packageParts) assert.ok(part.byteLength > 0, `empty part ${part.name}`);
  t.diagnostic(`office-kit package parts: ${packageParts.length}`);

  // Media parts must be the PNGs the fragment wrote, and read back byte-identical.
  const media = [...(await pptxUnzip(bytes)).entries()].filter(([name]) => name.startsWith('ppt/media/'));
  assert.equal(media.length, 2);
  for (const [, data] of media) {
    const decoded = await pptxDecodePng(data);
    assert.ok(decoded.width > 0 && decoded.height > 0);
  }
});
