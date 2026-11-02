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
