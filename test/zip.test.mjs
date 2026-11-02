/**
 * Strict structural checks for the ZIP container that pptx.frag.js writes.
 *
 * office-kit, bsdtar and PowerShell all happened to accept a ZIP whose EOCD
 * "central directory size" field was wrong, but the LibreOffice render bridge
 * (deck-render.mjs in dsh-ppt) validates `cdStart + cdSize <= eocdOffset` and
 * rejected the deck with "PPTX ZIP 目录无效或超过限制". This test parses the
 * archive the same strict way so the bug cannot come back.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadFrags, loadPptx } from './harness.mjs';

const pptx = loadFrags(['pptx.frag.js'], ['pptxZip', 'pptxUnzip', 'pptxEncodePng']);

function parseContainer(bytes) {
	let end = -1;
	for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset--) {
		if (
			bytes[offset] === 0x50 && bytes[offset + 1] === 0x4b && bytes[offset + 2] === 0x05 && bytes[offset + 3] === 0x06 &&
			offset + 22 + bytes[offset + 20] + (bytes[offset + 21] << 8) === bytes.length
		) {
			end = offset;
			break;
		}
	}
	assert.ok(end >= 0, 'end of central directory record must sit exactly at the end of the file');
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const count = view.getUint16(end + 10, true);
	const cdSize = view.getUint32(end + 12, true);
	const cdStart = view.getUint32(end + 16, true);
	assert.equal(view.getUint16(end + 4, true), 0, 'disk number must be 0');
	assert.equal(view.getUint16(end + 6, true), 0, 'central directory disk must be 0');
	assert.ok(count < 65535, 'no ZIP64 marker');
	assert.ok(count <= 20000, 'entry count must stay under the render bridge limit');
	assert.ok(cdStart + cdSize <= end, `central directory must end before the EOCD record (${cdStart}+${cdSize}<=${end})`);
	assert.equal(cdStart + cdSize, end, 'central directory must end exactly where the EOCD record starts');
	assert.equal(cdSize, end - cdStart);

	// every central record must point at its own local header, inside the data area
	const entries = [];
	let offset = cdStart;
	for (let index = 0; index < count; index++) {
		assert.equal(view.getUint32(offset, true), 0x02014b50, `central record ${index} signature`);
		const flags = view.getUint16(offset + 8, true);
		const method = view.getUint16(offset + 10, true);
		const compressed = view.getUint32(offset + 20, true);
		const expanded = view.getUint32(offset + 24, true);
		const nameLength = view.getUint16(offset + 28, true);
		const extra = view.getUint16(offset + 30, true);
		const comment = view.getUint16(offset + 32, true);
		const local = view.getUint32(offset + 42, true);
		const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
		assert.equal(extra, 0, `${name}: no extra field`);
		assert.equal(comment, 0, `${name}: no comment`);
		assert.ok(compressed < 0xffffffff && expanded < 0xffffffff && local < 0xffffffff, `${name}: no ZIP64 field`);
		assert.ok(local < cdStart, `${name}: local header inside the data area`);
		assert.equal(view.getUint32(local, true), 0x04034b50, `${name}: local header signature`);
		assert.equal(view.getUint16(local + 8, true), method, `${name}: compression method matches the local header`);
		assert.equal(view.getUint16(local + 6, true), flags, `${name}: flags match the local header`);
		assert.equal(view.getUint16(local + 26, true), nameLength, `${name}: name lengths match`);
		const localEnd = local + 30 + nameLength + compressed;
		assert.ok(localEnd <= cdStart, `${name}: payload must end before the central directory`);
		entries.push({ name, local, localEnd, method, compressed, expanded });
		offset += 46 + nameLength + extra + comment;
	}
	assert.equal(offset, end, 'central directory records must fill the declared size exactly');
	// local payloads must tile the data area without gaps or overlaps
	entries.sort((a, b) => a.local - b.local);
	let cursor = 0;
	for (const entry of entries) {
		assert.equal(entry.local, cursor, `${entry.name}: local headers must be contiguous`);
		cursor = entry.localEnd;
	}
	assert.equal(cursor, cdStart, 'local headers must fill the data area exactly');
	return entries;
}

test('pptxZip writes a ZIP the strict render bridge accepts', async () => {
	const png = await pptx.pptxEncodePng(new Uint8ClampedArray([255, 0, 0, 255]), 1, 1);
	const bytes = await pptx.pptxZip([
		{ name: 'mimetype', data: new TextEncoder().encode('application/vnd.oasis.opendocument.presentation') },
		{ name: 'ppt/media/image1.png', data: png },
		{ name: '[Content_Types].xml', data: new TextEncoder().encode('<?xml version="1.0"?><Types/>') },
	]);
	const entries = parseContainer(bytes);
	assert.deepEqual(entries.map((entry) => entry.name).sort(), ['[Content_Types].xml', 'mimetype', 'ppt/media/image1.png']);
	// inflating every entry with the archive's own reader is still the round-trip check
	const map = await pptx.pptxUnzip(bytes);
	assert.equal(map.size, 3);
});

test('pptxZip stays structurally valid for a large deck', async () => {
	// 700 entries push the central directory well past a single 32 KiB chunk
	const files = [];
	for (let index = 0; index < 700; index++) {
		files.push({ name: `ppt/media/pad-${index}.txt`, data: new TextEncoder().encode(`payload ${index} `.repeat(index % 17)) });
	}
	const bytes = await pptx.pptxZip(files);
	const entries = parseContainer(bytes);
	assert.equal(entries.length, 700);
	const map = await pptx.pptxUnzip(bytes);
	assert.equal(map.size, 700);
});
