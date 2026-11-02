// Static + pure-function tests for the Lead-owned fragments (entry.frag.js, ui.frag.js)
// and for the generated bundle. There is no React in this environment, so the panel
// component tree is checked structurally: the fragment must load, expose its pieces,
// and every string/CSS/theme reference must resolve -- the browser check
// (tools/browser-check.mjs) covers actual rendering.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { frag, repoRoot } from './harness.mjs';

const FRAGMENTS = ['entry.frag.js', 'scene.frag.js', 'pptx.frag.js', 'ui.frag.js'];

const WANTED = [
  'coreDefaults',
  'coreModes',
  'coreMaxColorChoices',
  'coreNormalizeOptions',
  'coreScene',
  'coreSummary',
  'coreSafeName',
  'coreDeckFileName',
  'coreBuildDeck',
  'coreSaveBytes',
  'coreIsHexColor',
  'uiPanelId',
  'uiSidebarOrder',
  'uiStyleId',
  'uiStrings',
  'uiCss',
  'uiInitialState',
  'uiPanel',
  'uiSidebarIcon',
  'uiApply',
  'uiAnalyze',
  'uiDiffIgnoring',
  'uiImageHasAlpha',
];

/** Minimal React stand-in: the fragments only need hook identities at load time. */
function reactStub() {
  const noop = () => {};
  return {
    createElement: (type, props, ...children) => ({ type, props, children }),
    useState: (initial) => [typeof initial === 'function' ? initial() : initial, noop],
    useEffect: noop,
    useLayoutEffect: noop,
    useRef: (initial) => ({ current: initial }),
    useMemo: (factory) => factory(),
    useCallback: (fn) => fn,
  };
}

function loadAll() {
  const code = FRAGMENTS.map((name) => frag(name)).join('\n;\n');
  // eslint-disable-next-line no-new-func
  const factory = new Function(
    'React',
    'require',
    `"use strict";\nconst { useState, useEffect, useLayoutEffect, useRef, useMemo, useCallback } = React;\nconst h = React.createElement;\n${code}\nreturn {${WANTED.join(', ')}};`,
  );
  return factory(reactStub(), () => ({}));
}

const api = loadAll();

test('bundle: generated lib/client.js matches the fragment contract', () => {
  const bundle = readFileSync(path.join(repoRoot, 'lib', 'client.js'), 'utf8');
  assert.equal(bundle.match(/__ModuleLoader__\.load\(/g).length, 1, 'exactly one loader call');
  assert.match(bundle, /id: 'dsh-pixeltrace'/);
  assert.match(bundle, /inject: \['slots', 'locale'\]/);
  assert.match(bundle, /const \{ useState, useEffect,/, 'hooks must be destructured from React');
  for (const name of FRAGMENTS) {
    assert.ok(bundle.includes(`src/${name}`), `bundle carries the ${name} banner`);
  }
  // The bundle is what the browser loads: it has to parse as a plain script.
  assert.doesNotThrow(() => new Function(bundle));
});

test('options: defaults are applied to an empty request', () => {
  assert.deepEqual(api.coreNormalizeOptions(undefined), {
    mode: 'smart',
    maxColors: 0,
    mergeTolerance: 0,
    shapeBudget: 8000,
    detailToImage: true,
    textRegions: false,
    alphaBackground: '#ffffff',
    maxSlidePx: 5376,
  });
});

test('options: hostile values are clamped, not trusted', () => {
  const opts = api.coreNormalizeOptions({
    mode: 'nope',
    maxColors: 7,
    mergeTolerance: 9999,
    shapeBudget: -40,
    alphaBackground: 'red',
    maxSlidePx: 99999,
    detailToImage: 'yes',
    textRegions: 'yes',
  });
  assert.equal(opts.mode, 'smart');
  assert.equal(opts.maxColors, 0, 'off-menu palette sizes fall back to 0');
  assert.equal(opts.mergeTolerance, 255);
  assert.equal(opts.shapeBudget, 0);
  assert.equal(opts.alphaBackground, '#ffffff');
  assert.equal(opts.maxSlidePx, 8192);
  assert.equal(opts.detailToImage, true);
  assert.equal(opts.textRegions, false, 'only a real boolean true enables text detection');
});

test('options: legal palette choices survive, colours are normalised', () => {
  const opts = api.coreNormalizeOptions({ maxColors: 32, alphaBackground: '#ABCDEF', mode: 'vector' });
  assert.equal(opts.maxColors, 32);
  assert.equal(opts.alphaBackground, '#abcdef');
  assert.equal(opts.mode, 'vector');
});

test('names: unsafe file names are cleaned and stay short', () => {
  assert.equal(api.coreSafeName('a/b:c*l?.png'), 'a-b-c-l-');
  assert.equal(api.coreSafeName('   ', 'fallback'), 'fallback');
  assert.equal(api.coreSafeName('', ''), 'pixeltrace');
  assert.ok(api.coreSafeName('x'.repeat(400)).length <= 80);
  assert.equal(api.coreDeckFileName('shot.png', 1), 'shot.pptx');
  assert.equal(api.coreDeckFileName('shot.png', 3), 'shot-3p.pptx');
});

test('summary: aggregates worst-case fidelity across slides', () => {
  const summary = api.coreSummary([
    { scene: { stats: { shapes: 5, rects: 4, images: 1, texts: 0, paletteSize: 6, exactRatio: 1, meanAbs: 0, maxAbs: 0, ms: 10, notes: ['a'] } } },
    { scene: { stats: { shapes: 3, rects: 3, images: 0, texts: 0, paletteSize: 9, exactRatio: 0.5, meanAbs: 3, maxAbs: 40, ms: 5, notes: ['a', 'b'] } } },
    { skipped: true },
  ]);
  assert.equal(summary.slides, 2);
  assert.equal(summary.shapes, 8);
  assert.equal(summary.images, 1);
  assert.equal(summary.palette, 9);
  assert.equal(summary.exactRatio, 0.5);
  assert.equal(summary.maxAbs, 40);
  assert.equal(summary.ms, 15);
  assert.deepEqual(summary.notes, ['a', 'b']);
});

test('deck: one processed image becomes a valid single-slide pptx', async () => {
  const width = 6;
  const height = 4;
  const data = new Uint8ClampedArray(width * height * 4);
  for (let i = 0; i < width * height; i++) {
    data[i * 4] = 200;
    data[i * 4 + 1] = 30;
    data[i * 4 + 2] = 90;
    data[i * 4 + 3] = 255;
  }
  const { scene } = api.coreScene({ width, height, data }, { mode: 'smart' });
  const { bytes, slides } = await api.coreBuildDeck([{ name: 'tiny', scene }]);
  assert.equal(slides, 1);
  assert.ok(bytes.byteLength > 1000);
  assert.deepEqual([bytes[0], bytes[1]], [0x50, 0x4b], 'ZIP magic');
});

test('save: without a DOM the bytes are returned untouched', async () => {
  const bytes = new Uint8Array([1, 2, 3]);
  const result = await api.coreSaveBytes(bytes, 'x.pptx', 'application/octet-stream');
  assert.equal(result.saved, false);
  assert.equal(result.fileName, 'x.pptx');
  assert.equal(result.bytes, bytes);
});

test('ui: translations cover every literal used by the panel', () => {
  const source = frag('ui.frag.js');
  const keys = new Set();
  for (const match of source.matchAll(/uiT\('([^'\\]+)'\)/g)) keys.add(match[1]);
  assert.ok(keys.size > 40, `expected a real string table, saw ${keys.size} keys`);
  for (const key of keys) {
    assert.ok(Object.prototype.hasOwnProperty.call(api.uiStrings, key), `missing uiStrings entry for ${JSON.stringify(key)}`);
    const english = api.uiStrings[key];
    assert.equal(typeof english, 'string');
    assert.ok(english.length > 0, `empty translation for ${JSON.stringify(key)}`);
  }
});

test('ui: CSS is scoped, balanced, and only uses documented theme tokens', () => {
  const css = api.uiCss;
  assert.ok(css.includes('.pt-root'));
  assert.equal((css.match(/\{/g) || []).length, (css.match(/\}/g) || []).length, 'balanced braces');
  assert.ok(!css.includes('@import'), 'no remote imports');
  assert.ok(!/url\(\s*['"]?https?:/i.test(css), 'no remote assets');
  const tokens = new Set([
    '--dsw-alias-bg-base',
    '--dsw-alias-bg-layer-1',
    '--dsw-alias-bg-layer-2',
    '--dsw-alias-bg-overlay',
    '--dsw-alias-border-l1',
    '--dsw-alias-border-l2',
    '--dsw-alias-brand-primary',
    '--dsw-alias-label-primary',
    '--dsw-alias-label-secondary',
    '--dsw-alias-state-error-primary',
    '--dsw-alias-state-idle-primary',
    '--dsw-alias-state-success-primary',
    '--dsw-alias-state-warn-primary',
    '--dsw-specific-sidebar-fill',
  ]);
  for (const match of css.matchAll(/var\((--dsw-[a-z0-9-]+)/g)) {
    assert.ok(tokens.has(match[1]), `undocumented theme token ${match[1]}`);
  }
});

test('ui: initial state and exported pieces line up with the contract', () => {
  const state = api.uiInitialState();
  assert.ok(Array.isArray(state.items));
  assert.equal(state.items.length, 0);
  assert.equal(state.mode, 'smart');
  assert.equal(state.shapeBudget, 8000);
  assert.equal(state.view, 'rebuilt');
  assert.equal(typeof api.uiPanel, 'function');
  assert.equal(typeof api.uiSidebarIcon, 'function');
  assert.equal(typeof api.uiApply, 'function');
  assert.equal(api.uiPanelId, 'pixeltrace');
  assert.equal(api.uiStyleId, 'dsh-pixeltrace-style');
  assert.ok(api.uiSidebarOrder > 0);
});

function solidImage(width, height, rgba) {
  const data = new Uint8ClampedArray(width * height * 4);
  for (let index = 0; index < width * height; index += 1) {
    data[index * 4] = rgba[0];
    data[index * 4 + 1] = rgba[1];
    data[index * 4 + 2] = rgba[2];
    data[index * 4 + 3] = rgba[3];
  }
  return { width, height, data };
}

test('ui: transparency detection follows the alpha channel', () => {
  assert.equal(api.uiImageHasAlpha(solidImage(4, 4, [1, 2, 3, 255])), false);
  const translucent = solidImage(4, 4, [1, 2, 3, 255]);
  translucent.data[20 + 3] = 128;
  assert.equal(api.uiImageHasAlpha(translucent), true);
  assert.equal(api.uiImageHasAlpha({ width: 0, height: 0, data: null }), false);
});

test('ui: diff honours the ignored rectangles', () => {
  const reference = solidImage(4, 4, [0, 0, 0, 255]);
  const candidate = solidImage(4, 4, [0, 0, 0, 255]);
  // one changed pixel at (3,3) and one at (0,0)
  candidate.data[15 * 4] = 200;
  candidate.data[0] = 40;
  const raw = api.uiDiffIgnoring(reference, candidate, []);
  assert.equal(raw.pixels, 16);
  assert.equal(raw.mismatchPixels, 2);
  assert.equal(raw.exactRatio, 14 / 16);
  assert.equal(raw.maxAbs, 200);
  const ignoring = api.uiDiffIgnoring(reference, candidate, [{ x: 0, y: 0, w: 1, h: 1 }, { x: 3, y: 3, w: 1, h: 1 }]);
  assert.equal(ignoring.pixels, 14);
  assert.equal(ignoring.mismatchPixels, 0);
  assert.equal(ignoring.exactRatio, 1);
  assert.equal(ignoring.maxAbs, 0);
});

test('ui: analysis reports the engine fidelity for a flat image', async () => {
  const image = solidImage(8, 8, [10, 20, 30, 255]);
  const items = await api.uiAnalyze([{ id: 'a', name: 'flat', image }], api.coreNormalizeOptions({ mode: 'smart' }));
  assert.equal(items.length, 1);
  const [item] = items;
  assert.equal(item.error, null);
  assert.equal(item.stats.exactRatio, 1);
  assert.equal(item.stats.mismatchPixels, 0);
  assert.ok(item.scene.shapes.length >= 1);
  assert.ok(item.raster && item.raster.data.length === 8 * 8 * 4);
  assert.equal(item.diff.exactRatio, 1);
});

