#!/usr/bin/env node
/**
 * Concatenate the plain script fragments in src/ into lib/client.js.
 *
 * There is no bundler and no dependency: the fragments are designed to be pasted
 * into one closure, which is exactly what the browser module loader expects.
 * Run: node tools/build.mjs
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const PKG = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'));

/** Load order matters: helpers first, UI last. */
const FRAGMENTS = ['entry.frag.js', 'scene.frag.js', 'pptx.frag.js', 'ui.frag.js'];

/** `await` cannot be detected by a line regex (async functions legitimately use it);
 * a stray top-level await inside the factory is a SyntaxError that the parse check
 * below reports anyway. */
const FORBIDDEN = [
  { re: /^[ \t]*import\s[\s\S]*?from\s+['"]/m, what: 'import statement' },
  { re: /^[ \t]*export\s/m, what: 'export statement' },
];

function readFragment(name) {
  const file = path.join(root, 'src', name);
  let code;
  try {
    code = readFileSync(file, 'utf8');
  } catch (error) {
    throw new Error(`missing fragment ${name}: ${error.message}`);
  }
  // Strip a UTF-8 BOM the editor may add.
  code = code.replace(/^\uFEFF/, '');
  for (const rule of FORBIDDEN) {
    if (rule.re.test(code)) {
      throw new Error(`src/${name} contains a forbidden ${rule.what}; fragments must be plain script bodies`);
    }
  }
  return code.trimEnd();
}

const parts = FRAGMENTS.map((name) => {
  const code = readFragment(name);
  return `/* ==================== src/${name} ==================== */\n${code}\n`;
});

const bundle = `/* dsh-pixeltrace — ${PKG.version}
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

${parts.join('\n')}
    function apply(ctx) {
      return uiApply(ctx);
    }

    return { inject: ['slots', 'locale'], apply };
  },
});
`;

// Parse-only check: never throws on valid syntax even though `window` does not exist here.
try {
  // eslint-disable-next-line no-new-func
  new Function(bundle);
} catch (error) {
  console.error('lib/client.js would not parse:', error.message);
  process.exit(1);
}

mkdirSync(path.join(root, 'lib'), { recursive: true });
const target = path.join(root, 'lib', 'client.js');
writeFileSync(target, bundle, 'utf8');

const lines = bundle.split('\n').length;
const bytes = Buffer.byteLength(bundle, 'utf8');
console.log(`wrote lib/client.js  (${lines} lines, ${(bytes / 1024).toFixed(1)} KiB)`);
for (const [index, name] of FRAGMENTS.entries()) {
  const size = Buffer.byteLength(parts[index], 'utf8');
  console.log(`  ${String(index + 1).padStart(2)}. src/${name.padEnd(18)} ${(size / 1024).toFixed(1)} KiB`);
}
