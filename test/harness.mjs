// Shared test harness: evaluates the plain script fragments exactly the way the
// browser bundle does (one scope, no imports/exports).
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));

/** Read one fragment body from src/. */
export function frag(name) {
  return readFileSync(path.join(root, 'src', name), 'utf8');
}

/**
 * Evaluate fragments in one scope and return the named top-level bindings.
 * @param {string[]} names fragment file names, in dependency order
 * @param {string[]} wanted top-level identifiers to expose
 */
export function loadFrags(names, wanted) {
  const code = names.map((n) => `//#region ${n}\n${frag(n)}\n//#endregion`).join('\n;\n');
  // eslint-disable-next-line no-new-func
  const factory = new Function(
    'React',
    'require',
    `"use strict";\n${code}\nreturn {${wanted.join(', ')}};`,
  );
  return factory(undefined, (id) => {
    if (id === 'react') return undefined;
    throw new Error(`fragment asked for unexpected module: ${id}`);
  });
}

/** scene engine only. */
export function loadScene() {
  return loadFrags(
    ['scene.frag.js'],
    ['sceneBuild', 'sceneRasterize', 'sceneDiff', 'sceneDefaults'],
  );
}

/** pptx writer only. */
export function loadPptx() {
  return loadFrags(
    ['pptx.frag.js'],
    ['pptxBuild', 'pptxEncodePng', 'pptxDecodePng', 'pptxZip', 'pptxUnzip'],
  );
}

/** Both engines together (scene functions first, as in the bundle). */
export function loadEngine() {
  return loadFrags(
    ['scene.frag.js', 'pptx.frag.js'],
    [
      'sceneBuild',
      'sceneRasterize',
      'sceneDiff',
      'sceneDefaults',
      'pptxBuild',
      'pptxEncodePng',
      'pptxDecodePng',
      'pptxZip',
      'pptxUnzip',
    ],
  );
}

export const repoRoot = root;
