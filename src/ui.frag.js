/* PixelTrace panel: drop an image, rebuild it with PowerPoint objects, export .pptx.
 *
 * Plain fragment: no import/export, no top-level await. React arrives from the
 * browser module table as `React`/`h` in the surrounding factory closure.
 * Every top-level name starts with `ui`.
 */

const uiPanelId = 'pixeltrace';
const uiSidebarOrder = 32;
const uiStyleId = 'dsh-pixeltrace-style';

/** zh -> en. Untranslated keys fall back to the Chinese key itself. */
const uiStrings = {
  '像素复刻 PixelTrace': 'PixelTrace',
  '把一张图还原成 PPT 里可编辑的形状 · 像素级一致':
    'Rebuild any image as editable PowerPoint objects, pixel for pixel',
  '拖入图片，或点击选择': 'Drop an image here, or click to choose',
  'PNG · JPEG · WebP · GIF · 可多选（每张一页）· 也可以直接 Ctrl+V 粘贴':
    'PNG · JPEG · WebP · GIF · multiple files (one slide each) · or Ctrl+V to paste',
  '选择图片': 'Choose images',
  '清空': 'Clear',
  '输出参数': 'Output',
  '输出模式': 'Mode',
  '智能分层': 'Smart layers',
  '纯矢量': 'Vector only',
  '整图保真': 'Whole picture',
  '平面区域变成可编辑色块，细节区域变成精确裁剪的图片，整页仍逐像素一致':
    'Flat areas become editable rectangles, detail areas become exactly cropped pictures; the slide stays pixel-identical',
  '全部变成可编辑矩形：每个色块都能单独改，但形状数量最多':
    'Everything becomes editable rectangles: every block is editable, but the shape count is highest',
  '每页只放一张图片对象：100% 一致，但不能分层编辑':
    'One picture object per slide: 100% identical, but nothing to edit layer by layer',
  '颜色量化': 'Colour quantisation',
  '0 = 保留原始颜色': '0 = keep original colours',
  '颜色容差': 'Colour tolerance',
  '合并相邻矩形时允许的单通道色差，0 = 完全精确': 'Per-channel delta allowed when merging rectangles; 0 = exact',
  '形状预算/页': 'Shape budget / slide',
  '0 = 不限制；超出预算时自动降级并在统计里说明': '0 = unlimited; over budget it degrades and says so in the notes',
  '细节区域转为精确裁剪图': 'Turn detail areas into exact crops',
  '文字识别为可编辑文本框': 'Detect text as editable text boxes',
  '实验性：文字墨迹被擦成背景色，留下一个空文本框，方便你自己输入原文':
    'Experimental: ink is erased to the background colour and an empty text box is left for you to type',
  '透明背景色': 'Transparent background',
  '重新分析': 'Re-analyse',
  '分析中…': 'Analysing…',
  '图片': 'Images',
  '还没有图片': 'No image yet',
  '移除': 'Remove',
  '预览': 'Preview',
  '原图': 'Source',
  '重建': 'Rebuilt',
  '差异': 'Diff',
  '适应': 'Fit',
  '保真统计': 'Fidelity',
  '对象': 'objects',
  '矩形': 'rects',
  '裁剪图': 'crops',
  '文本框': 'texts',
  '颜色': 'colours',
  '完全一致像素': 'Exact pixels',
  '不一致像素': 'Mismatched px',
  '平均误差': 'Mean error',
  '最大误差': 'Max error',
  '耗时': 'Time',
  '备注': 'Notes',
  '导出 PPTX': 'Export PPTX',
  '正在生成…': 'Building…',
  '文件名': 'File name',
  '参数已修改，点「重新分析」刷新预览': 'Settings changed — press Re-analyse to refresh the preview',
  '先添加图片才能导出': 'Add an image first',
  '已导出': 'Exported',
  '已取消': 'Cancelled',
  '出错': 'Error',
  '拖入图片后在这里查看逐像素预览': 'Drop an image, then inspect the pixel-exact preview here',
  '提示：导出的每个色块、裁剪图、文本框都是独立对象，可在 PowerPoint / WPS / Keynote 里单独编辑。':
    'Tip: every rectangle, crop and text box is a separate object you can edit in PowerPoint / WPS / Keynote.',
  '正在读取图片…': 'Reading image…',
  '正在分析…': 'Analysing…',
  '正在生成 PPTX…': 'Building the .pptx…',
  '图片超过 56 英寸幻灯片上限，已等比缩放': 'Image exceeds the 56-inch slide cap and was scaled down',
  '图片过大，已降采样以获得可用速度': 'Very large image downsampled for usable speed',
  '全部精确重建': 'Reconstructed exactly',
  '页': 'slides',
  '张': '',
  '无法读取该图片': 'Could not read this image',
  '该浏览器不支持画布读取像素': 'This browser cannot read canvas pixels',
};

let uiLocaleService = null;

function uiSetLocaleService(service) {
  uiLocaleService = service || null;
}

function uiLanguageOf(locale) {
  const source = locale || uiLocaleService;
  try {
    const snapshot = source && typeof source.getSnapshot === 'function' ? source.getSnapshot() : null;
    const active = snapshot && typeof snapshot === 'object' ? snapshot.active || snapshot.locale || '' : '';
    return String(active || '').toLowerCase().startsWith('en') ? 'en' : 'zh';
  } catch {
    return 'zh';
  }
}

function uiT(zh) {
  if (uiLanguageOf() !== 'en') {
    return zh;
  }
  return Object.prototype.hasOwnProperty.call(uiStrings, zh) ? uiStrings[zh] : zh;
}

function uiUseLanguage(locale) {
  const [, bump] = useState(0);
  useEffect(() => {
    if (!locale || typeof locale.subscribe !== 'function') {
      return undefined;
    }
    let alive = true;
    const update = () => {
      if (alive) {
        bump((n) => n + 1);
      }
    };
    const off = locale.subscribe(update);
    return () => {
      alive = false;
      if (typeof off === 'function') {
        off();
      }
    };
  }, [locale]);
}

const uiCss = `
.pt-root{display:flex;flex-direction:column;gap:14px;height:100%;min-height:0;box-sizing:border-box;padding:18px 22px 22px;overflow:auto;background:var(--dsw-alias-bg-base,transparent);color:var(--dsw-alias-label-primary,#1f2328);font:13px/1.55 "Segoe UI","PingFang SC","Microsoft YaHei",system-ui,sans-serif;}
.pt-head{display:flex;align-items:center;gap:12px;flex:0 0 auto;}
.pt-logo{width:36px;height:36px;border-radius:10px;display:grid;place-items:center;flex:0 0 auto;color:var(--dsw-alias-brand-primary,#3b82f6);background:linear-gradient(135deg,rgba(59,130,246,.22),rgba(168,85,247,.22));border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.12));}
.pt-h1{margin:0;font-size:15px;font-weight:650;letter-spacing:.2px;}
.pt-sub{margin:2px 0 0;font-size:12px;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-drop{flex:0 0 auto;display:flex;align-items:center;gap:14px;padding:14px 16px;border-radius:12px;border:1.5px dashed var(--dsw-alias-border-l2,rgba(0,0,0,.2));background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.06));cursor:pointer;transition:border-color .15s ease,background .15s ease;outline:none;}
.pt-drop:hover,.pt-drop.is-over{border-color:var(--dsw-alias-brand-primary,#3b82f6);background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 9%,transparent);}
.pt-drop-icon{width:34px;height:34px;flex:0 0 auto;display:grid;place-items:center;border-radius:9px;color:var(--dsw-alias-brand-primary,#3b82f6);background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 14%,transparent);}
.pt-drop-title{font-weight:600;font-size:13px;}
.pt-drop-hint{font-size:11.5px;color:var(--dsw-alias-label-secondary,#6b7280);margin-top:2px;}
.pt-drop-spacer{flex:1 1 auto;}
.pt-form{flex:1 1 auto;min-height:0;display:grid;grid-template-columns:290px minmax(0,1fr);gap:16px;align-items:start;overflow:auto;padding-right:2px;}
@media (max-width:880px){.pt-form{grid-template-columns:minmax(0,1fr);}}
.pt-side{display:flex;flex-direction:column;gap:12px;min-width:0;}
.pt-stage{display:flex;flex-direction:column;gap:12px;min-width:0;min-height:0;overflow:auto;}
.pt-card{border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.12));background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.05));border-radius:12px;padding:12px;}
.pt-card-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:10px;}
.pt-card-title{font-size:11.5px;font-weight:650;letter-spacing:.6px;text-transform:uppercase;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-field{display:flex;flex-direction:column;gap:5px;margin-top:10px;}
.pt-field:first-child{margin-top:0;}
.pt-label{font-size:12px;font-weight:550;display:flex;justify-content:space-between;align-items:center;gap:8px;}
.pt-value{font-variant-numeric:tabular-nums;font-size:11.5px;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-hint{font-size:11px;color:var(--dsw-alias-label-secondary,#6b7280);line-height:1.45;}
.pt-seg{display:flex;gap:4px;padding:3px;border-radius:9px;background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.1));}
.pt-seg button{flex:1 1 0;appearance:none;border:0;background:transparent;color:inherit;font:inherit;font-size:12px;padding:6px 4px;border-radius:6px;cursor:pointer;transition:background .12s ease,color .12s ease;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.pt-seg button.is-on{background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-brand-primary,#3b82f6);font-weight:600;box-shadow:0 1px 2px rgba(0,0,0,.12);}
.pt-seg button:hover:not(.is-on){background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 10%,transparent);}
.pt-stage-bar .pt-seg{flex:0 0 auto;}
.pt-stage-bar .pt-seg button{flex:0 0 auto;padding:6px 10px;}
.pt-input,.pt-select{width:100%;box-sizing:border-box;padding:6px 9px;border-radius:8px;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.16));background:var(--dsw-alias-bg-base,#fff);color:inherit;font:inherit;font-size:12.5px;outline:none;}
.pt-input:focus,.pt-select:focus{border-color:var(--dsw-alias-brand-primary,#3b82f6);box-shadow:0 0 0 3px color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 18%,transparent);}
.pt-row{display:flex;align-items:center;gap:8px;}
.pt-check{display:flex;align-items:flex-start;gap:8px;margin-top:10px;cursor:pointer;}
.pt-check input{margin:2px 0 0;accent-color:var(--dsw-alias-brand-primary,#3b82f6);}
.pt-check-text{display:flex;flex-direction:column;gap:2px;}
.pt-check-title{font-size:12px;font-weight:550;}
.pt-swatch{width:34px;height:26px;padding:0;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.16));border-radius:7px;background:transparent;cursor:pointer;}
.pt-btn{appearance:none;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.16));background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.08));color:inherit;font:inherit;font-size:12.5px;font-weight:550;padding:7px 13px;border-radius:9px;cursor:pointer;transition:transform .06s ease,background .12s ease,border-color .12s ease;white-space:nowrap;}
.pt-btn:hover:not(:disabled){background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 12%,transparent);border-color:var(--dsw-alias-brand-primary,#3b82f6);}
.pt-btn:active:not(:disabled){transform:translateY(1px);}
.pt-btn:disabled{opacity:.45;cursor:not-allowed;}
.pt-btn.pt-primary{background:var(--dsw-alias-brand-primary,#3b82f6);border-color:transparent;color:#fff;}
.pt-btn.pt-primary:hover:not(:disabled){background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 86%,#000);}
.pt-btn.pt-ghost{border-color:transparent;background:transparent;color:var(--dsw-alias-label-secondary,#6b7280);padding:5px 8px;}
.pt-btn.pt-ghost:hover:not(:disabled){color:inherit;background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.12));}
.pt-btn.pt-small{font-size:11.5px;padding:4px 9px;border-radius:7px;}
.pt-files{display:flex;flex-direction:column;gap:5px;max-height:200px;overflow:auto;}
.pt-file{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:8px;cursor:pointer;border:1px solid transparent;background:var(--dsw-alias-bg-base,transparent);}
.pt-file.is-on{border-color:var(--dsw-alias-brand-primary,#3b82f6);background:color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 9%,transparent);}
.pt-file-body{min-width:0;flex:1 1 auto;}
.pt-file-name{display:block;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pt-file-meta{display:block;font-size:10.5px;color:var(--dsw-alias-label-secondary,#6b7280);font-variant-numeric:tabular-nums;}
.pt-x{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary,#6b7280);cursor:pointer;font-size:15px;line-height:1;padding:3px 5px;border-radius:6px;}
.pt-x:hover{color:var(--dsw-alias-state-error-primary,#dc2626);background:color-mix(in srgb,var(--dsw-alias-state-error-primary,#dc2626) 12%,transparent);}
.pt-empty{padding:22px 12px;text-align:center;font-size:12px;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-stage-bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
.pt-stage-title{font-size:12px;font-weight:600;margin-right:auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pt-canvas-wrap{position:relative;border-radius:12px;border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.12));background-color:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.08));background-image:linear-gradient(45deg,rgba(127,127,127,.13) 25%,transparent 25%),linear-gradient(-45deg,rgba(127,127,127,.13) 25%,transparent 25%),linear-gradient(45deg,transparent 75%,rgba(127,127,127,.13) 75%),linear-gradient(-45deg,transparent 75%,rgba(127,127,127,.13) 75%);background-size:16px 16px;background-position:0 0,0 8px,8px -8px,-8px 0;overflow:auto;padding:10px;max-height:46vh;}
.pt-canvas{display:block;margin:0 auto;border-radius:4px;box-shadow:0 2px 12px rgba(0,0,0,.18);}
.pt-stats{display:flex;flex-wrap:wrap;gap:6px;}
.pt-stat{display:flex;flex-direction:column;gap:1px;padding:6px 9px;border-radius:10px;background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.07));border:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.1));min-width:74px;}
.pt-stat-k{font-size:10.5px;letter-spacing:.4px;text-transform:uppercase;color:var(--dsw-alias-label-secondary,#6b7280);}
.pt-stat-v{font-size:14px;font-weight:650;font-variant-numeric:tabular-nums;}
.pt-stat.is-good .pt-stat-v{color:var(--dsw-alias-state-success-primary,#16a34a);}
.pt-stat.is-bad .pt-stat-v{color:var(--dsw-alias-state-error-primary,#dc2626);}
.pt-notes{font-size:11.5px;color:var(--dsw-alias-label-secondary,#6b7280);display:flex;flex-direction:column;gap:2px;margin-top:2px;}
.pt-notes code{font-size:11px;background:var(--dsw-alias-bg-layer-2,rgba(127,127,127,.14));border-radius:4px;padding:1px 4px;}
.pt-foot{flex:0 0 auto;display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding-top:4px;border-top:1px solid var(--dsw-alias-border-l1,rgba(0,0,0,.1));}
.pt-foot-name{display:flex;align-items:center;gap:7px;}
.pt-foot-name .pt-input{width:190px;}
.pt-msg{font-size:12px;border-radius:8px;padding:7px 11px;display:flex;align-items:center;gap:8px;}
.pt-msg.ok{color:var(--dsw-alias-state-success-primary,#16a34a);background:color-mix(in srgb,var(--dsw-alias-state-success-primary,#16a34a) 12%,transparent);}
.pt-msg.error{color:var(--dsw-alias-state-error-primary,#dc2626);background:color-mix(in srgb,var(--dsw-alias-state-error-primary,#dc2626) 12%,transparent);}
.pt-msg.info{color:var(--dsw-alias-label-secondary,#6b7280);background:var(--dsw-alias-bg-layer-1,rgba(127,127,127,.1));}
.pt-warn{font-size:11.5px;color:var(--dsw-alias-state-warn-primary,#b45309);}
.pt-spin{width:13px;height:13px;border-radius:50%;border:2px solid color-mix(in srgb,var(--dsw-alias-brand-primary,#3b82f6) 30%,transparent);border-top-color:var(--dsw-alias-brand-primary,#3b82f6);animation:pt-rot .8s linear infinite;flex:0 0 auto;}
@keyframes pt-rot{to{transform:rotate(360deg);}}
.pt-sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;}
`;

function uiMountStyle() {
  if (typeof document === 'undefined') {
    return () => {};
  }
  const existing = document.getElementById(uiStyleId);
  if (existing) {
    return () => {
      /* owned by an earlier mount */
    };
  }
  const style = document.createElement('style');
  style.id = uiStyleId;
  style.textContent = uiCss;
  document.head.appendChild(style);
  return () => {
    try {
      style.remove();
    } catch {
      /* ignore */
    }
  };
}

/* ---------------------------------------------------------------- image input */

const uiMaxDecodeSide = 6000;

function uiDecodeViaImageTag(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error(uiT('无法读取该图片')));
    };
    img.src = url;
  });
}

/** Decode any image-ish source into { width, height, data, downscaled }. */
async function uiRgbaFromSource(source) {
  const width = source.width || source.naturalWidth || 0;
  const height = source.height || source.naturalHeight || 0;
  if (!width || !height) {
    throw new Error(uiT('无法读取该图片'));
  }
  const ratio = Math.min(1, uiMaxDecodeSide / Math.max(width, height));
  const w = Math.max(1, Math.round(width * ratio));
  const h = Math.max(1, Math.round(height * ratio));
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const context2d = canvas.getContext('2d', { willReadFrequently: true });
  if (!context2d) {
    throw new Error(uiT('该浏览器不支持画布读取像素'));
  }
  context2d.imageSmoothingEnabled = ratio < 1;
  context2d.imageSmoothingQuality = 'high';
  if (ratio < 1) {
    context2d.fillStyle = '#ffffff';
    context2d.fillRect(0, 0, w, h);
  }
  context2d.drawImage(source, 0, 0, w, h);
  const imageData = context2d.getImageData(0, 0, w, h);
  return { width: w, height: h, data: imageData.data, downscaled: ratio < 1 };
}

async function uiDecodeFile(file) {
  let source = null;
  if (typeof createImageBitmap === 'function') {
    try {
      source = await createImageBitmap(file);
    } catch {
      source = null;
    }
  }
  let rgba;
  if (source) {
    try {
      rgba = await uiRgbaFromSource(source);
    } finally {
      if (source && typeof source.close === 'function') {
        source.close();
      }
    }
  } else {
    const img = await uiDecodeViaImageTag(file);
    rgba = await uiRgbaFromSource(img);
  }
  return rgba;
}

/* --------------------------------------------------------------- processing */

function uiNow() {
  try {
    return typeof performance !== 'undefined' && performance && typeof performance.now === 'function'
      ? performance.now()
      : Date.now();
  } catch {
    return Date.now();
  }
}

/** Any pixel that is not fully opaque? Cheap scan, used to pick the reference. */
function uiImageHasAlpha(image) {
  const data = image && image.data;
  if (!data || !data.length) return false;
  for (let index = 3; index < data.length; index += 4) {
    if (data[index] < 255) return true;
  }
  return false;
}

/**
 * Diff two RGBA buffers while ignoring the given rectangles completely, so the
 * report stays self-consistent when text boxes deliberately drop their ink.
 */
function uiDiffIgnoring(reference, candidate, rects) {
  const width = reference.width;
  const height = reference.height;
  const mask = new Uint8Array(width * height);
  for (const rect of rects || []) {
    const x0 = Math.max(0, Math.round(rect.x));
    const y0 = Math.max(0, Math.round(rect.y));
    const x1 = Math.min(width, Math.round(rect.x + rect.w));
    const y1 = Math.min(height, Math.round(rect.y + rect.h));
    for (let y = y0; y < y1; y++) mask.fill(1, y * width + x0, y * width + x1);
  }
  let pixels = 0;
  let mismatchPixels = 0;
  let sum = 0;
  let maxAbs = 0;
  const a = reference.data;
  const b = candidate.data;
  for (let index = 0; index < mask.length; index++) {
    if (mask[index]) continue;
    const p = index * 4;
    let differs = false;
    for (let channel = 0; channel < 4; channel++) {
      const delta = a[p + channel] > b[p + channel] ? a[p + channel] - b[p + channel] : b[p + channel] - a[p + channel];
      if (delta > 0) differs = true;
      sum += delta;
      if (delta > maxAbs) maxAbs = delta;
    }
    pixels += 1;
    if (differs) mismatchPixels += 1;
  }
  return {
    pixels,
    mismatchPixels,
    exactRatio: pixels ? (pixels - mismatchPixels) / pixels : 1,
    meanAbs: pixels ? sum / (pixels * 4) : 0,
    maxAbs,
  };
}

/**
 * Analyse every item with the same normalised options.
 * Pure enough for Node tests: it only calls the scene engine.
 */
async function uiAnalyze(items, options, onProgress) {
  const out = [];
  const list = items || [];
  for (let index = 0; index < list.length; index += 1) {
    const item = list[index];
    if (typeof onProgress === 'function') {
      onProgress(uiT('正在分析…') + ' ' + (index + 1) + '/' + list.length);
    }
    const started = uiNow();
    try {
      const built = sceneBuild(item.image, options);
      const raster = sceneRasterize(built.scene);
      const textShapes = built.scene.shapes.filter((shape) => shape.k === 't');
      // Diff against what the slide really shows, not against raw pixels:
      // transparent areas become the chosen background colour, and a text box is an
      // empty editable box over its ink, so that ink is not a fidelity failure.
      let reference = item.image;
      if (textShapes.length || uiImageHasAlpha(item.image)) {
        reference = sceneRasterize(sceneBuild(item.image, { ...options, mode: 'image' }).scene);
      }
      const diff = sceneDiff(reference, raster, textShapes.length ? { ignoreRects: textShapes } : undefined);
      let stats = built.scene.stats;
      if (textShapes.length) {
        const exact = uiDiffIgnoring(reference, raster, textShapes);
        stats = {
          ...stats,
          exactRatio: exact.exactRatio,
          mismatchPixels: exact.mismatchPixels,
          meanAbs: exact.meanAbs,
          maxAbs: exact.maxAbs,
          notes: [
            ...(stats.notes || []),
            'fidelity ignores ' + textShapes.length + ' text box(es): their ink is stored as an empty editable box',
          ],
        };
      }
      out.push({
        ...item,
        scene: built.scene,
        scale: built.scale,
        stats,
        raster,
        diff,
        ms: Math.round(uiNow() - started),
        error: null,
      });
    } catch (error) {
      out.push({
        ...item,
        scene: null,
        scale: 1,
        stats: null,
        raster: null,
        diff: null,
        ms: Math.round(uiNow() - started),
        error: (error && error.message) || String(error),
      });
    }
  }
  return out;
}

function uiNewItemId() {
  uiItemSeq += 1;
  return 'pt-' + Date.now().toString(36) + '-' + uiItemSeq;
}

let uiItemSeq = 0;

function uiPreviewFrame(item, view) {
  if (!item) {
    return null;
  }
  if (view === 'original') {
    return { data: item.image.data, width: item.image.width, height: item.image.height };
  }
  if (view === 'diff' && item.diff && item.diff.heatmap) {
    return { data: item.diff.heatmap, width: item.image.width, height: item.image.height };
  }
  if (item.raster) {
    return { data: item.raster.data, width: item.raster.width, height: item.raster.height };
  }
  return { data: item.image.data, width: item.image.width, height: item.image.height };
}

/* ------------------------------------------------------------------- format */

function uiPct(value) {
  const ratio = typeof value === 'number' && Number.isFinite(value) ? value : 1;
  const pct = Math.max(0, Math.min(1, ratio)) * 100;
  if (pct >= 99.99995) {
    return '100%';
  }
  if (pct >= 99) {
    return pct.toFixed(2) + '%';
  }
  return pct.toFixed(1) + '%';
}

function uiNum(value, digits) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return '—';
  }
  return digits == null ? String(Math.round(number)) : number.toFixed(digits);
}

function uiFileSize(bytes) {
  const size = Number(bytes) || 0;
  if (size >= 1024 * 1024) {
    return (size / (1024 * 1024)).toFixed(1) + ' MB';
  }
  if (size >= 1024) {
    return Math.round(size / 1024) + ' KB';
  }
  return size + ' B';
}

/* ---------------------------------------------------------------- small UI */

function uiSection(props) {
  return h(
    'section',
    { className: 'pt-card' },
    h(
      'div',
      { className: 'pt-card-head' },
      h('span', { className: 'pt-card-title' }, props.title),
      props.action || null,
    ),
    props.children,
  );
}

function uiField(props) {
  return h(
    'div',
    { className: 'pt-field' },
    h(
      'label',
      { className: 'pt-label' },
      h('span', null, props.label),
      props.value == null ? null : h('span', { className: 'pt-value' }, props.value),
    ),
    props.children,
    props.hint ? h('span', { className: 'pt-hint' }, props.hint) : null,
  );
}

function uiSegmented(props) {
  return h(
    'div',
    { className: 'pt-seg', role: 'group' },
    (props.options || []).map((option) =>
      h(
        'button',
        {
          key: option.value,
          type: 'button',
          className: option.value === props.value ? 'is-on' : '',
          title: option.title || '',
          onClick: () => props.onChange(option.value),
        },
        option.label,
      ),
    ),
  );
}

function uiToggle(props) {
  return h(
    'label',
    { className: 'pt-check' },
    h('input', {
      type: 'checkbox',
      checked: props.checked === true,
      onChange: (event) => props.onChange(event.target.checked),
    }),
    h(
      'span',
      { className: 'pt-check-text' },
      h('span', { className: 'pt-check-title' }, props.label),
      props.hint ? h('span', { className: 'pt-hint' }, props.hint) : null,
    ),
  );
}

function uiPreviewCanvas(props) {
  const frame = props.frame;
  const canvasRef = useRef(null);
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !frame || !frame.data) {
      return;
    }
    if (canvas.width !== frame.width || canvas.height !== frame.height) {
      canvas.width = frame.width;
      canvas.height = frame.height;
    }
    const context2d = canvas.getContext('2d');
    if (!context2d) {
      return;
    }
    try {
      const pixels = frame.data instanceof Uint8ClampedArray ? frame.data : new Uint8ClampedArray(frame.data);
      if (pixels.length !== frame.width * frame.height * 4) {
        return;
      }
      context2d.putImageData(new ImageData(pixels, frame.width, frame.height), 0, 0);
    } catch {
      /* keep the previous frame */
    }
  }, [frame]);
  if (!frame) {
    return null;
  }
  const style =
    props.zoom === 'actual'
      ? { width: frame.width + 'px', height: frame.height + 'px', maxWidth: 'none', imageRendering: 'pixelated' }
      : { width: '100%', height: 'auto', maxWidth: '100%' };
  return h('canvas', { ref: canvasRef, className: 'pt-canvas', width: frame.width, height: frame.height, style });
}

function uiStat(props) {
  const tone = props.tone === 'good' ? ' is-good' : props.tone === 'bad' ? ' is-bad' : '';
  return h(
    'div',
    { className: 'pt-stat' + tone },
    h('span', { className: 'pt-stat-k' }, props.label),
    h('span', { className: 'pt-stat-v' }, props.value),
  );
}

/* --------------------------------------------------------------- main panel */

function uiInitialState() {
  return {
    items: [],
    selected: 0,
    mode: coreDefaults.mode,
    maxColors: coreDefaults.maxColors,
    mergeTolerance: coreDefaults.mergeTolerance,
    shapeBudget: coreDefaults.shapeBudget,
    detailToImage: coreDefaults.detailToImage,
    textRegions: coreDefaults.textRegions,
    alphaBackground: coreDefaults.alphaBackground,
    fileName: '',
    view: 'rebuilt',
    zoom: 'fit',
    busy: false,
    progress: '',
    dirty: false,
    dragOver: false,
    message: null,
  };
}

function uiPanel(props) {
  const locale = props && props.locale ? props.locale : null;
  const [state, setState] = useState(uiInitialState);
  const stateRef = useRef(state);
  stateRef.current = state;
  const fileInputRef = useRef(null);
  const addFilesRef = useRef(null);
  const messageTimerRef = useRef(null);
  uiUseLanguage(locale);

  const patch = (delta) => setState((current) => ({ ...current, ...delta }));

  const flash = (kind, text) => {
    setState((current) => ({ ...current, message: { kind, text } }));
    if (messageTimerRef.current) {
      clearTimeout(messageTimerRef.current);
    }
    messageTimerRef.current = setTimeout(() => {
      messageTimerRef.current = null;
      setState((current) => (current.message && current.message.text === text ? { ...current, message: null } : current));
    }, 6000);
  };
  useEffect(
    () => () => {
      if (messageTimerRef.current) {
        clearTimeout(messageTimerRef.current);
      }
    },
    [],
  );

  const currentOptions = () =>
    coreNormalizeOptions({
      mode: state.mode,
      maxColors: state.maxColors,
      mergeTolerance: state.mergeTolerance,
      shapeBudget: state.shapeBudget,
      detailToImage: state.detailToImage,
      textRegions: state.textRegions,
      alphaBackground: state.alphaBackground,
    });

  const runAnalysis = async (items, options) => {
    if (!items.length) {
      return items;
    }
    setState((current) => ({ ...current, busy: true, progress: uiT('正在分析…'), message: null }));
    const analysed = await uiAnalyze(items, options, (text) =>
      setState((current) => ({ ...current, progress: text })),
    );
    setState((current) => ({ ...current, items: analysed, busy: false, progress: '', dirty: false }));
    return analysed;
  };

  const addFiles = async (fileLike) => {
    const list = Array.from(fileLike || []).filter(
      (file) => file && (/^image\//.test(file.type || '') || /\.(png|jpe?g|webp|gif|bmp|avif)$/i.test(file.name || '')),
    );
    if (!list.length) {
      return;
    }
    setState((current) => ({ ...current, busy: true, progress: uiT('正在读取图片…'), message: null }));
    const added = [];
    for (const file of list) {
      try {
        const rgba = await uiDecodeFile(file);
        added.push({
          id: uiNewItemId(),
          name: coreSafeName(file.name, 'image'),
          width: rgba.width,
          height: rgba.height,
          image: { width: rgba.width, height: rgba.height, data: rgba.data },
          downscaled: rgba.downscaled === true,
          bytes: file.size || 0,
          scene: null,
          scale: 1,
          raster: null,
          diff: null,
          ms: 0,
          error: null,
        });
      } catch (error) {
        flash('error', (error && error.message) || uiT('无法读取该图片'));
      }
    }
    if (!added.length) {
      setState((current) => ({ ...current, busy: false, progress: '' }));
      return;
    }
    const merged = stateRef.current.items.concat(added);
    const options = currentOptions();
    const analysed = await runAnalysis(merged, options);
    const firstNew = analysed.findIndex((item) => item.id === added[0].id);
    setState((current) => ({
      ...current,
      selected: firstNew >= 0 ? firstNew : current.selected,
      fileName: current.fileName || coreSafeName(added[0].name, 'pixeltrace'),
    }));
  };
  addFilesRef.current = addFiles;

  const removeAt = (index) => {
    setState((current) => {
      const items = current.items.filter((_, itemIndex) => itemIndex !== index);
      let selected = current.selected;
      if (index < current.selected) {
        selected -= 1;
      } else if (index === current.selected) {
        selected = Math.min(index, items.length - 1);
      }
      return { ...current, items, selected: Math.max(0, Math.min(selected, Math.max(0, items.length - 1))) };
    });
  };

  const clearAll = () => {
    setState((current) => ({
      ...current,
      items: [],
      selected: 0,
      fileName: '',
      dirty: false,
      message: null,
      progress: '',
      busy: false,
    }));
  };

  const reanalyse = async () => {
    const items = stateRef.current.items;
    if (!items.length) {
      return items;
    }
    return runAnalysis(items, currentOptions());
  };

  const exportDeck = async () => {
    const snapshot = stateRef.current;
    if (!snapshot.items.length) {
      flash('info', uiT('先添加图片才能导出'));
      return;
    }
    const options = coreNormalizeOptions({
      mode: snapshot.mode,
      maxColors: snapshot.maxColors,
      mergeTolerance: snapshot.mergeTolerance,
      shapeBudget: snapshot.shapeBudget,
      detailToImage: snapshot.detailToImage,
      textRegions: snapshot.textRegions,
      alphaBackground: snapshot.alphaBackground,
    });
    const stale = snapshot.items.some((item) => !item.scene || snapshot.dirty);
    let items = snapshot.items;
    if (stale) {
      items = await runAnalysis(snapshot.items, options);
    }
    const usable = items.filter((item) => item.scene && !item.error);
    if (!usable.length) {
      flash('error', uiT('出错'));
      return;
    }
    setState((current) => ({ ...current, busy: true, progress: uiT('正在生成 PPTX…'), message: null }));
    try {
      const result = await coreBuildDeck(usable, {
        title: snapshot.fileName || usable[0].name,
      });
      const baseName = coreSafeName(snapshot.fileName || usable[0].name, 'pixeltrace');
      const fileName = coreDeckFileName(baseName, usable.length);
      const saved = await coreSaveBytes(
        result.bytes,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      );
      setState((current) => ({ ...current, busy: false, progress: '' }));
      if (saved && saved.cancelled) {
        flash('info', uiT('已取消'));
      } else {
        flash('ok', uiT('已导出') + ' · ' + fileName + ' · ' + uiFileSize(result.bytes.length));
      }
    } catch (error) {
      setState((current) => ({ ...current, busy: false, progress: '' }));
      flash('error', (error && error.message) || uiT('出错'));
    }
  };

  useEffect(() => {
    const onPaste = (event) => {
      const data = event.clipboardData;
      if (!data) {
        return;
      }
      const files = [];
      const items = data.items || [];
      for (let index = 0; index < items.length; index += 1) {
        const entry = items[index];
        if (entry && entry.kind === 'file' && /^image\//.test(entry.type || '')) {
          const file = entry.getAsFile();
          if (file) {
            files.push(file);
          }
        }
      }
      if (!files.length) {
        for (const file of data.files || []) {
          if (/^image\//.test(file.type || '')) {
            files.push(file);
          }
        }
      }
      if (files.length) {
        event.preventDefault();
        if (addFilesRef.current) {
          addFilesRef.current(files);
        }
      }
    };
    document.addEventListener('paste', onPaste);
    return () => document.removeEventListener('paste', onPaste);
  }, []);

  const options = currentOptions();
  const selectedItem = state.items[state.selected] || null;
  const frame = uiPreviewFrame(selectedItem, state.view);
  const stats = selectedItem && (selectedItem.stats || selectedItem.scene) ? selectedItem.stats || selectedItem.scene.stats : null;
  const summary = coreSummary(state.items);
  const canExport = state.items.some((item) => item.scene && !item.error) && !state.busy;
  const modeHint =
    state.mode === 'vector'
      ? uiT('全部变成可编辑矩形：每个色块都能单独改，但形状数量最多')
      : state.mode === 'image'
        ? uiT('每页只放一张图片对象：100% 一致，但不能分层编辑')
        : uiT('平面区域变成可编辑色块，细节区域变成精确裁剪的图片，整页仍逐像素一致');

  return h(
    'div',
    { className: 'pt-root' },
    h(
      'header',
      { className: 'pt-head' },
      h(
        'span',
        { className: 'pt-logo', 'aria-hidden': true },
        h(
          'svg',
          { viewBox: '0 0 24 24', width: 21, height: 21, fill: 'currentColor' },
          h('path', { d: 'M3 3h5v5H3zM3 10h5v5H3zM3 17h5v5H3zM10 3h5v5h-5zM10 10h5v5h-5zM10 17h5v5h-5zM17 3h5v5h-5zM17 10h5v5h-5z', opacity: 0.45 }),
          h('path', { d: 'M17 17h5v5h-5z' }),
        ),
      ),
      h(
        'div',
        null,
        h('h1', { className: 'pt-h1' }, uiT('像素复刻 PixelTrace')),
        h(
          'p',
          { className: 'pt-sub' },
          uiT('把一张图还原成 PPT 里可编辑的形状 · 像素级一致'),
        ),
      ),
    ),
    h('input', {
      ref: fileInputRef,
      type: 'file',
      accept: 'image/*',
      multiple: true,
      className: 'pt-sr',
      onChange: (event) => {
        const files = event.target.files;
        addFiles(files);
        event.target.value = '';
      },
    }),
    h(
      'div',
      {
        className: 'pt-drop' + (state.dragOver ? ' is-over' : ''),
        role: 'button',
        tabIndex: 0,
        onClick: () => {
          if (fileInputRef.current) {
            fileInputRef.current.click();
          }
        },
        onKeyDown: (event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            if (fileInputRef.current) {
              fileInputRef.current.click();
            }
          }
        },
        onDragOver: (event) => {
          event.preventDefault();
          if (!state.dragOver) {
            patch({ dragOver: true });
          }
        },
        onDragLeave: () => {
          if (state.dragOver) {
            patch({ dragOver: false });
          }
        },
        onDrop: (event) => {
          event.preventDefault();
          patch({ dragOver: false });
          if (event.dataTransfer && event.dataTransfer.files) {
            addFiles(event.dataTransfer.files);
          }
        },
      },
      h(
        'span',
        { className: 'pt-drop-icon', 'aria-hidden': true },
        h(
          'svg',
          { viewBox: '0 0 24 24', width: 19, height: 19, fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round' },
          h('path', { d: 'M12 16V4' }),
          h('path', { d: 'm7 9 5-5 5 5' }),
          h('path', { d: 'M4 17v2a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-2' }),
        ),
      ),
      h(
        'div',
        { className: 'pt-drop-spacer' },
        h('div', { className: 'pt-drop-title' }, uiT('拖入图片，或点击选择')),
        h(
          'div',
          { className: 'pt-drop-hint' },
          uiT('PNG · JPEG · WebP · GIF · 可多选（每张一页）· 也可以直接 Ctrl+V 粘贴'),
        ),
      ),
      h(
        'button',
        {
          type: 'button',
          className: 'pt-btn pt-primary',
          disabled: state.busy,
          onClick: (event) => {
            event.stopPropagation();
            if (fileInputRef.current) {
              fileInputRef.current.click();
            }
          },
        },
        uiT('选择图片'),
      ),
    ),
    h(
      'div',
      { className: 'pt-form' },
      h(
        'aside',
        { className: 'pt-side' },
        h(
          uiSection,
          {
            title: uiT('图片'),
            action: state.items.length
              ? h(
                  'button',
                  { type: 'button', className: 'pt-btn pt-ghost pt-small', onClick: clearAll, disabled: state.busy },
                  uiT('清空'),
                )
              : null,
          },
          state.items.length
            ? h(
                'div',
                { className: 'pt-files' },
                state.items.map((item, index) =>
                  h(
                    'div',
                    {
                      key: item.id,
                      className: 'pt-file' + (index === state.selected ? ' is-on' : ''),
                      role: 'button',
                      tabIndex: 0,
                      onClick: () => patch({ selected: index }),
                      onKeyDown: (event) => {
                        if (event.key === 'Enter') {
                          patch({ selected: index });
                        }
                      },
                    },
                    h(
                      'span',
                      { className: 'pt-file-body' },
                      h('span', { className: 'pt-file-name', title: item.name }, item.name),
                      h(
                        'span',
                        { className: 'pt-file-meta' },
                        item.width +
                          '×' +
                          item.height +
                          (item.scene ? ' · ' + item.scene.stats.shapes + ' ' + uiT('对象') : '') +
                          (item.error ? ' · ' + uiT('出错') : ''),
                      ),
                    ),
                    h(
                      'button',
                      {
                        type: 'button',
                        className: 'pt-x',
                        title: uiT('移除'),
                        onClick: (event) => {
                          event.stopPropagation();
                          removeAt(index);
                        },
                      },
                      '×',
                    ),
                  ),
                ),
              )
            : h('div', { className: 'pt-empty' }, uiT('还没有图片')),
        ),
        h(
          uiSection,
          { title: uiT('输出参数') },
          h(
            uiField,
            { label: uiT('输出模式'), hint: modeHint },
            h(uiSegmented, {
              value: state.mode,
              options: [
                { value: 'smart', label: uiT('智能分层'), title: uiT('智能分层') },
                { value: 'vector', label: uiT('纯矢量'), title: uiT('纯矢量') },
                { value: 'image', label: uiT('整图保真'), title: uiT('整图保真') },
              ],
              onChange: (value) => patch({ mode: value, dirty: !!state.items.length }),
            }),
          ),
          h(
            uiField,
            { label: uiT('形状预算/页'), value: String(state.shapeBudget), hint: uiT('0 = 不限制；超出预算时自动降级并在统计里说明') },
            h('input', {
              type: 'number',
              className: 'pt-input',
              min: 0,
              max: 200000,
              step: 500,
              value: state.shapeBudget,
              onChange: (event) => patch({ shapeBudget: event.target.value, dirty: !!state.items.length }),
            }),
          ),
          h(
            uiField,
            { label: uiT('颜色容差'), value: String(state.mergeTolerance), hint: uiT('合并相邻矩形时允许的单通道色差，0 = 完全精确') },
            h('input', {
              type: 'range',
              min: 0,
              max: 64,
              step: 1,
              value: state.mergeTolerance,
              style: { width: '100%', accentColor: 'var(--dsw-alias-brand-primary,#3b82f6)' },
              onChange: (event) => patch({ mergeTolerance: event.target.value, dirty: !!state.items.length }),
            }),
          ),
          h(
            uiField,
            { label: uiT('颜色量化'), hint: uiT('0 = 保留原始颜色') },
            h(
              'select',
              {
                className: 'pt-select',
                value: String(state.maxColors),
                onChange: (event) => patch({ maxColors: Number(event.target.value), dirty: !!state.items.length }),
              },
              [0, 16, 32, 64, 128, 256].map((count) =>
                h('option', { key: count, value: String(count) }, count === 0 ? uiT('0 = 保留原始颜色') : String(count)),
              ),
            ),
          ),
          h(uiToggle, {
            label: uiT('细节区域转为精确裁剪图'),
            checked: state.detailToImage,
            onChange: (checked) => patch({ detailToImage: checked, dirty: !!state.items.length }),
          }),
          h(uiToggle, {
            label: uiT('文字识别为可编辑文本框'),
            hint: uiT('实验性：文字墨迹被擦成背景色，留下一个空文本框，方便你自己输入原文'),
            checked: state.textRegions,
            onChange: (checked) => patch({ textRegions: checked, dirty: !!state.items.length }),
          }),
          h(
            uiField,
            { label: uiT('透明背景色') },
            h(
              'div',
              { className: 'pt-row' },
              h('input', {
                type: 'color',
                className: 'pt-swatch',
                value: options.alphaBackground,
                onChange: (event) => patch({ alphaBackground: event.target.value, dirty: !!state.items.length }),
              }),
              h('input', {
                type: 'text',
                className: 'pt-input',
                value: state.alphaBackground,
                spellCheck: false,
                onChange: (event) => patch({ alphaBackground: event.target.value, dirty: !!state.items.length }),
              }),
            ),
          ),
        ),
      ),
      h(
        'main',
        { className: 'pt-stage' },
        h(
          'div',
          { className: 'pt-stage-bar' },
          h(
            'span',
            { className: 'pt-stage-title' },
            selectedItem
              ? selectedItem.name + ' · ' + selectedItem.width + '×' + selectedItem.height
              : uiT('预览'),
          ),
          h(uiSegmented, {
            value: state.view,
            options: [
              { value: 'original', label: uiT('原图') },
              { value: 'rebuilt', label: uiT('重建') },
              { value: 'diff', label: uiT('差异') },
            ],
            onChange: (value) => patch({ view: value }),
          }),
          h(uiSegmented, {
            value: state.zoom,
            options: [
              { value: 'fit', label: uiT('适应') },
              { value: 'actual', label: '1:1' },
            ],
            onChange: (value) => patch({ zoom: value }),
          }),
          h(
            'button',
            { type: 'button', className: 'pt-btn pt-small', onClick: reanalyse, disabled: state.busy || !state.items.length },
            uiT('重新分析'),
          ),
        ),
        h(
          'div',
          { className: 'pt-canvas-wrap' },
          frame
            ? h(uiPreviewCanvas, { frame, zoom: state.zoom })
            : h('div', { className: 'pt-empty' }, uiT('拖入图片后在这里查看逐像素预览')),
        ),
        state.dirty
          ? h('div', { className: 'pt-warn' }, uiT('参数已修改，点「重新分析」刷新预览'))
          : null,
        selectedItem && selectedItem.error
          ? h('div', { className: 'pt-msg error' }, selectedItem.error)
          : null,
        selectedItem && selectedItem.scene
          ? h(
              uiSection,
              { title: uiT('保真统计') },
              h(
                'div',
                { className: 'pt-stats' },
                h(uiStat, { label: uiT('完全一致像素'), value: uiPct(stats.exactRatio), tone: stats.exactRatio >= 1 ? 'good' : 'bad' }),
                h(uiStat, { label: uiT('不一致像素'), value: uiNum(stats.mismatchPixels) }),
                h(uiStat, { label: uiT('对象'), value: uiNum(stats.shapes) }),
                h(uiStat, { label: uiT('矩形'), value: uiNum(stats.rects) }),
                h(uiStat, { label: uiT('裁剪图'), value: uiNum(stats.images) }),
                h(uiStat, { label: uiT('文本框'), value: uiNum(stats.texts) }),
                h(uiStat, { label: uiT('颜色'), value: uiNum(stats.paletteSize) }),
                h(uiStat, { label: uiT('平均误差'), value: uiNum(stats.meanAbs, 2) }),
                h(uiStat, { label: uiT('最大误差'), value: uiNum(stats.maxAbs) }),
                h(uiStat, { label: uiT('耗时'), value: uiNum(selectedItem.ms) + ' ms' }),
              ),
              selectedItem.downscaled
                ? h('div', { className: 'pt-notes' }, h('span', null, uiT('图片过大，已降采样以获得可用速度')))
                : null,
              stats.notes && stats.notes.length
                ? h(
                    'div',
                    { className: 'pt-notes' },
                    h('span', null, uiT('备注') + ':'),
                    stats.notes.map((note, index) => h('code', { key: index }, note)),
                  )
                : null,
            )
          : null,
        state.items.length > 1
          ? h(
              'div',
              { className: 'pt-hint' },
              summary.slides + ' ' + uiT('页') + ' · ' + summary.shapes + ' ' + uiT('对象') + ' · ' + uiPct(summary.exactRatio) + ' · ' + uiNum(summary.ms) + ' ms',
            )
          : null,
      ),
    ),
    state.busy && state.progress
      ? h('div', { className: 'pt-msg info' }, h('span', { className: 'pt-spin' }), state.progress)
      : null,
    state.message
      ? h('div', { className: 'pt-msg ' + (state.message.kind || 'info') }, state.message.text)
      : null,
    h(
      'footer',
      { className: 'pt-foot' },
      h(
        'div',
        { className: 'pt-foot-name' },
        h('span', { className: 'pt-hint' }, uiT('文件名')),
        h('input', {
          type: 'text',
          className: 'pt-input',
          value: state.fileName,
          spellCheck: false,
          onChange: (event) => patch({ fileName: event.target.value }),
        }),
      ),
      h(
        'button',
        { type: 'button', className: 'pt-btn', onClick: reanalyse, disabled: state.busy || !state.items.length },
        uiT('重新分析'),
      ),
      h(
        'button',
        { type: 'button', className: 'pt-btn pt-primary', onClick: exportDeck, disabled: !canExport },
        state.busy && state.progress === uiT('正在生成 PPTX…') ? uiT('正在生成…') : uiT('导出 PPTX'),
      ),
      state.items.length
        ? h('span', { className: 'pt-hint' }, state.items.filter((item) => item.scene).length + ' ' + uiT('页'))
        : null,
      h('span', { className: 'pt-hint', style: { flex: '1 1 auto' } }, ''),
      h('span', { className: 'pt-hint' }, uiT('提示：导出的每个色块、裁剪图、文本框都是独立对象，可在 PowerPoint / WPS / Keynote 里单独编辑。')),
    ),
  );
}

function uiSidebarIcon(props) {
  const raw = props && props.size != null ? props.size : 20;
  const size = typeof raw === 'number' ? raw : parseFloat(raw) || 20;
  return h(
    'svg',
    {
      viewBox: '0 0 24 24',
      width: size,
      height: size,
      fill: 'currentColor',
      'aria-hidden': true,
      style: { display: 'block' },
    },
    h('path', {
      d: 'M3 3h5v5H3zM3 10h5v5H3zM3 17h5v5H3zM10 3h5v5h-5zM10 10h5v5h-5zM10 17h5v5h-5zM17 3h5v5h-5zM17 10h5v5h-5z',
      opacity: 0.45,
    }),
    h('path', { d: 'M17 17h5v5h-5z' }),
  );
}

/** Register the sidebar entry and the main panel. */
function uiApply(ctx) {
  let localeService = null;
  try {
    localeService = (ctx && typeof ctx.get === 'function' && ctx.get('locale')) || (ctx && ctx.locale) || null;
  } catch {
    localeService = (ctx && ctx.locale) || null;
  }
  uiSetLocaleService(localeService);

  if (typeof ctx.effect === 'function') {
    ctx.effect(() => uiMountStyle());
  } else {
    uiMountStyle();
  }

  const openPanel = () => {
    try {
      ctx.layout.selectPanel(uiPanelId);
    } catch {
      /* layout service may not be ready yet */
    }
  };

  ctx.slots.inject('main', () =>
    ctx.slots.inject('sidebar.panellist', () => {
      const stopMain = ctx.slots.register(
        {
          name: 'main',
          key: uiPanelId,
          inject: () => ({ locale: localeService, openPanel }),
        },
        uiPanel,
      );
      const stopIcon = ctx.slots.register(
        {
          name: 'sidebar.panellist',
          id: uiPanelId,
          order: uiSidebarOrder,
          label: () => (uiLanguageOf() === 'en' ? 'PixelTrace' : '像素复刻'),
        },
        uiSidebarIcon,
      );
      return () => {
        stopMain();
        stopIcon();
      };
    }),
  );
}
