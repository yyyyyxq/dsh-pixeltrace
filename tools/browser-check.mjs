#!/usr/bin/env node
/**
 * Headless browser smoke test for the PixelTrace client plugin.
 *
 *   node tools/browser-check.mjs "http://127.0.0.1:PORT/?token=..." [--keep]
 *
 * What it does:
 *   1. launches headless Microsoft Edge with a throw-away profile
 *   2. opens the authenticated DSH Web GUI URL (token in the query string)
 *   3. waits for the sidebar, clicks the PixelTrace panel entry
 *   4. drops a synthetic test image (canvas drawn PNG) onto the drop zone
 *   5. waits for the analysis, screenshots the panel, reads the stats
 *   6. clicks "export" with the file picker stubbed out, captures the .pptx
 *      download and validates it with @office-kit/pptx
 *
 * Everything it produces lands in out/shots/. Zero npm dependencies: the CDP
 * client is a bare WebSocket built on the Node 22 global.
 */
import { spawn } from "node:child_process";
import { mkdirSync, writeFileSync, readFileSync, readdirSync, existsSync, rmSync } from "node:fs";
import { join, resolve } from "node:path";
import { setTimeout as sleep } from "node:timers/promises";
import { repoRoot } from "../test/harness.mjs";

const EDGE = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const OUT = join(repoRoot, "out", "shots");
const DOWNLOADS = join(repoRoot, "out", "downloads");

function fail(message) {
	console.error(`browser-check: ${message}`);
	process.exit(1);
}

const args = process.argv.slice(2);
const url = args.find((a) => a.startsWith("http"));
if (!url) fail("pass the authenticated web URL as the first argument");
const keep = args.includes("--keep");
/** Optional real screenshot to drop instead of the built-in canvas image. */
const imageIndex = args.indexOf("--image");
const imagePath = imageIndex >= 0 ? args[imageIndex + 1] : null;
if (imageIndex >= 0 && (!imagePath || !existsSync(imagePath))) fail(`--image needs an existing file, got ${imagePath}`);
const mimeOf = (file) => (/\.jpe?g$/i.test(file) ? "image/jpeg" : /\.webp$/i.test(file) ? "image/webp" : "image/png");
const dataUrlOf = (file) => `data:${mimeOf(file)};base64,${readFileSync(file).toString("base64")}`;

class CDP {
	constructor(socket) {
		this.socket = socket;
		this.nextId = 1;
		this.pending = new Map();
		this.listeners = new Set();
		socket.addEventListener("message", (event) => {
			const message = JSON.parse(String(event.data));
			if (message.id !== undefined) {
				const pending = this.pending.get(message.id);
				if (!pending) return;
				this.pending.delete(message.id);
				if (message.error) pending.reject(new Error(`${pending.method}: ${message.error.message}`));
				else pending.resolve(message.result);
				return;
			}
			for (const listener of this.listeners) listener(message);
		});
	}
	send(method, params = {}, sessionId) {
		const id = this.nextId++;
		const payload = { id, method, params };
		if (sessionId !== undefined) payload.sessionId = sessionId;
		return new Promise((resolve, reject) => {
			this.pending.set(id, { resolve, reject, method });
			this.socket.send(JSON.stringify(payload));
		});
	}
	on(listener) {
		this.listeners.add(listener);
		return () => this.listeners.delete(listener);
	}
}

async function launchEdge(port) {
	const profile = join(OUT, "edge-profile");
	rmSync(profile, { recursive: true, force: true });
	const child = spawn(EDGE, [
		"--headless=new",
		`--remote-debugging-port=${port}`,
		`--user-data-dir=${profile}`,
		"--no-first-run",
		"--no-default-browser-check",
		"--disable-gpu",
		"--hide-scrollbars",
		"--window-size=1600,1000",
		"--force-device-scale-factor=1",
		"about:blank",
	], { stdio: ["ignore", "pipe", "pipe"] });
	let stderr = "";
	child.stderr.setEncoding("utf8");
	child.stderr.on("data", (chunk) => { stderr += chunk; });
	for (let attempt = 0; attempt < 100; attempt++) {
		try {
			const response = await fetch(`http://127.0.0.1:${port}/json/version`);
			if (response.ok) {
				const info = await response.json();
				if (typeof info.webSocketDebuggerUrl === "string") return { child, wsUrl: info.webSocketDebuggerUrl };
			}
		} catch {}
		await sleep(100);
	}
	child.kill();
	fail(`headless Edge did not expose CDP: ${stderr.slice(-400)}`);
}

/** Builds the test image inside the page and returns its base64 PNG. */
const DRAW_TEST_IMAGE = `(() => {
	const c = document.createElement('canvas');
	c.width = 800; c.height = 500;
	const g = c.getContext('2d');
	g.fillStyle = '#ffffff'; g.fillRect(0, 0, 800, 500);
	g.fillStyle = '#1b4f9c'; g.fillRect(0, 0, 800, 56);
	g.fillStyle = '#ffffff'; g.font = 'bold 26px "Segoe UI", Arial'; g.textBaseline = 'middle';
	g.fillText('PixelTrace fidelity test', 24, 28);
	g.fillStyle = '#eef2f8'; g.fillRect(0, 56, 240, 444);
	g.fillStyle = '#22303f'; g.font = '16px "Segoe UI", Arial';
	g.fillText('Sidebar', 24, 100);
	g.fillText('Item one', 24, 140);
	g.fillText('Item two', 24, 176);
	const grad = g.createLinearGradient(280, 90, 760, 300);
	grad.addColorStop(0, '#ff7a18'); grad.addColorStop(1, '#af002d');
	g.fillStyle = grad; g.fillRect(280, 90, 480, 120);
	g.fillStyle = '#2bb673'; g.beginPath(); g.arc(380, 320, 70, 0, Math.PI * 2); g.fill();
	g.fillStyle = '#f5c518'; g.fillRect(500, 250, 260, 140);
	g.fillStyle = '#111111'; g.font = 'bold 34px "Segoe UI", Arial';
	g.fillText('Hello PPT', 520, 300);
	g.font = '18px "Segoe UI", Arial'; g.fillText('editable text region', 520, 340);
	g.strokeStyle = '#5b6b7c'; g.lineWidth = 2;
	for (let i = 0; i < 12; i++) { g.beginPath(); g.moveTo(280, 400 + i * 6); g.lineTo(760, 400 + i * 6); g.stroke(); }
	return c.toDataURL('image/png');
})()`;

const DROP_TEST_IMAGE = (dataUrl) => `(() => {
	const bytes = Uint8Array.from(atob(${JSON.stringify(dataUrl)}.split(',')[1]), (ch) => ch.charCodeAt(0));
	const file = new File([bytes], 'pixeltrace-test.png', { type: 'image/png' });
	const zone = document.querySelector('.pt-drop');
	if (!zone) return { ok: false, reason: 'drop zone missing' };
	const dataTransfer = new DataTransfer();
	dataTransfer.items.add(file);
	for (const type of ['dragenter', 'dragover', 'drop']) {
		const event = new Event(type, { bubbles: true, cancelable: true });
		Object.defineProperty(event, 'dataTransfer', { value: dataTransfer });
		zone.dispatchEvent(event);
	}
	return { ok: true };
})()`;

async function main() {
	mkdirSync(OUT, { recursive: true });
	mkdirSync(DOWNLOADS, { recursive: true });
	for (const leftover of readdirSync(DOWNLOADS)) rmSync(join(DOWNLOADS, leftover), { force: true, recursive: true });

	const { child, wsUrl } = await launchEdge(9331);
	const socket = new WebSocket(wsUrl);
	await new Promise((resolve, reject) => {
		socket.addEventListener("open", resolve, { once: true });
		socket.addEventListener("error", () => reject(new Error("CDP socket failed")), { once: true });
	});
	const cdp = new CDP(socket);
	const problems = [];
	cdp.on((message) => {
		if (message.method === "Runtime.exceptionThrown") {
			const detail = message.params?.exceptionDetails;
			problems.push(`exception: ${detail?.exception?.description ?? detail?.text ?? "unknown"}`);
		}
		if (message.method === "Runtime.consoleAPICalled" && message.params?.type === "error") {
			problems.push(`console.error: ${(message.params.args ?? []).map((a) => a.value ?? a.description ?? "").join(" ").slice(0, 300)}`);
		}
	});

	const { targetId } = await cdp.send("Target.createTarget", { url: "about:blank" });
	const { sessionId } = await cdp.send("Target.attachToTarget", { targetId, flatten: true });
	const page = (method, params) => cdp.send(method, params, sessionId);
	await page("Page.enable");
	await page("Runtime.enable");
	await page("Log.enable").catch(() => {});
	await page("Browser.setDownloadBehavior", { behavior: "allow", downloadPath: DOWNLOADS }).catch(() => {});

	const evaluate = async (expression, awaitPromise = false) => {
		const result = await page("Runtime.evaluate", { expression, returnByValue: true, awaitPromise, timeout: 120000 });
		if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
		return result.result?.value;
	};
	const shot = async (name) => {
		const { data } = await page("Page.captureScreenshot", { format: "png", captureBeyondViewport: false });
		const file = join(OUT, `${name}.png`);
		writeFileSync(file, Buffer.from(data, "base64"));
		return file;
	};

	const report = { url, problems, sidebarFound: false, panelOpened: false, analysis: null, exported: null, shots: [] };

	await page("Page.navigate", { url });
	// Wait until the shell rendered something meaningful.
	for (let attempt = 0; attempt < 120; attempt++) {
		await sleep(500);
		const text = (await evaluate("document.body ? document.body.innerText : ''")) ?? "";
		if (text.length > 120) break;
	}
	report.shots.push(await shot("01-shell"));

	const sidebar = await evaluate(`(() => {
		const nodes = Array.from(document.querySelectorAll('button, [role="button"], a'));
		return nodes.map((node) => ({
			label: (node.getAttribute('aria-label') || node.getAttribute('title') || node.textContent || '').trim().slice(0, 60),
			cls: typeof node.className === 'string' ? node.className.slice(0, 80) : ''
		})).filter((entry) => entry.label !== '');
	})()`);
	report.sidebarEntries = sidebar;

	const clicked = await evaluate(`(() => {
		const wanted = ['像素复刻', 'PixelTrace', 'pixeltrace'];
		const nodes = Array.from(document.querySelectorAll('button, [role="button"], a'));
		for (const node of nodes) {
			const label = (node.getAttribute('aria-label') || node.getAttribute('title') || node.textContent || '').trim();
			if (wanted.some((needle) => label === needle || label.includes(needle))) {
				node.click();
				return { clicked: true, label };
			}
		}
		return { clicked: false };
	})()`);
	report.sidebarFound = clicked.clicked === true;
	report.sidebarLabel = clicked.label ?? null;
	await sleep(2500);
	report.panelOpened = (await evaluate("!!document.querySelector('.pt-root')")) === true;
	report.shots.push(await shot("02-panel"));

	if (!report.panelOpened) {
		report.bodyText = ((await evaluate("document.body.innerText")) ?? "").slice(0, 4000);
	} else {
		const dataUrl = imagePath ? dataUrlOf(imagePath) : await evaluate(DRAW_TEST_IMAGE);
		const dropped = await evaluate(DROP_TEST_IMAGE(dataUrl));
		report.dropped = dropped;
		for (let attempt = 0; attempt < 240; attempt++) {
			await sleep(500);
			const busy = await evaluate("!!document.querySelector('.pt-busy')");
			const stats = await evaluate(`(() => {
				const node = document.querySelector('.pt-stats');
				return node ? node.innerText : '';
			})()`);
			if (!busy && stats && /%/.test(stats)) {
				report.analysis = stats;
				break;
			}
		}
		report.shots.push(await shot("03-analysis"));

		// Optional: flip "text regions become editable text boxes" on and re-read the stats.
		if (args.includes("--text")) {
			const toggled = await evaluate(`(() => {
				const labels = Array.from(document.querySelectorAll('.pt-root label'));
				const target = labels.find((node) => /文字/.test(node.textContent || ''));
				if (!target) return { clicked: false, labels: labels.map((node) => (node.textContent || '').trim().slice(0, 24)) };
				const input = target.querySelector('input[type=checkbox]') || target.querySelector('input');
				if (!input) return { clicked: false, reason: 'no input in label' };
				input.click();
				return { clicked: true, text: (target.textContent || '').trim().slice(0, 40), checked: !!input.checked };
			})()`);
			report.textToggle = toggled;
			// Changing a parameter only marks the panel dirty; re-analysis is explicit.
			await evaluate(`(() => {
				const buttons = Array.from(document.querySelectorAll('.pt-root button'));
				const target = buttons.find((node) => /重新分析|Re-analyse|Re-analyze/i.test(node.textContent || ''));
				if (target) target.click();
				return !!target;
			})()`);
			for (let attempt = 0; attempt < 240; attempt++) {
				await sleep(500);
				const busy = await evaluate("!!document.querySelector('.pt-busy')");
				const stats = await evaluate(`(() => {
					const node = document.querySelector('.pt-stats');
					return node ? node.innerText : '';
				})()`);
				if (!busy && stats && /%/.test(stats)) {
					report.textAnalysis = stats;
					break;
				}
			}
			report.shots.push(await shot("05-text-regions"));
		}

		// Force the plain-download fallback and press the export button.
		await evaluate("window.showSaveFilePicker = undefined; true");
		const exported = await evaluate(`(() => {
			const buttons = Array.from(document.querySelectorAll('.pt-foot button, .pt-btn'));
			const target = buttons.find((node) => /导出|Export/.test(node.textContent || ''));
			if (!target) return { clicked: false, text: buttons.map((n) => (n.textContent || '').trim()) };
			target.click();
			return { clicked: true, text: (target.textContent || '').trim() };
		})()`);
		report.exportClicked = exported;
		for (let attempt = 0; attempt < 60; attempt++) {
			await sleep(500);
			const files = readdirSync(DOWNLOADS).filter((name) => name.endsWith(".pptx"));
			if (files.length > 0) {
				report.exported = join(DOWNLOADS, files[0]);
				break;
			}
		}
		report.shots.push(await shot("04-after-export"));
		report.footText = ((await evaluate(`(() => { const node = document.querySelector('.pt-foot'); return node ? node.innerText : ''; })()`)) ?? "").slice(0, 600);
	}

	// Validate the produced file with the independent OOXML reader.
	if (report.exported) {
		try {
			const office = await import("file:///C:/Users/yxq/.dsh/profiles/web/node_modules/@office-kit/pptx/dist/index.js");
			const bytes = readFileSync(report.exported);
			const presentation = await office.loadPresentation(new Uint8Array(bytes));
			const issues = office.validatePresentation(presentation) ?? [];
			report.ooxml = {
				bytes: bytes.byteLength,
				slides: office.getSlides(presentation).length,
				shapes: office.getSlides(presentation).map((slide) => office.getSlideShapes(slide).length),
				issues: issues.map((issue) => String(issue.message ?? issue)).slice(0, 8),
			};
		} catch (error) {
			report.ooxml = { error: String(error?.message ?? error) };
		}
	}

	writeFileSync(join(OUT, "report.json"), JSON.stringify(report, null, 2));
	console.log(JSON.stringify(report, null, 2));

	if (!keep) {
		socket.close();
		child.kill();
	}
}

main().catch((error) => {
	console.error(error);
	process.exit(1);
});
