#!/usr/bin/env node
/* eslint-disable no-console */
import fs from "fs";
import path from "path";
import { performance } from "perf_hooks";
import { fileURLToPath } from "url";
import XLSX from "xlsx";
import { chromium } from "playwright";
const ROOT = path.dirname(fileURLToPath(import.meta.url));
const SCENARIO_PATH = path.join(ROOT, "scenarios", "polaria-simulaciones.json");
const EVIDENCE_PATH = path.join(ROOT, "evidencia-entrega.png");
const DIAGNOSTICS_DIR = path.join(ROOT, "diagnosticos");

const DEFAULT_BROWSER_PATHS = {
  opera: "C:\\Program Files\\Opera GX\\opera.exe",
  chrome: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  edge: "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
};

function parseArgs(argv) {
  const out = {
    headed: true,
    frontUrl: "",
    downloadsDir: "",
    timeoutMs: 90_000
  };

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "--help") {
      printHelp();
      process.exit(0);
    } else if (token === "--headed") {
      out.headed = true;
    } else if (token === "--headless") {
      out.headed = false;
    } else if (token === "--front-url") {
      out.frontUrl = argv[++i] || "";
    } else if (token === "--downloads") {
      out.downloadsDir = argv[++i] || "";
    } else if (token === "--timeout-ms") {
      out.timeoutMs = Number.parseInt(argv[++i] || "", 10);
    }
  }
  return out;
}

function printHelp() {
  console.log(`
Runner externo UI Polaria

Uso:
  node run.mjs --downloads "C:\\Users\\TU_USUARIO\\Downloads" [opciones]

Opciones:
  --front-url <url>      URL frontend (default: la del escenario)
  --downloads <path>     Carpeta donde están Andino(1).xlsx, Mar_Azul(1).xlsx, Aves_Dorada(1).xlsx
  --headed               Navegadores visibles (default)
  --headless             Navegadores ocultos
  --timeout-ms <n>       Timeout base por acción (default 35000)
  --help                 Mostrar ayuda
`);
}

function normalizePathFromFileUrl(p) {
  if (process.platform === "win32") return p;
  return p;
}

function readScenario() {
  const raw = fs.readFileSync(SCENARIO_PATH, "utf8");
  return JSON.parse(raw);
}

function ensureFile(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`No existe ${label}: ${filePath}`);
  }
}

function parseCatalogMeta(excelPath) {
  const workbook = XLSX.readFile(excelPath);
  const firstSheet = workbook.SheetNames[0];
  const rows = firstSheet
    ? XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" })
    : [];
  return { sheet: firstSheet || "(sin hoja)", rows: rows.length };
}

function ensureEvidenceImage() {
  if (fs.existsSync(EVIDENCE_PATH)) return EVIDENCE_PATH;
  const pngBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
  fs.writeFileSync(EVIDENCE_PATH, Buffer.from(pngBase64, "base64"));
  return EVIDENCE_PATH;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function formatRegexList(regexes) {
  return regexes.map((rx) => String(rx)).join(", ");
}

function rxContains(value) {
  return new RegExp(escapeRegExp(value), "i");
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function todayInput() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

const SLOW_FACTOR = Math.max(
  1,
  Number.parseFloat(process.env.POLARIA_SLOW_FACTOR || "2.4") || 2.4
);
const SLOW_MO_MS = Math.max(
  0,
  Number.parseInt(process.env.POLARIA_SLOW_MO_MS || "350", 10) || 350
);

function sleep(ms) {
  const scaled = Math.ceil(Math.max(0, ms) * SLOW_FACTOR);
  return new Promise((resolve) => setTimeout(resolve, scaled));
}


const RUN_DIAG = {
  startedAt: new Date().toISOString(),
  meta: {},
  sims: new Map(),
  events: [],
  finalized: false,

  setMeta(meta) {
    this.meta = meta || {};
  },

  registerSimulation(sim, browserType) {
    if (!sim || !sim.id) return;
    if (this.sims.has(sim.id)) return;
    this.sims.set(sim.id, {
      id: sim.id,
      nombre: sim.nombre || sim.id,
      browserType: browserType || "desconocido",
      estado: "pendiente",
      total: 0,
      ok: 0,
      pasoActual: "-",
      startedAt: null,
      endedAt: null,
      error: null,
      pasos: []
    });
  },

  setTotal(simId, total) {
    const s = this.sims.get(simId);
    if (!s) return;
    s.total = total;
  },

  stepStart(simId, paso) {
    const s = this.sims.get(simId);
    if (!s) return;
    const now = new Date().toISOString();
    if (!s.startedAt) s.startedAt = now;
    s.estado = "ejecutando";
    s.pasoActual = paso;
    s.pasos.push({ paso, estado: "ejecutando", startedAt: now, endedAt: null, error: null });
    this.events.push({ at: now, simId, tipo: "step_start", paso });
  },

  stepOk(simId, paso) {
    const s = this.sims.get(simId);
    if (!s) return;
    const now = new Date().toISOString();
    s.ok += 1;
    for (let i = s.pasos.length - 1; i >= 0; i -= 1) {
      if (s.pasos[i].paso === paso && s.pasos[i].estado === "ejecutando") {
        s.pasos[i].estado = "ok";
        s.pasos[i].endedAt = now;
        break;
      }
    }
    this.events.push({ at: now, simId, tipo: "step_ok", paso });
  },

  stepError(simId, paso, error) {
    const s = this.sims.get(simId);
    if (!s) return;
    const now = new Date().toISOString();
    const msg = error?.message || String(error || "error");
    for (let i = s.pasos.length - 1; i >= 0; i -= 1) {
      if (s.pasos[i].paso === paso && s.pasos[i].estado === "ejecutando") {
        s.pasos[i].estado = "error";
        s.pasos[i].endedAt = now;
        s.pasos[i].error = msg;
        break;
      }
    }
    this.events.push({ at: now, simId, tipo: "step_error", paso, error: msg });
  },

  completeSimulation(simId) {
    const s = this.sims.get(simId);
    if (!s) return;
    s.estado = "ok";
    s.pasoActual = "finalizado";
    s.endedAt = new Date().toISOString();
  },

  failSimulation(simId, error) {
    const s = this.sims.get(simId);
    if (!s) return;
    s.estado = "error";
    s.error = error?.stack || error?.message || String(error || "error");
    s.endedAt = new Date().toISOString();
  },

  writeFinal(results, globalError) {
    if (this.finalized) return null;
    this.finalized = true;

    fs.mkdirSync(DIAGNOSTICS_DIR, { recursive: true });

    const now = new Date();
    const stamp =
      String(now.getFullYear()) +
      String(now.getMonth() + 1).padStart(2, "0") +
      String(now.getDate()).padStart(2, "0") + "-" +
      String(now.getHours()).padStart(2, "0") +
      String(now.getMinutes()).padStart(2, "0") +
      String(now.getSeconds()).padStart(2, "0");

    const endedAt = new Date().toISOString();
    const sims = [...this.sims.values()];
    const payload = {
      run: {
        startedAt: this.startedAt,
        endedAt,
        frontUrl: this.meta.frontUrl,
        downloadsDir: this.meta.downloadsDir,
        headed: this.meta.headed,
        timeoutMs: this.meta.timeoutMs,
        scenarioPath: this.meta.scenarioPath
      },
      summary: {
        total: sims.length,
        ok: sims.filter((x) => x.estado === "ok").length,
        failed: sims.filter((x) => x.estado === "error").length
      },
      simulations: sims,
      results: results || [],
      globalError: globalError?.stack || globalError?.message || (globalError ? String(globalError) : null),
      events: this.events
    };

    const md = [];
    md.push("# Diagnóstico de ejecución - Polaria UI Runner");
    md.push("");
    md.push("## Resumen");
    md.push("");
    md.push("- Inicio: " + payload.run.startedAt);
    md.push("- Fin: " + payload.run.endedAt);
    md.push("- Frontend: " + (payload.run.frontUrl || "-"));
    md.push("- Descargas: " + (payload.run.downloadsDir || "-"));
    md.push("");
    md.push("| Simulación | Navegador | Estado | Progreso | Paso actual |");
    md.push("|---|---|---|---:|---|");
    for (const s of sims) {
      md.push("| " + s.id + " | " + s.browserType + " | " + s.estado + " | " + s.ok + "/" + (s.total || "?") + " | " + s.pasoActual + " |");
    }
    md.push("");

    for (const s of sims) {
      md.push("## " + s.id + " (" + s.nombre + ")");
      md.push("");
      md.push("- Estado final: " + s.estado);
      md.push("- Inicio: " + (s.startedAt || "-"));
      md.push("- Fin: " + (s.endedAt || "-"));
      if (s.error) md.push("- Error final: " + s.error.replace(/\n/g, " "));
      md.push("");
      md.push("| Paso | Estado | Error |");
      md.push("|---|---|---|");
      for (const p of s.pasos) {
        md.push("| " + p.paso + " | " + p.estado + " | " + (p.error || "") + " |");
      }
      md.push("");
    }

    if (payload.globalError) {
      md.push("## Error global");
      md.push("");
      md.push("~~~text");
      md.push(payload.globalError);
      md.push("~~~");
      md.push("");
    }

    const mdPath = path.join(DIAGNOSTICS_DIR, "diagnostico-" + stamp + ".md");
    const jsonPath = path.join(DIAGNOSTICS_DIR, "diagnostico-" + stamp + ".json");
    fs.writeFileSync(mdPath, md.join("\n"), "utf8");
    fs.writeFileSync(jsonPath, JSON.stringify(payload, null, 2), "utf8");
    return { mdPath, jsonPath };
  }
};

const setupUsersGate = {
  chain: Promise.resolve()
};

const writeOpsGate = {
  chain: Promise.resolve()
};

const catalogImportGate = {
  chain: Promise.resolve()
};

const setupFlowGate = {
  chain: Promise.resolve()
};

async function runInGate(gate, task) {
  const previous = gate.chain;
  let release;
  gate.chain = new Promise((resolve) => {
    release = resolve;
  });
  await previous;
  try {
    return await task();
  } finally {
    release();
  }
}

class Monitor {
  constructor(sims) {
    this.start = performance.now();
    this.rows = new Map(
      sims.map((sim) => [
        sim.id,
        { id: sim.id, nombre: sim.nombre, estado: "pendiente", paso: "-", ok: 0, total: 0, error: "" }
      ])
    );
    this.ticker = null;
  }

  setTotal(id, total) {
    const row = this.rows.get(id);
    if (!row) return;
    row.total = total;
  }

  begin(id, paso) {
    const row = this.rows.get(id);
    if (!row) return;
    row.estado = "ejecutando";
    row.paso = paso;
    this.render();
  }

  done(id) {
    const row = this.rows.get(id);
    if (!row) return;
    row.ok += 1;
    this.render();
  }

  complete(id) {
    const row = this.rows.get(id);
    if (!row) return;
    row.estado = "ok";
    row.paso = "finalizado";
    this.render();
  }

  fail(id, error) {
    const row = this.rows.get(id);
    if (!row) return;
    row.estado = "error";
    row.error = String(error?.message || error || "fallo");
    this.render();
  }

  startLoop() {
    if (this.ticker) return;
    this.ticker = setInterval(() => this.render(), 1200);
    this.render();
  }

  stopLoop() {
    if (this.ticker) clearInterval(this.ticker);
    this.ticker = null;
    this.render();
  }

  render() {
    const elapsed = ((performance.now() - this.start) / 1000).toFixed(1);
    const lines = [];
    lines.push("=== POLARIA UI RUNNER (TERCERO) ===");
    lines.push(`Tiempo: ${elapsed}s`);
    lines.push("");
    lines.push("simulacion | estado     | progreso | paso actual                    | error");
    lines.push("-----------+------------+----------+-------------------------------+-------------------------");
    for (const row of this.rows.values()) {
      const progress = `${row.ok}/${row.total || "?"}`;
      lines.push(
        `${row.id.padEnd(10, " ")} | ${row.estado.padEnd(10, " ")} | ${progress.padEnd(8, " ")} | ${row.paso.slice(0, 29).padEnd(29, " ")} | ${row.error.slice(0, 23)}`
      );
    }
    process.stdout.write(`\x1Bc${lines.join("\n")}\n`);
  }
}

class RolePages {
  constructor(browser, frontUrl, timeoutMs) {
    this.browser = browser;
    this.frontUrl = frontUrl;
    this.timeoutMs = timeoutMs;
    this.map = new Map();
  }

  async get(key, email, password, codigoEmpresa = "") {
    if (this.map.has(key)) return this.map.get(key).page;
    const context = await this.browser.newContext({
      viewport: { width: 1536, height: 920 }
    });
    const page = await context.newPage();
    page.setDefaultTimeout(this.timeoutMs);
    await login(page, this.frontUrl, email, password, codigoEmpresa, this.timeoutMs);
    this.map.set(key, { context, page });
    return page;
  }

  async closeAll() {
    const entries = [...this.map.values()];
    for (const entry of entries) {
      await entry.context.close().catch(() => {});
    }
    this.map.clear();
  }
}

async function activeDialog(page, timeoutMs = 12_000) {
  const dialog = page.locator('[role="dialog"]').last();
  await dialog.waitFor({ state: "visible", timeout: timeoutMs });
  return dialog;
}

async function clickButtonByName(scope, regexes, timeoutMs = 20_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    for (const rx of regexes) {
      const btn = scope.getByRole("button", { name: rx }).first();
      if ((await btn.count()) === 0) continue;
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      const disabled = await btn.isDisabled().catch(() => false);
      if (disabled) continue;
      await btn.click();
      return true;
    }
    await sleep(220);
  }
  return false;
}

async function collectVisibleButtonLabels(page) {
  return page.evaluate(() => {
    const labels = [];
    const nodes = Array.from(document.querySelectorAll("button"));
    for (const node of nodes) {
      const text = (node.textContent || "").replace(/\s+/g, " ").trim();
      if (!text) continue;
      const style = window.getComputedStyle(node);
      if (style.visibility === "hidden" || style.display === "none") continue;
      labels.push(text);
    }
    return labels.slice(0, 20);
  });
}

async function clickActionButtonOrFail(page, buttonRegexes, timeoutMs, actionLabel) {
  const ok = await clickButtonByName(page, buttonRegexes, timeoutMs);
  if (ok) return;
  const labels = await collectVisibleButtonLabels(page);
  throw new Error(
    `No encontré botón de acción "${actionLabel}". URL=${page.url()} Botones visibles=${labels.join(" | ")}`
  );
}
async function openCreateModal(page) {
  const ok = await clickButtonByName(
    page,
    [
      /^Agregar$/i,
      /^Nuevo proveedor$/i,
      /^Nuevo cliente$/i,
      /^Nuevo comprador$/i,
      /^Nuevo cami[oó]n$/i,
      /^Nuevo producto$/i,
      /^Nueva solicitud$/i,
      /^Nueva orden$/i,
      /^Nueva venta$/i,
      /Nueva/i,
      /Nuevo/i,
      /Crear/i,
      /Agregar/i
    ],
    28_000
  );
  if (ok) return;
  const labels = await collectVisibleButtonLabels(page);
  throw new Error(
    `No encontré botón para abrir modal de creación. URL=${page.url()} Botones visibles=${labels.join(" | ")}`
  );
}

async function fillByLabel(scope, labelRx, value) {
  const locator = scope.getByLabel(labelRx).first();
  if ((await locator.count()) === 0) {
    throw new Error(`No encontré campo ${labelRx}`);
  }

  const valueStr = String(value);
  const tag = await locator
    .evaluate((el) => el.tagName.toLowerCase())
    .catch(() => "");

  // Caso normal: input/textarea
  if (tag === "input" || tag === "textarea") {
    await locator.fill(valueStr);
    return;
  }

  // Caso PhoneInput: el label puede resolver al select de país
  if (tag === "select") {
    const telCandidates = [
      scope.locator("input[type='tel']").first(),
      scope.locator("input[autocomplete='tel']").first(),
      scope.locator("input[id*='telefono']").first(),
      scope.locator("input[id*='phone']").first(),
    ];

    for (const candidate of telCandidates) {
      if ((await candidate.count()) === 0) continue;
      const visible = await candidate.isVisible().catch(() => false);
      if (!visible) continue;
      await candidate.fill(valueStr);
      return;
    }

    // Si realmente era un select normal, intentar selección por texto
    const selected = await selectOptionLike(locator, valueStr);
    if (selected) return;

    throw new Error(
      `El campo ${labelRx} resolvió a <select> y no encontré input telefónico asociado.`
    );
  }

  throw new Error(
    `El campo ${labelRx} no es editable con fill(). tag=${tag || "desconocido"}`
  );
}

async function selectOptionLike(selectLocator, textNeedle) {
  const options = selectLocator.locator("option");
  const count = await options.count();
  for (let i = 0; i < count; i += 1) {
    const option = options.nth(i);
    const txt = ((await option.textContent()) || "").toLowerCase();
    if (txt.includes(textNeedle.toLowerCase())) {
      const value = await option.getAttribute("value");
      if (value) {
        await selectLocator.selectOption(value);
        return true;
      }
    }
  }
  return false;
}


async function selectFromPicker(page, searchButtonRx, preferText = "", options = {}) {
  const { strictPrefer = false, triggerTimeoutMs = 22000, pickerTimeoutMs = 20000 } = options;
  const triggerRegexes = Array.isArray(searchButtonRx) ? searchButtonRx : [searchButtonRx];
  const exists = await clickButtonByName(page, triggerRegexes, triggerTimeoutMs);
  if (!exists) {
    const labels = await collectVisibleButtonLabels(page);
    throw new Error(
      "No encontré botón de picker " + formatRegexList(triggerRegexes) + ". URL=" + page.url() + " Botones visibles=" + labels.join(" | ")
    );
  }

  const picker = await activeDialog(page, 12000);

  const clickFirstVisible = async (locator) => {
    const count = await locator.count();
    for (let i = 0; i < count; i += 1) {
      const node = locator.nth(i);
      const visible = await node.isVisible().catch(() => false);
      if (!visible) continue;
      await node.click();
      return true;
    }
    return false;
  };

  if (preferText) {
    const searchInput = picker.getByRole("searchbox", { name: /Buscar/i }).first();
    if ((await searchInput.count()) > 0) {
      await searchInput.fill(String(preferText));
      await sleep(180);
    }

    const byRoleName = picker.getByRole("button", { name: rxContains(preferText) });
    if (await clickFirstVisible(byRoleName)) return;

    const byRowText = picker.locator('tr[role="button"]', {
      hasText: new RegExp(escapeRegExp(preferText), "i")
    });
    if (await clickFirstVisible(byRowText)) return;

    const rows = picker.locator('tr[role="button"]');
    const rowsCount = await rows.count();
    const needle = normalizeText(preferText);
    for (let i = 0; i < rowsCount; i += 1) {
      const row = rows.nth(i);
      const [aria, txt] = await Promise.all([
        row.getAttribute("aria-label").catch(() => ""),
        row.textContent().catch(() => "")
      ]);
      const haystack = normalizeText((aria || "") + " " + (txt || ""));
      if (haystack.includes(needle)) {
        await row.click();
        return;
      }
    }

    if (strictPrefer) {
      throw new Error('No encontré opción "' + preferText + '" en picker ' + formatRegexList(triggerRegexes));
    }
  }

  const startedAt = Date.now();
  while (Date.now() - startedAt < pickerTimeoutMs) {
    const rowCount = await picker.locator('tr[role="button"]').count();
    if (rowCount > 0) {
      await picker.locator('tr[role="button"]').first().click();
      return;
    }
    const cardCount = await picker.getByRole("button", { name: /Seleccionar/i }).count();
    if (cardCount > 0) {
      await picker.getByRole("button", { name: /Seleccionar/i }).first().click();
      return;
    }
    const empty = await picker
      .getByText(/sin resultados|sin datos|no hay registros|no hay datos/i)
      .first()
      .isVisible()
      .catch(() => false);
    if (empty) break;
    await sleep(260);
  }

  const preview = ((await picker.textContent().catch(() => "")) || "").replace(/\s+/g, " ").trim().slice(0, 220);
  throw new Error("No hay filas ni tarjetas seleccionables en el picker. Contexto=" + preview);
}

async function hasButtonByName(scope, regexes, timeoutMs = 3_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    for (const rx of regexes) {
      const btn = scope.getByRole("button", { name: rx }).first();
      if ((await btn.count()) === 0) continue;
      const visible = await btn.isVisible().catch(() => false);
      if (!visible) continue;
      const disabled = await btn.isDisabled().catch(() => false);
      if (disabled) continue;
      return true;
    }
    await sleep(220);
  }
  return false;
}

async function collectControlHints(scope) {
  return scope
    .evaluate(() => {
      const hints = [];
      const push = (kind, text) => {
        const value = String(text || "").replace(/\s+/g, " ").trim();
        if (value) hints.push(`${kind}:${value}`);
      };

      for (const el of document.querySelectorAll("label")) push("label", el.textContent);
      for (const el of document.querySelectorAll("select")) {
        push("select", el.getAttribute("aria-label") || el.name || el.id || "select");
      }
      for (const el of document.querySelectorAll("[role='combobox']")) {
        push("combobox", el.getAttribute("aria-label") || el.textContent);
      }
      for (const el of document.querySelectorAll("button")) {
        const style = window.getComputedStyle(el);
        if (style.display === "none" || style.visibility === "hidden") continue;
        push("button", el.textContent);
      }
      return hints.slice(0, 25);
    })
    .catch(() => []);
}

async function selectFirstVisibleOption(selectLocator) {
  const options = selectLocator.locator("option");
  const count = await options.count();
  for (let i = 0; i < count; i += 1) {
    const option = options.nth(i);
    const value = await option.getAttribute("value").catch(() => "");
    const text = ((await option.textContent()) || "").trim();
    if (!value && !text) continue;
    if (/seleccion|elige|choose/i.test(text)) continue;
    if (value) {
      await selectLocator.selectOption(value);
      return true;
    }
  }
  return false;
}

async function pickListboxOption(page, preferText = "") {
  const listbox = page.getByRole("listbox").last();
  if ((await listbox.count()) === 0) return false;

  if (preferText) {
    const byName = listbox.getByRole("option", { name: rxContains(preferText) }).first();
    if ((await byName.count()) > 0) {
      await byName.click();
      return true;
    }
    const count = await listbox.getByRole("option").count();
    const needle = normalizeText(preferText);
    for (let i = 0; i < count; i += 1) {
      const option = listbox.getByRole("option").nth(i);
      const text = normalizeText((await option.textContent().catch(() => "")) || "");
      if (text.includes(needle)) {
        await option.click();
        return true;
      }
    }
    return false;
  }

  const firstOption = listbox.getByRole("option").first();
  if ((await firstOption.count()) > 0) {
    await firstOption.click();
    return true;
  }
  return false;
}

async function selectComboboxByLabel(page, scope, labelRx, preferText = "") {
  const combobox = scope.getByRole("combobox", { name: labelRx }).first();
  if ((await combobox.count()) === 0) return false;
  await combobox.click();
  await sleep(220);
  return pickListboxOption(page, preferText);
}

async function selectFieldByControl(page, scope, labelRegexes, preferText = "", options = {}) {
  const { strictPrefer = false } = options;
  const labels = Array.isArray(labelRegexes) ? labelRegexes : [labelRegexes];

  for (const labelRx of labels) {
    const byLabel = scope.getByLabel(labelRx).first();
    if ((await byLabel.count()) > 0) {
      const tag = await byLabel.evaluate((el) => el.tagName.toLowerCase()).catch(() => "");
      if (tag === "select") {
        if (preferText) {
          const matched = await selectOptionLike(byLabel, preferText);
          if (matched) return true;
          if (strictPrefer) return false;
        }
        if (await selectFirstVisibleOption(byLabel)) return true;
        continue;
      }

      await byLabel.click().catch(() => {});
      await sleep(220);
      if (await pickListboxOption(page, preferText)) return true;
      if (!preferText && (await selectFirstVisibleOption(byLabel))) return true;
    }

    if (await selectComboboxByLabel(page, scope, labelRx, preferText)) return true;
  }

  return false;
}

async function selectFieldWithPickerFallback(page, scope, options = {}) {
  const {
    fieldLabel = "campo",
    pickerRegexes = [],
    labelRegexes = [],
    preferText = "",
    strictPrefer = false
  } = options;
  const triggerRegexes = Array.isArray(pickerRegexes) ? pickerRegexes : [pickerRegexes];
  const controlLabels = Array.isArray(labelRegexes) ? labelRegexes : [labelRegexes];

  if (await hasButtonByName(page, triggerRegexes, 4_000)) {
    return selectFromPicker(page, triggerRegexes, preferText, { strictPrefer });
  }

  const selected = await selectFieldByControl(page, scope, controlLabels, preferText, { strictPrefer });
  if (selected) return;

  const hints = await collectControlHints(scope);
  throw new Error(
    `No pude seleccionar ${fieldLabel}. Sin picker (${formatRegexList(triggerRegexes)}) ni control directo (${formatRegexList(controlLabels)}). Hints=${hints.join(" | ")}`
  );
}


async function submitModal(page, submitRx, allowDuplicate = true) {
  return runInGate(writeOpsGate, async () => {
    const dialog = await activeDialog(page, 15_000);
    const submit = dialog.getByRole("button", { name: submitRx }).first();
    if ((await submit.count()) === 0) throw new Error(`No encontré botón submit ${submitRx}`);
    const throttleRx = /(ThrottlerException|Too Many Requests|429|demasiadas solicitudes)/i;
    const maxAttempts = 10;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      const disabled = await submit.isDisabled().catch(() => false);
      if (disabled) {
        await sleep(300);
      } else {
        await submit.click();
      }

      await sleep(450);
      const dialogAlert = dialog.getByRole("alert").first();
      const pageAlert = page.getByRole("alert").first();
      const dialogAlertText =
        (await dialogAlert.count()) > 0 ? ((await dialogAlert.textContent()) || "").trim() : "";
      const pageAlertText =
        (await pageAlert.count()) > 0 ? ((await pageAlert.textContent()) || "").trim() : "";
      const alertTextNow = `${dialogAlertText} ${pageAlertText}`.trim();

      if (allowDuplicate && /(existe|duplic|ya est[aá]|ya fue)/i.test(alertTextNow)) {
        await clickButtonByName(dialog, [/Cerrar/i, /Cancelar/i]);
        await dialog.waitFor({ state: "hidden", timeout: 5_000 }).catch(() => {});
        return;
      }

      if (throttleRx.test(alertTextNow) && attempt < maxAttempts) {
        const backoffMs = Math.min(30_000, 1_500 * 2 ** (attempt - 1)) + Math.floor(Math.random() * 1400);
        await sleep(backoffMs);
        continue;
      }

      try {
        await dialog.waitFor({ state: "hidden", timeout: 16_000 });
        return;
      } catch {
        const dialogAlert2 = dialog.getByRole("alert").first();
        const pageAlert2 = page.getByRole("alert").first();
        const dialogAlertText2 =
          (await dialogAlert2.count()) > 0 ? ((await dialogAlert2.textContent()) || "").trim() : "";
        const pageAlertText2 =
          (await pageAlert2.count()) > 0 ? ((await pageAlert2.textContent()) || "").trim() : "";
        const alertText = `${dialogAlertText2} ${pageAlertText2}`.trim();

        if (allowDuplicate && /(existe|duplic|ya est[aá]|ya fue)/i.test(alertText)) {
          await clickButtonByName(dialog, [/Cerrar/i, /Cancelar/i]);
          await dialog.waitFor({ state: "hidden", timeout: 5_000 }).catch(() => {});
          return;
        }

        if (throttleRx.test(alertText) && attempt < maxAttempts) {
          const backoffMs = Math.min(30_000, 1_500 * 2 ** (attempt - 1)) + Math.floor(Math.random() * 1400);
          await sleep(backoffMs);
          continue;
        }

        throw new Error(
          alertText || `El modal no cerró después de enviar (intento ${attempt}/${maxAttempts}).`
        );
      }
    }

    throw new Error(`No se pudo enviar el modal tras ${maxAttempts} intentos.`);
  });
}

async function login(page, frontUrl, email, password, codigoEmpresa = "", timeoutMs = 30_000) {
  const maxAttempts = 3;
  const redirectTimeout = Math.max(timeoutMs, 90_000);

  async function fillReactInput(locator, value) {
    await locator.waitFor({ state: "visible", timeout: timeoutMs });
    await locator.click({ clickCount: 3 }).catch(() => {});
    await locator.fill("").catch(() => {});
    await locator.type(String(value), { delay: 20 });

    // Fuerza evento input/change para React
    await locator.evaluate((el, v) => {
      if ("value" in el) {
        el.value = v;
      }
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
      el.dispatchEvent(new Event("blur", { bubbles: true }));
    }, String(value));
  }

  async function readAlertText() {
    const alert = page.getByRole("alert").first();
    if ((await alert.count()) === 0) return "";
    return ((await alert.textContent()) || "").trim();
  }

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      await page.goto(`${frontUrl}/login`, { waitUntil: "domcontentloaded" });

      const currentPath = getPathnameFromUrl(page.url());
      if (currentPath.startsWith("/configurador") || currentPath.startsWith("/dashboard")) {
        return;
      }

      const emailInput = page.locator("#identificador").first();
      await fillReactInput(emailInput, email);

      // Intento submit por Enter primero
      await emailInput.press("Enter").catch(() => {});
      await page.waitForTimeout(350);

      const companySelect = page.locator("#codigoEmpresa").first();
      const continueBtn = page.getByRole("button", { name: /Continuar/i }).first();

      // Si sigue en login y no avanzó, click en Continuar
      if (getPathnameFromUrl(page.url()).startsWith("/login")) {
        const hasContinue = (await continueBtn.count()) > 0;
        if (hasContinue) {
          const disabled = await continueBtn.isDisabled().catch(() => false);
          if (!disabled) {
            await continueBtn.click();
          }
        }
      }

      const companyVisible = await companySelect.isVisible().catch(() => false);
      if (companyVisible) {
        if (codigoEmpresa) {
          await companySelect.selectOption({ value: codigoEmpresa }).catch(async () => {
            await companySelect.selectOption({ label: codigoEmpresa });
          });
        } else {
          // Si no vino código, solo autoselecciona si hay una única opción válida
          const options = companySelect.locator("option");
          const count = await options.count();
          const validValues = [];
          for (let i = 0; i < count; i += 1) {
            const val = (await options.nth(i).getAttribute("value")) || "";
            if (val.trim()) validValues.push(val);
          }
          if (validValues.length === 1) {
            await companySelect.selectOption({ value: validValues[0] });
          } else {
            throw new Error("Login requiere selección de empresa y no se pudo resolver automáticamente.");
          }
        }

        await page.waitForTimeout(250);
        const disabled = await continueBtn.isDisabled().catch(() => false);
        if (!disabled) {
          await continueBtn.click();
        } else {
          await companySelect.press("Enter").catch(() => {});
        }
      }

      const passwordInput = page.locator("#password").first();
      await fillReactInput(passwordInput, password);

      // Submit por Enter + fallback click
      await passwordInput.press("Enter").catch(() => {});
      await page.waitForTimeout(300);

      if (getPathnameFromUrl(page.url()).startsWith("/login")) {
        const loginBtn = page.getByRole("button", { name: /Iniciar sesi[oó]n/i }).first();
        if ((await loginBtn.count()) > 0) {
          const disabled = await loginBtn.isDisabled().catch(() => false);
          if (!disabled) await loginBtn.click();
        }
      }

      await page.waitForURL(/\/(dashboard|configurador)/, { timeout: redirectTimeout });
      return;
    } catch (error) {
      if (attempt === maxAttempts) {
        const alertText = await readAlertText();
        const url = page.url();
        throw new Error(
          `Falló login (${email}) tras ${maxAttempts} intentos. URL=${url}${alertText ? ` | alert=${alertText}` : ""} | causa=${error?.message || error}`
        );
      }

      // Limpieza/reintento
      await page.context().clearCookies().catch(() => {});
      await page.goto("about:blank").catch(() => {});
      await page.waitForTimeout(1200 * attempt);
    }
  }
}

function getPathnameFromUrl(url) {
  try {
    return new URL(url).pathname;
  } catch {
    return url;
  }
}

async function waitForPathname(page, expectedRoute, timeoutMs = 30_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const pathname = getPathnameFromUrl(page.url());
    if (pathname.startsWith(expectedRoute)) return;

    if (expectedRoute.startsWith("/configurador") && pathname.startsWith("/dashboard")) {
      throw new Error(
        `Se esperaba ruta de configurador (${expectedRoute}), pero se llegó a ${pathname}.`
      );
    }

    if (pathname.startsWith("/login")) {
      throw new Error("Sesión no válida o expirada: redirigió a /login.");
    }

    await sleep(180);
  }

  throw new Error(
    `No se alcanzó la ruta esperada ${expectedRoute}. Ruta actual: ${getPathnameFromUrl(page.url())}`
  );
}

async function go(page, frontUrl, route) {
  await page.goto(`${frontUrl}${route}`, { waitUntil: "domcontentloaded" });
  await waitForPathname(page, route, 30_000);
}
async function setupCore(configPage, sim, frontUrl) {
  await go(configPage, frontUrl, "/configurador/creacion/empresas");
  await openCreateModal(configPage);
  {
    const dialog = await activeDialog(configPage);
    await fillByLabel(dialog, /Raz[oó]n social/i, sim.setup.empresa.razonSocial);
    await fillByLabel(dialog, /Tel[eé]fono/i, sim.setup.empresa.telefono);
  }
  await submitModal(configPage, /Crear/i, true);
    await sleep(8500 + Math.floor(Math.random() * 2500));

  await go(configPage, frontUrl, "/configurador/creacion/cuentas");
  await openCreateModal(configPage);
  {
    const dialog = await activeDialog(configPage);
    await selectFromPicker(configPage, /Buscar Empresa/i, sim.setup.empresa.razonSocial);
    await fillByLabel(dialog, /^Nombre$/i, sim.setup.cuenta.nombre);
  }
  await submitModal(configPage, /Crear/i, true);
    await sleep(8500 + Math.floor(Math.random() * 2500));

  await go(configPage, frontUrl, "/configurador/creacion/bodega-interna");
  await openCreateModal(configPage);
  {
    const dialog = await activeDialog(configPage);
    await selectFromPicker(configPage, /Buscar Cuenta destino/i, sim.setup.cuenta.nombre);
    await fillByLabel(dialog, /^Nombre$/i, sim.setup.bodega.nombre);
    await fillByLabel(dialog, /Capacidad/i, sim.setup.bodega.capacidad);
  }
  await submitModal(configPage, /Crear/i, true);
    await sleep(8500 + Math.floor(Math.random() * 2500));
}

async function setupUsers(configPage, sim, frontUrl) {
  const users = [
    { key: "adminCuenta", role: "administrador_cuenta", roleLabel: "Administrador de cuenta", assign: "cuenta" },
    { key: "operadorCuenta", role: "operador_cuenta", roleLabel: "Operador de cuenta", assign: "cuenta" },
    { key: "adminBodega", role: "administrador_bodega", roleLabel: "Administrador de bodega", assign: "bodega" },
    { key: "jefeBodega", role: "jefe_bodega", roleLabel: "Jefe de bodega", assign: "bodega" },
    { key: "custodio", role: "custodio", roleLabel: "Custodio", assign: "bodega" },
    { key: "operario", role: "operario", roleLabel: "Operario", assign: "bodega" },
    { key: "procesador", role: "procesador", roleLabel: "Procesador", assign: "bodega" },
    { key: "transportista", role: "transportista", roleLabel: "Transportista", assign: "bodega" }
  ];

  await go(configPage, frontUrl, "/configurador/asignacion/usuarios");

  for (const item of users) {
    const u = sim.usuarios[item.key];
    await openCreateModal(configPage);
    const dialog = await activeDialog(configPage);

    await fillByLabel(dialog, /^Nombre$/i, u.nombre);

    // Selección estricta de rol: nunca fallback a primera fila
    await selectFromPicker(configPage, /Buscar Rol/i, item.roleLabel, { strictPrefer: true });

    // Validación defensiva del rol realmente seleccionado
    const rolSeleccionado = await dialog.locator("#usuario-rol").inputValue().catch(() => "");
    const roleOk =
      normalizeText(rolSeleccionado).includes(normalizeText(item.role)) ||
      normalizeText(rolSeleccionado).includes(normalizeText(item.roleLabel));

    if (!roleOk) {
      throw new Error(
        `Rol incorrecto para ${u.email}. Esperado=${item.role} (${item.roleLabel}) | seleccionado="${rolSeleccionado}"`
      );
    }

    const asignado = dialog.locator("#usuario-asignado");
    if ((await asignado.count()) > 0) {
      const needle = item.assign === "cuenta" ? sim.setup.cuenta.nombre : sim.setup.bodega.nombre;
      const selected = await selectOptionLike(asignado, needle);
      if (!selected) {
        throw new Error(`No pude seleccionar "${needle}" en Asignado para ${u.email}`);
      }
    }

    await fillByLabel(dialog, /Correo/i, u.email);
    await fillByLabel(dialog, /^Clave$/i, sim.passwordUnica);
    await submitModal(configPage, /Crear/i, true);
    await sleep(8500 + Math.floor(Math.random() * 2500));
  }
}

async function setupMastersAndCatalog(adminPage, sim, frontUrl, excelPath) {
  // Proveedor
  await go(adminPage, frontUrl, "/dashboard/administracion/asignacion-creacion/proveedores");
  await openCreateModal(adminPage);
  {
    const d = await activeDialog(adminPage);
    await fillByLabel(d, /^Proveedor$/i, sim.maestros.proveedor.nombre);
    await fillByLabel(d, /^Nombre$/i, sim.maestros.proveedor.contacto);
    await fillByLabel(d, /Tel[eé]fono/i, sim.maestros.proveedor.telefono);
    await fillByLabel(d, /Email/i, sim.maestros.proveedor.email);
  }
  await submitModal(adminPage, /Crear/i, true);

  // Cliente
  await go(adminPage, frontUrl, "/dashboard/administracion/asignacion-creacion/clientes");
  await openCreateModal(adminPage);
  {
    const d = await activeDialog(adminPage);
    await fillByLabel(d, /^Nombre$/i, sim.maestros.cliente.nombre);
    await fillByLabel(d, /^NIT$/i, sim.maestros.cliente.nit);
    await fillByLabel(d, /Tel[eé]fono/i, sim.maestros.cliente.telefono);
  }
  await submitModal(adminPage, /Crear/i, true);

  // Comprador
  await go(adminPage, frontUrl, "/dashboard/administracion/asignacion-creacion/compradores");
  await openCreateModal(adminPage);
  {
    const d = await activeDialog(adminPage);
    await fillByLabel(d, /Nombre del comprador/i, sim.maestros.comprador.nombre);
    await fillByLabel(d, /Tel[eé]fono/i, sim.maestros.comprador.telefono);
  }
  await submitModal(adminPage, /Crear/i, true);

  // Camión
  await go(adminPage, frontUrl, "/dashboard/administracion/asignacion-creacion/camiones");
  await openCreateModal(adminPage);
  {
    const d = await activeDialog(adminPage);
    await fillByLabel(d, /Placa/i, sim.maestros.camion.placa);
    await selectFromPicker(adminPage, /Buscar Marca/i, "");
    await selectFromPicker(adminPage, /Buscar Modelo/i, "");
    await fillByLabel(d, /Peso m[aá]x/i, sim.maestros.camion.capacidadKg);
    await fillByLabel(d, /Volumen/i, sim.maestros.camion.capacidadM3);
    await fillByLabel(d, /Cap\. pallets/i, sim.maestros.camion.capacidadPallets);
  }
  await submitModal(adminPage, /Crear/i, true);

  // Importar catálogo Excel
  await go(adminPage, frontUrl, "/dashboard/administracion/catalogo");
  await adminPage.getByRole("button", { name: /Importar Excel/i }).click();
  const fileInput = adminPage.locator('input[type="file"]').first();
  await fileInput.setInputFiles(excelPath);
  await sleep(2500);
}

async function createAndEmitPurchaseOrder(operadorPage, adminPage, sim, frontUrl, compraKg) {
  // Crear SOL
  await go(operadorPage, frontUrl, "/dashboard/compras");
  await operadorPage.getByRole("button", { name: /Nueva solicitud/i }).click();
  {
    const d = await activeDialog(operadorPage);
    await selectFromPicker(operadorPage, /Buscar Proveedor/i, sim.maestros.proveedor.nombre);
    await selectFromPicker(operadorPage, /Buscar Producto/i, "");
    await fillByLabel(d, /Peso \(kg\)/i, compraKg);
    await d.getByRole("button", { name: /^Agregar$/i }).click();
  }
  await submitModal(operadorPage, /Guardar solicitud/i, false);

  // Enviar aprobación
  await go(operadorPage, frontUrl, "/dashboard/compras");
  await operadorPage.locator("table tbody tr").first().click();
  {
    const d = await activeDialog(operadorPage);
    await d.getByRole("button", { name: /Enviar aprobaci[oó]n/i }).click();
    await d.waitFor({ state: "hidden", timeout: 15_000 });
  }

  // Aprobar
  await go(adminPage, frontUrl, "/dashboard/compras");
  await adminPage.locator("table tbody tr").first().click();
  {
    const d = await activeDialog(adminPage);
    await d.getByRole("button", { name: /^Aprobar$/i }).click();
    await d.waitFor({ state: "hidden", timeout: 15_000 });
  }

  // Convertir
  await adminPage.locator("table tbody tr").first().click();
  {
    const d = await activeDialog(adminPage);
    await d.getByRole("button", { name: /Convertir a OC/i }).click();
    await d.waitFor({ state: "hidden", timeout: 15_000 });
  }

  // Emitir OC con destino
  await adminPage.getByRole("button", { name: /[ÓO]rdenes/i }).click();
  await adminPage.locator("table tbody tr").first().click();
  {
    const d = await activeDialog(adminPage);
    const date = d.locator("#orden-fecha-entrega");
    if ((await date.count()) > 0) await date.fill(todayInput());
    const tipo = d.locator("#orden-destino-tipo");
    if ((await tipo.count()) > 0) await selectOptionLike(tipo, "interna");
    const bodega = d.locator("#orden-destino-bodega");
    if ((await bodega.count()) > 0) {
      await selectOptionLike(bodega, sim.setup.bodega.nombre);
    }
    await sleep(900);
    await d.getByRole("button", { name: /Emitir orden/i }).click();
    await d.waitFor({ state: "hidden", timeout: 20_000 });
  }
}

async function receiveOrder(custodioPage, sim, frontUrl, tempC, pesoKg) {
  await go(custodioPage, frontUrl, "/dashboard/custodio/ingreso");
  await selectFromPicker(custodioPage, /Buscar Orden de compra/i, "");
  await sleep(800);

  const tempFields = custodioPage.getByLabel(/Temperatura \([°º]C\)/i);
  const tempCount = await tempFields.count();
  for (let i = 0; i < tempCount; i += 1) {
    await tempFields.nth(i).fill(String(tempC));
  }
  const pesoFields = custodioPage.getByLabel(/Peso recibido \(kg\)/i);
  const pesoCount = await pesoFields.count();
  for (let i = 0; i < pesoCount; i += 1) {
    await pesoFields.nth(i).fill(String(pesoKg));
  }
  await custodioPage.getByRole("button", { name: /Registrar ingreso/i }).first().click();
  await sleep(2400);
}

async function createIngresoAndCompleteTask(jefePage, operarioPage, sim, frontUrl) {
  await go(jefePage, frontUrl, "/dashboard/jefe-bodega/estado-bodega");
  await clickActionButtonOrFail(
    jefePage,
    [/^Ingresos$/i, /Ingresos/i, /Registrar entrada/i],
    30_000,
    "Ingresos"
  );
  await activeDialog(jefePage);
  await selectFromPicker(jefePage, /Buscar Producto en ingreso/i, "");
  await selectFromPicker(jefePage, /Buscar Posici[oó]n en bodega/i, "");
  await selectFromPicker(jefePage, /Buscar Operario/i, sim.usuarios.operario.nombre);
  await submitModal(jefePage, /Crear ingreso/i, false);

  await go(operarioPage, frontUrl, "/dashboard/operario/operacion");
  for (let i = 0; i < 8; i += 1) {
    const task = operarioPage.locator('button:has-text("ID de tarea")').first();
    if ((await task.count()) === 0) break;
    await task.click();
    await sleep(1800);
  }
}

async function createAndEmitSale(operadorPage, sim, frontUrl, ventaKg) {
  await go(operadorPage, frontUrl, "/dashboard/ventas/ordenes");
  await operadorPage.getByRole("button", { name: /Nueva venta/i }).click();
  {
    const d = await activeDialog(operadorPage);
    await selectFieldWithPickerFallback(operadorPage, d, {
      fieldLabel: "comprador",
      pickerRegexes: [/Buscar Comprador/i, /Seleccionar Comprador/i, /Buscar Cliente/i, /Comprador/i],
      labelRegexes: [/Comprador/i, /Cliente/i],
      preferText: sim.maestros.comprador.nombre
    });
    await selectFieldWithPickerFallback(operadorPage, d, {
      fieldLabel: "producto",
      pickerRegexes: [/Buscar Producto/i, /Seleccionar Producto/i, /Producto/i],
      labelRegexes: [/Producto/i],
      preferText: ""
    });
    await fillByLabel(d, /Cantidad \(kg\)/i, ventaKg);
    await d.getByRole("button", { name: /^Agregar$/i }).click();
    await fillByLabel(d, /Observaciones/i, `Venta auto ${sim.id}`);
  }
  await submitModal(operadorPage, /Crear venta/i, false);

  await operadorPage.locator("table tbody tr").first().click();
  {
    const d = await activeDialog(operadorPage);
    const emitir = d.getByRole("button", { name: /Emitir venta/i });
    if ((await emitir.count()) > 0) {
      await emitir.first().click();
      await sleep(2000);
    }
  }
}

async function createSalidaAndCompleteTask(jefePage, operarioPage, sim, frontUrl) {
  await go(jefePage, frontUrl, "/dashboard/jefe-bodega/estado-bodega");
  await clickActionButtonOrFail(
    jefePage,
    [/^Crear Salida$/i, /Crear Salida/i, /Registrar salida/i],
    30_000,
    "Crear Salida"
  );
  await activeDialog(jefePage);
  await selectFromPicker(jefePage, /Buscar Orden de venta/i, "");
  await selectFromPicker(jefePage, /Buscar Destino/i, "");
  await selectFromPicker(jefePage, /Buscar Operario/i, sim.usuarios.operario.nombre);
  await submitModal(jefePage, /Crear salida/i, false);

  await go(operarioPage, frontUrl, "/dashboard/operario/operacion");
  for (let i = 0; i < 8; i += 1) {
    const task = operarioPage.locator('button:has-text("ID de tarea")').first();
    if ((await task.count()) === 0) break;
    await task.click();
    await sleep(1800);
  }
}

async function createPaquete(custodioPage, sim, frontUrl) {
  await go(custodioPage, frontUrl, "/dashboard/custodio/ingreso");
  await selectFromPicker(custodioPage, /Buscar Ventas para el paquete/i, "");
  await custodioPage.getByRole("button", { name: /Armar paquete de despacho/i }).click();
  await sleep(900);
  await selectFromPicker(custodioPage, /Buscar Cami[oó]n asignado/i, sim.maestros.camion.placa);
  await custodioPage.getByRole("button", { name: /Enviar paquete al transporte/i }).click();
  await sleep(2400);
}

async function drawSignature(modal, page) {
  const canvas = modal.locator("canvas").first();
  const box = await canvas.boundingBox();
  if (!box) throw new Error("No se encontró canvas de firma.");
  await page.mouse.move(box.x + 24, box.y + 30);
  await page.mouse.down();
  await page.mouse.move(box.x + 120, box.y + 72);
  await page.mouse.move(box.x + 226, box.y + 54);
  await page.mouse.up();
}

async function registerDelivery(transportPage, sim, frontUrl, evidencePath, conforme, motivoNoConforme = "") {
  await go(transportPage, frontUrl, "/dashboard/transporte");
  await transportPage.getByRole("button", { name: /Realizar entrega/i }).first().click();
  const modal = await activeDialog(transportPage, 15_000);

  const checks = modal.locator('input[type="checkbox"]');
  const checksCount = await checks.count();
  for (let i = 0; i < checksCount; i += 1) {
    await checks.nth(i).check();
  }
  await modal.getByRole("button", { name: /Siguiente/i }).click();

  await modal.locator('input[type="file"]').setInputFiles(evidencePath);
  await modal.getByRole("button", { name: /Siguiente/i }).click();

  await drawSignature(modal, transportPage);
  await modal.getByRole("button", { name: /Siguiente/i }).click();

  if (conforme) {
    await modal.getByRole("button", { name: /S[ií], conforme/i }).click();
  } else {
    await modal.getByRole("button", { name: /No conforme/i }).click();
    await modal
      .getByPlaceholder(/Describ[ií] la incidencia/i)
      .fill(motivoNoConforme || "Cadena de frio interrumpida en ruta");
  }

  await modal.getByRole("button", { name: /Cerrar entrega/i }).click();
  await modal.waitFor({ state: "hidden", timeout: 60_000 });
}

async function runCycleA(sim, pages, frontUrl, evidencePath) {
  await createAndEmitPurchaseOrder(
    await pages.get("operador", sim.usuarios.operadorCuenta.email, sim.passwordUnica),
    await pages.get("admin", sim.usuarios.adminCuenta.email, sim.passwordUnica),
    sim,
    frontUrl,
    sim.flujo.compraKg
  );
  await receiveOrder(
    await pages.get("custodio", sim.usuarios.custodio.email, sim.passwordUnica),
    sim,
    frontUrl,
    sim.flujo.recepcionTempC,
    sim.flujo.recepcionKgA
  );
  await createIngresoAndCompleteTask(
    await pages.get("jefe", sim.usuarios.jefeBodega.email, sim.passwordUnica),
    await pages.get("operario", sim.usuarios.operario.email, sim.passwordUnica),
    sim,
    frontUrl
  );
  await createAndEmitSale(
    await pages.get("operador", sim.usuarios.operadorCuenta.email, sim.passwordUnica),
    sim,
    frontUrl,
    sim.flujo.ventaKg
  );
  await createSalidaAndCompleteTask(
    await pages.get("jefe", sim.usuarios.jefeBodega.email, sim.passwordUnica),
    await pages.get("operario", sim.usuarios.operario.email, sim.passwordUnica),
    sim,
    frontUrl
  );
  await createPaquete(
    await pages.get("custodio", sim.usuarios.custodio.email, sim.passwordUnica),
    sim,
    frontUrl
  );
  await registerDelivery(
    await pages.get("transportista", sim.usuarios.transportista.email, sim.passwordUnica),
    sim,
    frontUrl,
    evidencePath,
    true
  );
}

async function runCycleB(sim, pages, frontUrl, evidencePath) {
  if (sim.flujo.cicloB === "recepcion_diferencia") {
    await createAndEmitPurchaseOrder(
      await pages.get("operador", sim.usuarios.operadorCuenta.email, sim.passwordUnica),
      await pages.get("admin", sim.usuarios.adminCuenta.email, sim.passwordUnica),
      sim,
      frontUrl,
      sim.flujo.compraKg
    );
    await receiveOrder(
      await pages.get("custodio", sim.usuarios.custodio.email, sim.passwordUnica),
      sim,
      frontUrl,
      sim.flujo.recepcionTempC,
      sim.flujo.recepcionKgB
    );
    return;
  }

  if (sim.flujo.cicloB === "entrega_no_conforme") {
    await createAndEmitPurchaseOrder(
      await pages.get("operador", sim.usuarios.operadorCuenta.email, sim.passwordUnica),
      await pages.get("admin", sim.usuarios.adminCuenta.email, sim.passwordUnica),
      sim,
      frontUrl,
      sim.flujo.compraKg
    );
    await receiveOrder(
      await pages.get("custodio", sim.usuarios.custodio.email, sim.passwordUnica),
      sim,
      frontUrl,
      sim.flujo.recepcionTempC,
      sim.flujo.recepcionKgB
    );
    await createIngresoAndCompleteTask(
      await pages.get("jefe", sim.usuarios.jefeBodega.email, sim.passwordUnica),
      await pages.get("operario", sim.usuarios.operario.email, sim.passwordUnica),
      sim,
      frontUrl
    );
    await createAndEmitSale(
      await pages.get("operador", sim.usuarios.operadorCuenta.email, sim.passwordUnica),
      sim,
      frontUrl,
      sim.flujo.ventaKg
    );
    await createSalidaAndCompleteTask(
      await pages.get("jefe", sim.usuarios.jefeBodega.email, sim.passwordUnica),
      await pages.get("operario", sim.usuarios.operario.email, sim.passwordUnica),
      sim,
      frontUrl
    );
    await createPaquete(
      await pages.get("custodio", sim.usuarios.custodio.email, sim.passwordUnica),
      sim,
      frontUrl
    );
    await registerDelivery(
      await pages.get("transportista", sim.usuarios.transportista.email, sim.passwordUnica),
      sim,
      frontUrl,
      evidencePath,
      false,
      sim.flujo.motivoNoConforme
    );
  }
}

async function runOneSimulation(sim, opts) {
  const { browserType, executablePath, monitor, frontUrl, catalogPath, timeoutMs, headed, configurador, evidencePath } = opts;
  const steps = [
    "login-configurador",
    "setup-core",
    "setup-usuarios",
    "setup-admin-catalogo",
    "ciclo-a",
    "ciclo-b"
  ];
  monitor.setTotal(sim.id, steps.length);
  RUN_DIAG.setTotal(sim.id, steps.length);

  const browser = await chromium.launch({
    headless: !headed,
    executablePath,
    slowMo: SLOW_MO_MS
  });

  const pages = new RolePages(browser, frontUrl, timeoutMs);
  const runStep = async (label, fn) => {
    const stepName = browserType + ": " + label;
    monitor.begin(sim.id, stepName);
    RUN_DIAG.stepStart(sim.id, stepName);
    try {
      await fn();
      monitor.done(sim.id);
      RUN_DIAG.stepOk(sim.id, stepName);
    } catch (error) {
      RUN_DIAG.stepError(sim.id, stepName, error);
      throw error;
    }
  };

  try {
    await runInGate(setupFlowGate, async () => {
      await runStep("login-configurador", async () => {
        const p = await pages.get(
          "config",
          configurador.email,
          configurador.password,
          configurador.codigoEmpresa || ""
        );
        return p;
      });
      await sleep(2200);

      await runStep("setup-core", async () => {
        const p = await pages.get(
          "config",
          configurador.email,
          configurador.password,
          configurador.codigoEmpresa || ""
        );
        await setupCore(p, sim, frontUrl);
      });
      await sleep(2400);

      await runStep("setup-usuarios", async () => {
        await runInGate(setupUsersGate, async () => {
          const p = await pages.get(
            "config",
            configurador.email,
            configurador.password,
            configurador.codigoEmpresa || ""
          );
          await setupUsers(p, sim, frontUrl);
        });
      });
      await sleep(2500);

      await runStep("setup-admin-catalogo", async () => {
        const admin = await pages.get("admin", sim.usuarios.adminCuenta.email, sim.passwordUnica);
        await setupMastersAndCatalog(admin, sim, frontUrl, catalogPath);
      });
    });

    await runStep("ciclo-a", async () => {
      await runCycleA(sim, pages, frontUrl, evidencePath);
    });
    await runStep("ciclo-b", async () => {
      await runCycleB(sim, pages, frontUrl, evidencePath);
    });
    monitor.complete(sim.id);
    RUN_DIAG.completeSimulation(sim.id);
  } catch (error) {
    monitor.fail(sim.id, error);
    RUN_DIAG.failSimulation(sim.id, error);
    throw error;
  } finally {
    await pages.closeAll().catch(() => {});
    await browser.close().catch(() => {});
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const scenario = readScenario();

  const frontUrl = args.frontUrl || scenario.frontUrl || "http://localhost:3001";
  const defaultDownloads = process.env.USERPROFILE
    ? path.join(process.env.USERPROFILE, "Downloads")
    : "";
  const downloadsDir = args.downloadsDir || defaultDownloads;

  if (!downloadsDir) {
    throw new Error("Debes indicar --downloads con la carpeta de tus Excel.");
  }

  ensureFile(normalizePathFromFileUrl(downloadsDir), "carpeta downloads");
  const evidencePath = ensureEvidenceImage();

  const monitor = new Monitor(scenario.simulaciones);
  RUN_DIAG.setMeta({
    frontUrl,
    downloadsDir,
    headed: args.headed,
    timeoutMs: args.timeoutMs,
    scenarioPath: SCENARIO_PATH
  });
  monitor.startLoop();

  try {
    const jobs = scenario.simulaciones.map(async (sim) => {
      const browserType = scenario.browserMap[sim.id];
      const executablePath = DEFAULT_BROWSER_PATHS[browserType];
      if (!executablePath) {
        throw new Error(`No hay ruta de navegador para ${browserType} (${sim.id}).`);
      }
      ensureFile(executablePath, `navegador ${browserType}`);

      const catalogPath = path.join(downloadsDir, sim.catalogoExcel);
      ensureFile(catalogPath, `Excel de ${sim.id}`);
      const meta = parseCatalogMeta(catalogPath);
      console.log(`[${sim.id}] Excel OK -> hoja: ${meta.sheet}, filas: ${meta.rows}`);
      RUN_DIAG.registerSimulation(sim, browserType);

      return runOneSimulation(sim, {
        browserType,
        executablePath,
        monitor,
        frontUrl,
        catalogPath,
        timeoutMs: args.timeoutMs,
        headed: args.headed,
        configurador: scenario.configurador,
        evidencePath
      });
    });

    const results = await Promise.allSettled(jobs);
    monitor.stopLoop();
    const reportFiles = RUN_DIAG.writeFinal(results, null);
    if (reportFiles) {
      console.log("[diagnostico] Markdown: " + reportFiles.mdPath);
      console.log("[diagnostico] JSON: " + reportFiles.jsonPath);
    }
    const failures = results.filter((r) => r.status === "rejected");
    if (failures.length > 0) {
      const first = failures[0];
      const reason = first.status === "rejected" ? first.reason : "Error desconocido";
      console.error(
        `\n❌ Simulaciones con error: ${failures.length}/${results.length}. Primer error:`,
        reason?.stack || reason?.message || reason
      );
      process.exit(1);
    }
    console.log("\n✅ Simulaciones finalizadas.");
  } catch (error) {
    monitor.stopLoop();
    const reportFiles = RUN_DIAG.writeFinal([], error);
    if (reportFiles) {
      console.log("[diagnostico] Markdown: " + reportFiles.mdPath);
      console.log("[diagnostico] JSON: " + reportFiles.jsonPath);
    }
    console.error("\n❌ Error global:", error?.stack || error?.message || error);
    process.exit(1);
  }
}

main();








