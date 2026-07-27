$ErrorActionPreference = "Stop"

$RunnerPath = "C:\polaria-ui-runner\run.mjs"
if (-not (Test-Path $RunnerPath)) {
  throw "No existe: $RunnerPath"
}

# Backup
$backup = "$RunnerPath.bak-" + (Get-Date -Format "yyyyMMdd-HHmmss")
Copy-Item $RunnerPath $backup -Force
Write-Host "Backup creado en: $backup"

$patcher = Join-Path $env:TEMP "patch-polaria-runner-v4.mjs"

@'
import fs from "fs";
import path from "path";

const runnerPath = process.argv[2];
if (!runnerPath) throw new Error("Falta ruta run.mjs");
if (!fs.existsSync(runnerPath)) throw new Error("No existe run.mjs: " + runnerPath);

let code = fs.readFileSync(runnerPath, "utf8");

function replaceOrThrow(label, regex, replacement) {
  if (!regex.test(code)) throw new Error("No encontré bloque para: " + label);
  code = code.replace(regex, replacement);
}

function replaceIfMissing(marker, label, regex, replacement) {
  if (code.includes(marker)) return;
  replaceOrThrow(label, regex, replacement);
}

function insertBeforeOrThrow(label, anchor, block) {
  const idx = code.indexOf(anchor);
  if (idx < 0) throw new Error("No encontré ancla para: " + label);
  code = code.slice(0, idx) + block + "\n" + code.slice(idx);
}

// 1) Constante diagnósticos
if (!code.includes('const DIAGNOSTICS_DIR = path.join(ROOT, "diagnosticos");')) {
  replaceOrThrow(
    "const EVIDENCE_PATH",
    /const EVIDENCE_PATH = path\.join\(ROOT, "evidencia-entrega\.png"\);/,
    'const EVIDENCE_PATH = path.join(ROOT, "evidencia-entrega.png");\nconst DIAGNOSTICS_DIR = path.join(ROOT, "diagnosticos");'
  );
}

// 2) Bloque diagnóstico (insertado antes de setupUsersGate)
if (!code.includes("const RUN_DIAG = {")) {
  const diagBlock = `
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
      if (s.error) md.push("- Error final: " + s.error.replace(/\\n/g, " "));
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
    fs.writeFileSync(mdPath, md.join("\\n"), "utf8");
    fs.writeFileSync(jsonPath, JSON.stringify(payload, null, 2), "utf8");
    return { mdPath, jsonPath };
  }
};
`;
  insertBeforeOrThrow("bloque diagnóstico", "const setupUsersGate = {", diagBlock);
}

// 3) selectFromPicker robusto
if (!code.includes("triggerTimeoutMs = 22000")) {
  const selectFn = `
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

  const preview = ((await picker.textContent().catch(() => "")) || "").replace(/\\s+/g, " ").trim().slice(0, 220);
  throw new Error("No hay filas ni tarjetas seleccionables en el picker. Contexto=" + preview);
}
`;
  replaceOrThrow(
    "selectFromPicker",
    /async function selectFromPicker\(page, searchButtonRx, preferText = "", options = \{\}\) \{[\s\S]*?\n\}\n\nasync function submitModal/,
    selectFn + "\n\nasync function submitModal"
  );
}

// 4) createAndEmitSale fix comprador/producto
{
  const re = /async function createAndEmitSale\([\s\S]*?\n\}\n\nasync function createSalidaAndCompleteTask/;
  if (!re.test(code)) throw new Error("No encontré función createAndEmitSale");
  code = code.replace(re, (blk) => {
    let out = blk;
    out = out.replace(
      /await selectFromPicker\(operadorPage,\s*\/Buscar Comprador\/i,\s*sim\.maestros\.comprador\.nombre\);/,
`await selectFromPicker(
      operadorPage,
      [/Buscar Comprador/i, /Seleccionar Comprador/i, /Buscar Cliente/i, /Comprador/i],
      sim.maestros.comprador.nombre
    );`
    );
    out = out.replace(
      /await selectFromPicker\(operadorPage,\s*\/Buscar Producto\/i,\s*""\);/,
      'await selectFromPicker(operadorPage, [/Buscar Producto/i, /Seleccionar Producto/i, /Producto/i], "");'
    );
    return out;
  });
}

// 5) Hooks diagnóstico en runOneSimulation
replaceIfMissing(
  "RUN_DIAG.stepStart(sim.id",
  "runStep diagnóstico",
  /const runStep = async \(label, fn\) => \{[\s\S]*?\n  \};/,
`const runStep = async (label, fn) => {
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
  };`
);

replaceIfMissing(
  "RUN_DIAG.setTotal(sim.id, steps.length);",
  "setTotal diagnóstico",
  /monitor\.setTotal\(sim\.id, steps\.length\);/,
`monitor.setTotal(sim.id, steps.length);
  RUN_DIAG.setTotal(sim.id, steps.length);`
);

replaceIfMissing(
  "RUN_DIAG.completeSimulation(sim.id);",
  "complete diagnóstico",
  /monitor\.complete\(sim\.id\);/,
`monitor.complete(sim.id);
    RUN_DIAG.completeSimulation(sim.id);`
);

replaceIfMissing(
  "RUN_DIAG.failSimulation(sim.id, error);",
  "fail diagnóstico",
  /monitor\.fail\(sim\.id, error\);/,
`monitor.fail(sim.id, error);
    RUN_DIAG.failSimulation(sim.id, error);`
);

// 6) main: meta + register + writeFinal
replaceIfMissing(
  "RUN_DIAG.setMeta(",
  "meta diagnóstico",
  /const monitor = new Monitor\(scenario\.simulaciones\);\s*monitor\.startLoop\(\);/,
`const monitor = new Monitor(scenario.simulaciones);
  RUN_DIAG.setMeta({
    frontUrl,
    downloadsDir,
    headed: args.headed,
    timeoutMs: args.timeoutMs,
    scenarioPath: SCENARIO_PATH
  });
  monitor.startLoop();`
);

replaceIfMissing(
  "RUN_DIAG.registerSimulation(sim, browserType);",
  "register sim diagnóstico",
  /console\.log\(`\[\$\{sim\.id\}\] Excel OK -> hoja: \$\{meta\.sheet\}, filas: \$\{meta\.rows\}`\);/,
`console.log(\`[\${sim.id}] Excel OK -> hoja: \${meta.sheet}, filas: \${meta.rows}\`);
      RUN_DIAG.registerSimulation(sim, browserType);`
);

replaceIfMissing(
  "const reportFiles = RUN_DIAG.writeFinal(results, null);",
  "writeFinal success",
  /monitor\.stopLoop\(\);\s*const failures = results\.filter/,
`monitor.stopLoop();
    const reportFiles = RUN_DIAG.writeFinal(results, null);
    if (reportFiles) {
      console.log("[diagnostico] Markdown: " + reportFiles.mdPath);
      console.log("[diagnostico] JSON: " + reportFiles.jsonPath);
    }
    const failures = results.filter`
);

replaceIfMissing(
  "const reportFiles = RUN_DIAG.writeFinal([], error);",
  "writeFinal catch",
  /monitor\.stopLoop\(\);\s*console\.error\("\\n❌ Error global:", error\?\.stack \|\| error\?\.message \|\| error\);/,
`monitor.stopLoop();
    const reportFiles = RUN_DIAG.writeFinal([], error);
    if (reportFiles) {
      console.log("[diagnostico] Markdown: " + reportFiles.mdPath);
      console.log("[diagnostico] JSON: " + reportFiles.jsonPath);
    }
    console.error("\\n❌ Error global:", error?.stack || error?.message || error);`
);

fs.writeFileSync(runnerPath, code, "utf8");
console.log("Patch OK -> " + runnerPath);
'@ | Set-Content -Path $patcher -Encoding UTF8

Write-Host "Patcher escrito en: $patcher"

node $patcher $RunnerPath
if ($LASTEXITCODE -ne 0) {
  throw "Falló el patcher Node (exit code $LASTEXITCODE)"
}

node --check $RunnerPath
if ($LASTEXITCODE -ne 0) {
  throw "run.mjs quedó con error de sintaxis"
}

Write-Host ""
Write-Host "OK ✅ Parche aplicado correctamente."
Write-Host "Diagnósticos: C:\polaria-ui-runner\diagnosticos"