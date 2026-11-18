// Pyodide Web Worker: numpy 및 openpyxl 로드
importScripts("https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js");

let pyodide = null;
let initialized = false;

async function ensurePyodide() {
  if (initialized) return;
  self.postMessage({ type: "status", message: "Pyodide(WebAssembly Python) 로딩 중..." });
  pyodide = await loadPyodide();

  self.postMessage({ type: "status", message: "numpy 및 openpyxl 패키지 로딩 중..." });
  await pyodide.loadPackage("numpy");
  try {
    await pyodide.loadPackage("openpyxl");
  } catch (_) {
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");
  }

  self.postMessage({ type: "status", message: "model.py 분석 모듈 초기화 중..." });
  const resp = await fetch("./model.py?v=" + Date.now(), { cache: "no-store" });
  if (!resp.ok) throw new Error("model.py를 불러올 수 없습니다.");
  const modelCode = await resp.text();
  pyodide.FS.writeFile("model.py", modelCode);
  pyodide.runPython("import model");
  initialized = true;
}

self.onmessage = async (event) => {
  const { id, action, payload } = event.data;
  try {
    await ensurePyodide();

    let resultJson = "{}";
    if (action === "init") {
      resultJson = JSON.stringify({ ok: true });
    } else if (action === "load_file") {
      const { filename, bytes } = payload;
      const uint8 = new Uint8Array(bytes);
      const safePath = "/tmp/" + filename.replace(/[^a-zA-Z0-9._-]/g, "_") + "_" + Date.now() + ".xlsx";
      pyodide.FS.writeFile(safePath, uint8);
      pyodide.globals.set("_arg_filename", filename);
      pyodide.globals.set("_arg_path", safePath);
      resultJson = pyodide.runPython("model.load_excel_file(_arg_filename, _arg_path)");
      try { pyodide.FS.unlink(safePath); } catch (_) {}
    } else if (action === "get_state") {
      resultJson = pyodide.runPython("model.get_state_summary()");
    } else if (action === "get_sheet_preview") {
      pyodide.globals.set("_arg_json", JSON.stringify(payload));
      resultJson = pyodide.runPython("model.get_sheet_preview(_arg_json)");
    } else if (action === "search_classification") {
      pyodide.globals.set("_arg_json", JSON.stringify(payload));
      resultJson = pyodide.runPython("model.search_classification(_arg_json)");
    } else if (action === "get_structure") {
      pyodide.globals.set("_arg_json", JSON.stringify(payload));
      resultJson = pyodide.runPython("model.get_structure_analysis(_arg_json)");
    } else if (action === "get_sector_detail") {
      pyodide.globals.set("_arg_json", JSON.stringify(payload));
      resultJson = pyodide.runPython("model.get_sector_detail(_arg_json)");
    } else if (action === "run_simulation") {
      pyodide.globals.set("_arg_json", JSON.stringify(payload));
      resultJson = pyodide.runPython("model.run_simulation(_arg_json)");
    } else if (action === "verify_all") {
      resultJson = pyodide.runPython("model.verify_all_levels()");
    } else if (action === "drilldown") {
      pyodide.globals.set("_arg_json", JSON.stringify(payload));
      resultJson = pyodide.runPython("model.get_drilldown(_arg_json)");
    } else {
      throw new Error("알 수 없는 요청입니다: " + action);
    }

    self.postMessage({ id, ok: true, data: JSON.parse(resultJson) });
  } catch (err) {
    self.postMessage({ id, ok: false, error: err.message || String(err) });
  }
};
