// 별도 작업 공간(Web Worker)에서 Pyodide(Python)를 실행한다.
// 화면이 멈추지 않도록 엑셀 해석과 행렬 계산은 모두 여기서 처리한다.
const PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let py = null;
let chain = Promise.resolve();

async function boot() {
  if (py) return;
  postMessage({ type: 'status', text: 'Python 실행 환경(Pyodide) 내려받는 중…' });
  importScripts(PYODIDE + 'pyodide.js');
  py = await loadPyodide({ indexURL: PYODIDE });
  postMessage({ type: 'status', text: 'numpy · openpyxl 설치 중…' });
  await py.loadPackage(['numpy', 'micropip']);
  await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
  const res = await fetch('./model.py', { cache: 'no-cache' });
  if (!res.ok) throw Error('model.py를 불러올 수 없습니다 (HTTP ' + res.status + ')');
  await py.runPythonAsync(await res.text());
}

function callPy(fn, arg) {
  py.globals.set('_arg', arg);
  try {
    return py.runPython(fn + '(_arg)');
  } finally {
    py.globals.delete('_arg');
  }
}

async function handle({ id, type, payload }) {
  try {
    let result;
    if (type === 'boot') {
      await boot();
      result = '{}';
    } else if (type === 'bytes') {
      // payload: { fn: 'load_level' | 'load_classification', level?, buffer }
      await boot();
      py.globals.set('_raw', new Uint8Array(payload.buffer));
      await py.runPythonAsync('_raw_bytes = _raw.to_py()');
      result = payload.level
        ? py.runPython(`${payload.fn}(${JSON.stringify(payload.level)}, _raw_bytes)`)
        : py.runPython(`${payload.fn}(_raw_bytes)`);
      py.runPython('del _raw_bytes');
      py.globals.delete('_raw');
    } else if (type === 'call') {
      // payload: { fn, arg } — arg는 문자열(분류 코드 또는 JSON)
      result = callPy(payload.fn, payload.arg);
    } else {
      throw Error('알 수 없는 요청: ' + type);
    }
    postMessage({ id, ok: true, result: JSON.parse(result) });
  } catch (err) {
    let msg = String(err && err.message ? err.message : err);
    const m = msg.match(/(ValueError|KeyError|IndexError|ZeroDivisionError|LinAlgError)[^\n]*$/m);
    if (m) msg = m[0];
    postMessage({ id, ok: false, error: msg });
  }
}

onmessage = (e) => { chain = chain.then(() => handle(e.data)); };
