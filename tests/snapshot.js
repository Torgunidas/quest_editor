// Quest editor snapshot tests (v2, kalibracja: v31.22).
//
// Użycie:
//   node tests/snapshot.js [ścieżka/do/edytora.html]
//   (domyślnie: index.html w katalogu głównym repo)
//
// Co robi:
//   Wyciąga IIFE edytora z pliku HTML, uruchamia go w Node ze stubem DOM
//   i wstrzykniętym hookiem testowym (bez modyfikacji pliku produktu),
//   po czym sprawdza parser, lint, osiągalność BFS i layout Flow
//   na dwóch fixture'ach: syntetycznym mini-quescie i pełnych Tkaczach Snów.
//
// Czego NIE sprawdza:
//   - jakości wizualnej renderowania (to nadal feedback wzrokowy),
//   - wydajności (profiler w produkcie),
//   - handlerów zdarzeń DOM.
//
// Aktualizacja baseline: gdy ŚWIADOMIE zmieniasz strukturę (np. dopisujesz
// etap do Tkaczy), popraw wartości w obiekcie BASELINE poniżej.

'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ============================================================
// Baseline (zmierzone na v31.22 + tkacze_snow_full.txt)
// ============================================================
const BASELINE = {
  tkacze: {
    stages: 119,
    stageEdges: 126,
    lintIssues: 4,          // stan realny Tkaczy na v31.22: 2x branch-boundary (ETAP_33/46), 2x niezamkniety [o] (ETAP_76)
    reachableFromStart: 119,
    flowStagesWithLayout: 119
  },
  synthetic: {
    stages: 3,
    stageEdges: 3,
    lintIssues: 0
  }
};

// ============================================================
// CLI
// ============================================================
const fileArg = process.argv.slice(2).find(a => !a.startsWith('--'));
const editorPath = fileArg
  ? path.resolve(fileArg)
  : path.resolve(__dirname, '..', 'index.html');

if (!fs.existsSync(editorPath)) {
  console.error('Nie znaleziono pliku edytora: ' + editorPath);
  process.exit(2);
}

// ============================================================
// Stub DOM — samowystarczalny element zwracający no-opy
// ============================================================
function makeStubElement() {
  const el = {};
  const listHandler = {
    add() {}, remove() {}, toggle() {}, contains() { return false; }
  };
  return new Proxy(el, {
    get(target, prop) {
      if (prop === 'style') return target.__style || (target.__style = new Proxy({}, { get: () => '', set: () => true }));
      if (prop === 'classList') return listHandler;
      if (prop === 'dataset') return target.__dataset || (target.__dataset = {});
      if (prop === 'children' || prop === 'childNodes' || prop === 'options' ||
          prop === 'selectedOptions' || prop === 'rows' || prop === 'cells' ||
          prop === 'files' || prop === 'attributes') return [];
      if (prop === 'value' || prop === 'textContent' || prop === 'innerHTML' || prop === 'outerHTML') {
        return target['__' + String(prop)] !== undefined ? target['__' + String(prop)] : '';
      }
      if (prop === 'selectionStart' || prop === 'selectionEnd' || prop === 'scrollTop' ||
          prop === 'scrollLeft' || prop === 'scrollHeight' || prop === 'clientHeight' ||
          prop === 'clientWidth' || prop === 'offsetTop' || prop === 'offsetHeight' ||
          prop === 'offsetWidth') return 0;
      if (prop === 'getBoundingClientRect') return () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 });
      if (prop === 'querySelector') return () => null;
      if (prop === 'querySelectorAll') return () => [];
      if (prop === 'getContext') return () => null;
      if (prop === 'closest') return () => null;
      if (prop === 'focus' || prop === 'blur' || prop === 'click' || prop === 'remove' ||
          prop === 'appendChild' || prop === 'append' || prop === 'prepend' ||
          prop === 'insertBefore' || prop === 'removeChild' || prop === 'replaceChildren' ||
          prop === 'setAttribute' || prop === 'removeAttribute' || prop === 'addEventListener' ||
          prop === 'removeEventListener' || prop === 'setRangeText' ||
          prop === 'setSelectionRange' || prop === 'scrollIntoView' || prop === 'dispatchEvent') {
        return () => makeStubElement();
      }
      if (prop === 'getAttribute') return () => null;
      if (prop === 'hasAttribute') return () => false;
      if (prop === 'parentElement' || prop === 'parentNode' || prop === 'firstChild' ||
          prop === 'lastChild' || prop === 'nextSibling' || prop === 'previousSibling' ||
          prop === 'nextElementSibling' || prop === 'ownerDocument') return null;
      if (prop === Symbol.toPrimitive) return () => '';
      if (typeof prop === 'symbol') return undefined;
      if (target[prop] !== undefined) return target[prop];
      // Nieznana właściwość: zwróć no-op funkcję (najbezpieczniejsze dla łańcuchów wywołań)
      return () => makeStubElement();
    },
    set(target, prop, val) {
      if (prop === 'value' || prop === 'textContent' || prop === 'innerHTML' || prop === 'outerHTML') {
        target['__' + String(prop)] = val;
      } else {
        target[prop] = val;
      }
      return true;
    }
  });
}

function buildSandbox() {
  const storage = {};
  const sandbox = {
    console,
    setTimeout: () => 0,
    clearTimeout: () => {},
    setInterval: () => 0,
    clearInterval: () => {},
    requestAnimationFrame: () => 0,
    cancelAnimationFrame: () => {},
    localStorage: {
      getItem: k => (k in storage ? storage[k] : null),
      setItem: (k, v) => { storage[k] = String(v); },
      removeItem: k => { delete storage[k]; }
    },
    navigator: { platform: 'test', userAgent: 'node-test', clipboard: { writeText: () => Promise.resolve() } },
    location: { href: 'about:test', search: '', hash: '' },
    history: { replaceState: () => {}, pushState: () => {} },
    document: null,
    window: null,
    alert: () => {},
    confirm: () => true,
    prompt: () => null,
    Blob: function Blob() {},
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL: () => {} },
    FileReader: function FileReader() { this.readAsText = () => {}; },
    getComputedStyle: () => new Proxy({}, { get: () => '' }),
    matchMedia: () => ({ matches: false, addEventListener: () => {}, addListener: () => {} }),
    performance: { now: () => Date.now() },
    addEventListener: () => {},
    removeEventListener: () => {},
    dispatchEvent: () => true,
    innerWidth: 1920,
    innerHeight: 1080,
    devicePixelRatio: 2,
    scrollTo: () => {},
    getSelection: () => ({ removeAllRanges: () => {}, addRange: () => {}, rangeCount: 0 }),
    ResizeObserver: function ResizeObserver() { this.observe = () => {}; this.disconnect = () => {}; this.unobserve = () => {}; },
    MutationObserver: function MutationObserver() { this.observe = () => {}; this.disconnect = () => {}; },
    __qed_test: null
  };
  const doc = {
    getElementById: () => makeStubElement(),
    querySelector: () => makeStubElement(),
    querySelectorAll: () => [],
    createElement: () => makeStubElement(),
    createElementNS: () => makeStubElement(),
    createDocumentFragment: () => makeStubElement(),
    createTextNode: () => makeStubElement(),
    addEventListener: () => {},
    removeEventListener: () => {},
    body: makeStubElement(),
    documentElement: makeStubElement(),
    head: makeStubElement(),
    activeElement: null,
    caretRangeFromPoint: () => null,
    fonts: { ready: Promise.resolve() },
    hidden: false,
    visibilityState: 'visible'
  };
  sandbox.document = doc;
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  return sandbox;
}

// ============================================================
// Ekstrakcja i uruchomienie IIFE edytora z wstrzykniętym hookiem
// ============================================================
function loadEditorInternals(htmlPath) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const open = html.indexOf('<script>');
  const close = html.lastIndexOf('</script>');
  if (open === -1 || close === -1) throw new Error('Nie znaleziono bloku <script> w ' + htmlPath);
  let script = html.slice(open + '<script>'.length, close);

  // Wstrzyknij hook testowy tuż przed zamknięciem IIFE — tylko w pamięci,
  // plik produktu pozostaje nietknięty.
  const tail = '})();';
  const tailIdx = script.lastIndexOf(tail);
  if (tailIdx === -1) throw new Error('Nie znaleziono zamknięcia IIFE — zmieniła się struktura pliku?');
  const hook = `
  ;globalThis.__qed_test = {
    parse: typeof parse === 'function' ? parse : null,
    lint: typeof lint === 'function' ? lint : null,
    layoutFlow: typeof layoutFlow === 'function' ? layoutFlow : null,
    findReferencedStages: typeof findReferencedStages === 'function' ? findReferencedStages : null,
    findDefinedStages: typeof findDefinedStages === 'function' ? findDefinedStages : null
  };
  `;
  script = script.slice(0, tailIdx) + hook + script.slice(tailIdx);

  const sandbox = buildSandbox();
  vm.createContext(sandbox);
  try {
    vm.runInContext(script, sandbox, { filename: path.basename(htmlPath), timeout: 30000 });
  } catch (e) {
    // Inicjalizacja UI może się wywalić na stubie DOM PO zdefiniowaniu funkcji.
    // Jeśli hook zdążył się ustawić — kontynuujemy; jeśli nie — to realny błąd.
    if (!sandbox.__qed_test || !sandbox.__qed_test.parse) {
      throw new Error('Skrypt edytora nie doszedł do hooka testowego: ' + e.message);
    }
    console.log('  (uwaga: inicjalizacja UI rzuciła po zdefiniowaniu funkcji — ignoruję: ' + e.message + ')');
  }
  if (!sandbox.__qed_test || !sandbox.__qed_test.parse || !sandbox.__qed_test.lint) {
    throw new Error('Hook testowy nie wyeksportował parse/lint.');
  }
  return sandbox.__qed_test;
}

// ============================================================
// BFS po krawędziach etapów
// ============================================================
function bfsReachable(parsed, startId) {
  const adj = {};
  (parsed.stageEdges || []).forEach(e => {
    const from = e.from || e.source, to = e.to || e.target;
    if (!from || !to) return;
    (adj[from] = adj[from] || []).push(to);
  });
  const seen = new Set([startId]);
  const q = [startId];
  while (q.length) {
    const cur = q.shift();
    (adj[cur] || []).forEach(n => { if (!seen.has(n)) { seen.add(n); q.push(n); } });
  }
  return seen;
}

// ============================================================
// Fixture syntetyczny — 3 etapy, 2 przejścia, zero błędów
// ============================================================
const SYNTHETIC = `##QUEST###900
##NPC
TESTER###1001

##ETAP
ETAP_1 ###Start ###Opis startu.
##DIALOG *TESTER*
*TESTER*
Witaj w tescie.
*GRACZ*
Idz dalej.{ETAP_2}
Skocz na koniec.{ETAP_3}

##ETAP
ETAP_2 ###Srodek ###Opis srodka.
##DIALOG *TESTER*
*TESTER*
Jestes w srodku.
*GRACZ*
Koniec.{ETAP_3}

##ETAP
ETAP_3 ###Koniec ###Opis konca.
##DIALOG *TESTER*
*TESTER*
To juz koniec.
`;

// ============================================================
// Runner
// ============================================================
let passed = 0, failed = 0;
function assert(name, actual, expected) {
  const ok = actual === expected;
  if (ok) { passed++; console.log('  \u2713 ' + name + ' = ' + actual); }
  else { failed++; console.log('  \u2717 ' + name + ': oczekiwano ' + expected + ', jest ' + actual); }
}
function assertTrue(name, cond, detail) {
  if (cond) { passed++; console.log('  \u2713 ' + name); }
  else { failed++; console.log('  \u2717 ' + name + (detail ? ' — ' + detail : '')); }
}

console.log('Edytor: ' + editorPath);
const api = loadEditorInternals(editorPath);
console.log('Hook testowy załadowany. Funkcje: ' +
  Object.keys(api).filter(k => api[k]).join(', ') + '\n');

// --- Syntetyczny ---
console.log('[1/3] Fixture syntetyczny');
{
  const parsed = api.parse(SYNTHETIC);
  const issues = api.lint(parsed);
  assert('stages', (parsed.stages || []).length, BASELINE.synthetic.stages);
  assert('stageEdges', (parsed.stageEdges || []).length, BASELINE.synthetic.stageEdges);
  assert('lintIssues', issues.length, BASELINE.synthetic.lintIssues);
  if (issues.length) issues.forEach(i => console.log('      lint: ' + JSON.stringify(i).slice(0, 160)));
  const reach = bfsReachable(parsed, (parsed.stages[0] || {}).id);
  assert('reachableFromStart', reach.size, BASELINE.synthetic.stages);
}

// --- Tkacze Snów ---
console.log('[2/3] Tkacze Snów (pełny quest)');
{
  const tkaczePath = path.resolve(__dirname, 'fixtures', 'tkacze_snow_full.txt');
  assertTrue('fixture istnieje', fs.existsSync(tkaczePath), tkaczePath);
  const text = fs.readFileSync(tkaczePath, 'utf8');
  const parsed = api.parse(text);
  const issues = api.lint(parsed);
  assert('stages', (parsed.stages || []).length, BASELINE.tkacze.stages);
  assert('stageEdges', (parsed.stageEdges || []).length, BASELINE.tkacze.stageEdges);
  assert('lintIssues', issues.length, BASELINE.tkacze.lintIssues);
  if (issues.length && issues.length !== BASELINE.tkacze.lintIssues) {
    issues.slice(0, 5).forEach(i => console.log('      lint: ' + JSON.stringify(i).slice(0, 160)));
  }
  const startId = (parsed.stages[0] || {}).id;
  const reach = bfsReachable(parsed, startId);
  assert('reachableFromStart', reach.size, BASELINE.tkacze.reachableFromStart);

  // Spójność definicji/referencji na poziomie tekstu
  const defined = api.findDefinedStages(text);
  const referenced = api.findReferencedStages(text);
  const undefinedRefs = [...referenced].filter(r => !defined.has(r));
  assertTrue('brak referencji do niezdefiniowanych etapów', undefinedRefs.length === 0,
    undefinedRefs.slice(0, 5).join(', '));
}

// --- Layout Flow ---
console.log('[3/3] Integralność layoutu Flow (Tkacze)');
{
  const text = fs.readFileSync(path.resolve(__dirname, 'fixtures', 'tkacze_snow_full.txt'), 'utf8');
  const parsed = api.parse(text);
  if (!api.layoutFlow) {
    assertTrue('layoutFlow dostępny', false, 'hook nie wyeksportował layoutFlow');
  } else {
    const flow = api.layoutFlow(parsed);
    assertTrue('layoutFlow zwraca col + flowYUnits', !!(flow && flow.col && flow.flowYUnits),
      'kształt wyniku: ' + Object.keys(flow || {}).join(','));
    if (flow && flow.col && flow.flowYUnits) {
      const ids = (parsed.stages || []).map(s => s.id);
      const noCol = ids.filter(id => !Number.isFinite(flow.col[id]));
      const noY = ids.filter(id => !Number.isFinite(flow.flowYUnits[id]));
      assert('etapy z kolumną', ids.length - noCol.length, BASELINE.tkacze.flowStagesWithLayout);
      assert('etapy z pozycją Y', ids.length - noY.length, BASELINE.tkacze.flowStagesWithLayout);
      assertTrue('brak etapów bez layoutu', noCol.length === 0 && noY.length === 0,
        'bez kolumny: ' + noCol.slice(0,3).join(',') + ' | bez Y: ' + noY.slice(0,3).join(','));
    }
  }
}

console.log('\nWynik: ' + passed + ' \u2713 / ' + failed + ' \u2717');
process.exit(failed ? 1 : 0);
