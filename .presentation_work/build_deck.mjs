import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const WORK_DIR = "/Users/torgerd/Documents/GitHub/quest_editor/.presentation_work";
const SCREEN_DIR = path.join(WORK_DIR, "screens");
const RENDER_DIR = path.join(WORK_DIR, "rendered");
const FINAL_PPTX = "/Users/torgerd/Documents/GitHub/quest_editor/outputs/Quest_Editor_prezentacja_konferencyjna_v31_22.pptx";

const W = 1280;
const H = 720;
const M = 42;
const FONT = "Helvetica Neue";
const MONO = "SF Mono";
const C = {
  ink: "#111827",
  muted: "#6B7280",
  faint: "#9CA3AF",
  rule: "#D8DADF",
  panel: "#F2F2F2",
  paper: "#FFFFFF",
  warm: "#FAFAF7",
  purple: "#7C3AED",
  purpleSoft: "#F1EAFE",
  green: "#059669",
  greenSoft: "#E8F5EF",
  orange: "#EA580C",
  red: "#B42318",
  redSoft: "#FDECEC",
};

let elementCounter = 0;

function newSlide(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  return slide;
}

function addText(slide, text, x, y, w, h, options = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    name: options.name || `text-${++elementCounter}`,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = text;
  box.text.style = {
    fontSize: options.fontSize ?? 24,
    typeface: options.typeface || FONT,
    color: options.color || C.ink,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    alignment: options.align || "left",
    verticalAlignment: options.valign || "top",
    autoFit: options.autoFit || "none",
    wrap: "square",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return box;
}

function addRect(slide, x, y, w, h, options = {}) {
  return slide.shapes.add({
    geometry: options.geometry || "rect",
    name: options.name || `shape-${++elementCounter}`,
    position: { left: x, top: y, width: w, height: h },
    fill: options.fill || C.panel,
    line: options.line || { style: "solid", fill: "none", width: 0 },
    borderRadius: options.radius,
    shadow: options.shadow,
  });
}

function addRule(slide, x, y, w, options = {}) {
  return slide.shapes.add({
    geometry: "line",
    name: options.name || `rule-${++elementCounter}`,
    position: { left: x, top: y, width: w, height: 0 },
    fill: "none",
    line: { style: "solid", fill: options.color || C.rule, width: options.width || 1 },
  });
}

function addVerticalRule(slide, x, y, h, options = {}) {
  return slide.shapes.add({
    geometry: "line",
    name: options.name || `vrule-${++elementCounter}`,
    position: { left: x, top: y, width: 0, height: h },
    fill: "none",
    line: { style: "solid", fill: options.color || C.rule, width: options.width || 1 },
  });
}

function addTitle(slide, title, number) {
  addText(slide, title, M, 34, 1196, 98, { fontSize: 48, bold: true, name: `slide-${number}-title` });
}

function addFooter(slide, number, label = "QUEST EDITOR · v31.22") {
  addText(slide, label, M, 675, 500, 20, { fontSize: 14, color: C.faint, name: `footer-${number}` });
  addText(slide, String(number).padStart(2, "0"), 1184, 675, 54, 20, {
    fontSize: 14,
    color: C.faint,
    align: "right",
    name: `footer-num-${number}`,
  });
}

function addAccentLabel(slide, text, x, y, color = C.purple, width = 240) {
  addRect(slide, x, y + 2, 5, 30, { fill: color });
  addText(slide, text, x + 18, y, width, 34, { fontSize: 20, bold: true, color });
}

function addRailCallout(slide, x, y, w, label, body, color) {
  addAccentLabel(slide, label, x, y, color, w - 18);
  addText(slide, body, x + 18, y + 44, w - 18, 104, { fontSize: 22, color: C.ink });
}

async function addScreenshot(slide, fileName, x, y, w, h, alt) {
  const bytes = await fs.readFile(path.join(SCREEN_DIR, fileName));
  addRect(slide, x - 2, y - 2, w + 4, h + 4, {
    fill: C.paper,
    line: { style: "solid", fill: C.rule, width: 1 },
    shadow: "shadow-sm",
    radius: 8,
  });
  slide.images.add({
    blob: bytes,
    contentType: "image/png",
    alt,
    fit: "contain",
    geometry: "roundRect",
    borderRadius: 6,
    position: { left: x, top: y, width: w, height: h },
  });
}

function addNotes(slide, talkTrack, sources = []) {
  const lines = Array.isArray(talkTrack) ? talkTrack : [talkTrack];
  const sourceLines = sources.length ? ["", "[Sources]", ...sources.map((source) => `- ${source}`)] : [];
  slide.speakerNotes.textFrame.setText([...lines, ...sourceLines].join("\n"));
  slide.speakerNotes.setVisible(true);
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(RENDER_DIR, { recursive: true });
  await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: W, height: H } });

  // 01 — cover: Codex Grid slide-01 silhouette.
  {
    const slide = newSlide(presentation);
    addText(slide, "PREZENTACJA KONFERENCYJNA · v31.22", M, 42, 640, 42, {
      fontSize: 19,
      color: C.muted,
      bold: true,
      name: "cover-kicker",
    });
    addText(slide, "Quest Editor", M, 165, 1040, 104, { fontSize: 86, bold: true, name: "cover-title" });
    addText(slide, "od tekstu do działającej fabuły", M, 270, 1040, 104, {
      fontSize: 56,
      color: C.purple,
      name: "cover-subtitle",
    });
    addText(slide, "Lekki format tekstowy, kontrola struktury i symulacja wyborów w jednym lokalnym narzędziu.", M, 505, 720, 90, {
      fontSize: 28,
      color: C.muted,
      name: "cover-description",
    });
    addRule(slide, 934, 68, 304, { color: C.green, width: 5 });
    addNotes(slide, [
      "Otwórz pytaniem: czy narzędzie dla narracji może zachować szybkość zwykłego pliku tekstowego, a jednocześnie dawać kontrolę znaną z edytorów grafowych?",
      "Quest Editor powstał jako odpowiedź na ten konflikt. Nie zastępuje sposobu pisania — dodaje do niego strukturę, diagnostykę i możliwość wykonania.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html",
    ]);
  }

  // 02 — problem: Codex Grid slide-02 large-message silhouette.
  {
    const slide = newSlide(presentation);
    addText(slide, "PROBLEM", M, 42, 220, 36, { fontSize: 20, bold: true, color: C.red });
    addText(slide, "TEKST ≠ WIDOCZNA STRUKTURA", 828, 42, 410, 36, {
      fontSize: 18,
      color: C.muted,
      align: "right",
    });
    addText(slide, "Plik rośnie liniowo.\nZłożoność questa — wykładniczo.", M, 206, 1130, 260, {
      fontSize: 68,
      bold: true,
      valign: "bottom",
    });
    addText(slide, "Nie widać martwych gałęzi, błędnych skoków ani tego, co naprawdę przejdzie gracz.", M, 520, 770, 80, {
      fontSize: 28,
      color: C.muted,
    });
    addNotes(slide, [
      "Sam format tekstowy świetnie skaluje się dla autora: jest szybki, wersjonowalny i łatwy do przenoszenia.",
      "Problem pojawia się wtedy, gdy rośnie liczba etapów, kotwic i przejść. Autor nadal widzi linie, choć gracz zobaczy graf decyzji.",
      "To właśnie tę różnicę perspektyw ma zlikwidować narzędzie.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html",
      "/Users/torgerd/Documents/GitHub/quest_editor/tests/fixtures/tkacze_snow_full.txt",
    ]);
  }

  // 03 — idea: one source, multiple derived views.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Jeden tekst, pięć perspektyw", 3);
    // Connectors first, behind the nodes.
    addRule(slide, 376, 389, 144, { color: C.rule, width: 2 });
    [211, 287, 363, 439, 515].forEach((yy) => addRule(slide, 720, yy, 112, { color: C.rule, width: 2 }));
    addVerticalRule(slide, 720, 211, 304, { color: C.rule, width: 2 });
    addRect(slide, 42, 246, 334, 286, { fill: C.warm, line: { style: "solid", fill: C.rule, width: 1 } });
    addText(slide, "TEKST", 74, 280, 260, 58, { fontSize: 42, bold: true, color: C.purple });
    addText(slide, "Autor pisze dialogi, etapy, kotwice i warunki w lekkim DSL-u.", 74, 358, 260, 112, {
      fontSize: 24,
      color: C.ink,
    });
    addRect(slide, 520, 316, 200, 146, { fill: C.ink, radius: 8 });
    addText(slide, "MODEL\nQUESTA", 538, 342, 164, 92, {
      fontSize: 27,
      bold: true,
      color: C.paper,
      align: "center",
      valign: "middle",
    });
    const views = [
      ["Tekst", "pisanie i autocomplete", C.purple],
      ["Graf", "lokalne rozgałęzienia", C.purple],
      ["Flow questa", "mapa etapów i warunków", C.green],
      ["Character Lens", "review w kontekście postaci", C.green],
      ["Play", "wykonanie ścieżki gracza", C.orange],
    ];
    views.forEach(([label, desc, color], i) => {
      const y = 181 + i * 76;
      addText(slide, label, 832, y, 280, 34, { fontSize: 28, bold: true, color });
      addText(slide, desc, 832, y + 34, 360, 32, { fontSize: 20, color: C.muted });
    });
    addFooter(slide, 3);
    addNotes(slide, [
      "Centralne założenie brzmi: nie utrzymujemy pięciu kopii danych. Utrzymujemy jeden tekst i pięć perspektyw wyprowadzanych z tego samego modelu.",
      "Dzięki temu poprawka w dialogu natychmiast zmienia graf, mapę etapów, lens postaci i symulację.",
      "To eliminuje synchronizację między dokumentem pisarskim a osobnym schematem technicznym.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:1982",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:2856",
    ]);
  }

  // 04 — DSL: split layout derived from Codex Grid slide-08.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Prosty dla autora, jednoznaczny dla parsera", 4);
    addRect(slide, M, 157, 690, 468, { fill: "#F8F7F2", line: { style: "solid", fill: C.rule, width: 1 }, radius: 8 });
    const code = [
      "//BIT: Odprawa u kapitana",
      "##ETAP",
      "ETAP_1 ###Idź na odprawę ###Kapitan czeka.",
      "##DIALOG *WARNIK*",
      "*WARNIK*",
      "Wreszcie jesteście.",
      "*GRACZ*",
      "Pytam o szczegóły.{1.1}",
      "Ruszamy dalej.{ETAP_2}",
      "{1.1}*WARNIK*",
      "Utorczycy wzięli jeńców.",
    ].join("\n");
    addText(slide, code, 72, 185, 630, 410, { fontSize: 23, typeface: MONO, color: C.ink });
    const rules = [
      ["ETAP_n", "definiuje węzeł i jego opis", C.green],
      ["*GRACZ*", "zamienia kolejne linie w wybory", C.purple],
      ["{…}", "tworzy kotwicę, przejście lub koniec", C.orange],
    ];
    rules.forEach(([token, body, color], i) => {
      const y = 196 + i * 126;
      addText(slide, token, 784, y, 200, 40, { fontSize: 34, bold: true, color, typeface: MONO });
      addText(slide, body, 784, y + 48, 410, 58, { fontSize: 24, color: C.muted });
      if (i < rules.length - 1) addRule(slide, 784, y + 111, 410);
    });
    addFooter(slide, 4);
    addNotes(slide, [
      "Format wykorzystuje kilka rozpoznawalnych znaczników: etap, speaker, opcje gracza, kotwice i skoki.",
      "Autor nie klika w rozbudowane formularze; pisze naturalnym rytmem, a składnia dopowiada strukturę.",
      "Parser normalizuje identyfikatory etapów i rozwiązuje kotwice w obrębie konkretnego bloku dialogowego.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:1962",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:1982",
    ]);
  }

  // 05 — execution pipeline.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Jeden model napędza wszystkie widoki", 5);
    addRule(slide, 144, 368, 992, { color: C.rule, width: 3 });
    const stages = [
      ["01", "PARSE", "tekst → etapy, dialogi, opcje, krawędzie", C.purple],
      ["02", "MODEL", "dwa grafy: dialogowy i międzyetapowy", C.ink],
      ["03", "LINT", "błędy skoków, kotwic i granic gałęzi", C.red],
      ["04", "WIDOKI", "SVG, lens, play, review i eksport", C.green],
    ];
    stages.forEach(([num, title, body, color], i) => {
      const x = 74 + i * 305;
      addRect(slide, x + 60, 350, 20, 20, { geometry: "ellipse", fill: color });
      addText(slide, num, x, 214, 72, 34, { fontSize: 20, color: C.faint, bold: true });
      addText(slide, title, x, 254, 230, 48, { fontSize: 34, bold: true, color });
      addText(slide, body, x, 410, 236, 118, { fontSize: 23, color: C.ink });
    });
    addText(slide, "Zmiana źródła unieważnia cache; po debounce odświeżają się wyłącznie widoki pochodne.", 180, 576, 920, 52, {
      fontSize: 24,
      color: C.muted,
      align: "center",
    });
    addFooter(slide, 5);
    addNotes(slide, [
      "Silnik najpierw parsuje tekst do modelu, a potem buduje osobny graf dialogów i graf etapów.",
      "Linter pracuje na strukturze, więc potrafi nazwać błąd i wskazać linię źródła.",
      "Widoki są pochodne: po zmianie tekstu współdzielony parse i lint są buforowane, a odświeżenie jest opóźniane tak, by edytor pozostał responsywny.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:1982",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:2372",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:2748",
    ]);
  }

  // 06 — screenshot Text + Flow.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Piszesz tekst i od razu widzisz mapę questa", 6);
    await addScreenshot(slide, "01_text_flow.png", 42, 150, 918, 516, "Quest Editor: tekst po lewej i Flow questa po prawej");
    addRailCallout(slide, 992, 188, 246, "PISANIE", "Kolorowanie składni zachowuje rytm zwykłego pliku tekstowego.", C.purple);
    addRailCallout(slide, 992, 405, 246, "ORIENTACJA", "Flow grupuje etapy w bity i pokazuje relacje między nimi.", C.green);
    addFooter(slide, 6);
    addNotes(slide, [
      "To jest podstawowy tryb pracy: po lewej pozostaje pełny tekst, po prawej mapa etapów.",
      "Autor nie musi przerywać pisania, żeby przełączyć się do osobnego narzędzia diagramowego.",
      "Zaznaczenia, nawigacja i skupienie działają między widokami, więc mapa służy również do poruszania się po źródle.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/.presentation_work/screens/01_text_flow.png",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html",
    ]);
  }

  // 07 — screenshot Graph + Play.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Graf ujawnia błąd, Play pokazuje konsekwencję", 7);
    await addScreenshot(slide, "02_graph_play.png", 42, 150, 918, 516, "Quest Editor: graf dialogu, komunikat lintera i tryb Play");
    addRailCallout(slide, 992, 180, 246, "LINT", "Nielegalny skok speakera jest opisany i podlinkowany do źródła.", C.red);
    addRailCallout(slide, 992, 420, 246, "PLAY", "Ten sam model zatrzymuje się na wyborze i wykonuje decyzję gracza.", C.orange);
    addFooter(slide, 7);
    addNotes(slide, [
      "Na tym ujęciu widać najważniejszą pętlę jakościową: struktura mówi nie tylko, że coś jest niepoprawne, ale dlaczego.",
      "Obok działa symulator. Nie renderuje statycznego podglądu; wykonuje dialog, obsługuje kotwice, przejścia, warunki i cofanie do wcześniejszego wyboru.",
      "Autor może więc przejść od diagnozy do odtworzenia zachowania bez eksportu do gry.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/.presentation_work/screens/02_graph_play.png",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:2372",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:9311",
    ]);
  }

  // 08 — screenshot Character Lens.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Character Lens porządkuje review postaci", 8);
    await addScreenshot(slide, "03_character_text.png", 42, 150, 918, 516, "Quest Editor: Character Lens z blokami postaci i kontekstem źródła");
    addRailCallout(slide, 992, 180, 246, "SOCZEWKA", "Zbiera wszystkie kwestie postaci wraz z sąsiadującą interakcją.", C.purple);
    addRailCallout(slide, 992, 420, 246, "EDYCJA", "Zmiana bloku wraca do dokładnego zakresu linii w źródle.", C.green);
    addFooter(slide, 8);
    addNotes(slide, [
      "Character Lens odpowiada na realny problem review: kwestie jednej postaci są rozrzucone po wielu etapach i gałęziach.",
      "Widok indeksuje wystąpienia, liczy bloki i słowa, a następnie pokazuje kwestie postaci razem z sąsiadującymi reakcjami.",
      "Blok można poprawić bezpośrednio, a narzędzie bezpiecznie podmienia odpowiadający mu zakres linii.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/.presentation_work/screens/03_character_text.png",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:2288",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:6915",
    ]);
  }

  // 09 — four principles, Codex Grid slide-13 silhouette without cards.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Cztery zasady utrzymują spójność", 9);
    addRule(slide, 42, 382, 1196);
    addVerticalRule(slide, 640, 170, 430);
    const principles = [
      ["01", "Tekst jest źródłem prawdy", "Graf i symulacja nigdy nie żyją własnym życiem.", C.purple, 42, 178],
      ["02", "Błąd jest częścią modelu", "Diagnostyka zna typ problemu i linię powrotu.", C.red, 684, 178],
      ["03", "Widoki są pochodne", "Jeden parse zasila wszystkie perspektywy i cache.", C.green, 42, 418],
      ["04", "Local-first, eksport jawny", "Szkic zapisuje się lokalnie; plik i HTML są świadomą decyzją.", C.orange, 684, 418],
    ];
    principles.forEach(([num, title, body, color, x, y]) => {
      addText(slide, num, x, y, 64, 34, { fontSize: 20, bold: true, color: C.faint });
      addText(slide, title, x + 78, y - 4, 470, 44, { fontSize: 32, bold: true, color });
      addText(slide, body, x + 78, y + 54, 470, 78, { fontSize: 24, color: C.ink });
    });
    addFooter(slide, 9);
    addNotes(slide, [
      "Te założenia są ważniejsze niż lista funkcji, bo określają granice systemu.",
      "Tekst jest kanoniczny; błędy należą do modelu; widoki nie kopiują danych; lokalny autosave nie udaje świadomego zapisu pliku.",
      "Dzięki temu narzędzie pozostaje lekkie, a zachowanie poszczególnych trybów daje się przewidzieć.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:1381",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:2748",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:8348",
    ]);
  }

  // 10 — implementation metrics, Codex Grid slide-19 silhouette.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Wykonanie: lokalna aplikacja bez builda", 10);
    addText(slide, "HTML, CSS i JavaScript działają bez serwera, a SVG powstaje bezpośrednio z modelu questa.", 42, 126, 1040, 62, {
      fontSize: 26,
      color: C.muted,
    });
    const metrics = [
      ["1", "plik aplikacji", "Cały interfejs i silnik w index.html", C.purple],
      ["0", "zewnętrznych bibliotek", "Brak runtime'owych skryptów i CSS z sieci", C.green],
      ["5 s", "autosave", "Szkic w localStorage + jawny zapis pliku", C.orange],
    ];
    metrics.forEach(([stat, label, body, color], i) => {
      const x = 42 + i * 411;
      addRect(slide, x, 248, 374, 330, { fill: C.panel, radius: 8 });
      addText(slide, stat, x + 32, 286, 310, 108, { fontSize: 78, bold: true, color });
      addText(slide, label, x + 32, 408, 310, 46, { fontSize: 30, bold: true, color: C.ink });
      addText(slide, body, x + 32, 476, 310, 76, { fontSize: 22, color: C.muted });
    });
    addText(slide, "10 398 linii · 366 nazwanych funkcji · 133 rejestracje zdarzeń", 42, 616, 1100, 32, {
      fontSize: 19,
      color: C.faint,
    });
    addFooter(slide, 10);
    addNotes(slide, [
      "Wykonanie jest celowo pragmatyczne: jeden plik aplikacji, jedna IIFE, zero zewnętrznych skryptów i arkuszy stylów.",
      "To ogranicza narzut uruchomienia i dystrybucji, ale przenosi odpowiedzialność na dyscyplinę modułów wewnętrznych i testy strukturalne.",
      "Grafy są generowane jako SVG, a szkic jest automatycznie zapisywany co pięć sekund w localStorage.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html",
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html:8348",
    ]);
  }

  // 11 — test evidence.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Testujemy strukturę, nie zrzuty ekranu", 11);
    addText(slide, "Test snapshot uruchamia kod edytora w pamięci i kontroluje parser, lint, osiągalność BFS oraz layout Flow.", 42, 126, 1080, 70, {
      fontSize: 26,
      color: C.muted,
    });
    const metrics = [
      ["119", "etapów", "wszystkie osiągalne od startu", C.green],
      ["126", "krawędzi", "graf międzyetapowy pełnego questa", C.purple],
      ["14/14", "sprawdzeń", "zero niezaliczonych testów", C.ink],
    ];
    metrics.forEach(([stat, label, body, color], i) => {
      const x = 42 + i * 411;
      addRect(slide, x, 248, 374, 330, { fill: C.panel, radius: 8 });
      addText(slide, stat, x + 30, 286, 314, 112, { fontSize: 72, bold: true, color });
      addText(slide, label, x + 30, 410, 314, 42, { fontSize: 30, bold: true });
      addText(slide, body, x + 30, 476, 314, 70, { fontSize: 22, color: C.muted });
    });
    addText(slide, "Pełny fixture „Tkacze Snów” · 119/119 etapów otrzymuje kolumnę i pozycję Y w layoucie Flow", 42, 616, 1130, 34, {
      fontSize: 19,
      color: C.faint,
    });
    addFooter(slide, 11);
    addNotes(slide, [
      "Test nie porównuje obrazków. Wyciąga z aplikacji funkcje parse, lint i layoutFlow, a następnie uruchamia je na dwóch fixture'ach.",
      "Najważniejszy test to pełny quest Tkacze Snów: 119 etapów, 126 krawędzi i pełna osiągalność od startu.",
      "Dodatkowo każdy etap musi otrzymać poprawną pozycję w layoucie Flow. Błędy realnego fixture'a są jawnie kalibrowane w baseline, a nie zamiatane pod dywan.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/tests/snapshot.js",
      "/Users/torgerd/Documents/GitHub/quest_editor/tests/fixtures/tkacze_snow_full.txt",
    ]);
  }

  // 12 — future-state workflow, timeline silhouette derived from slide-17.
  {
    const slide = newSlide(presentation);
    addTitle(slide, "Nowy rytm pracy skraca pętlę do decyzji", 12);
    addRule(slide, 96, 350, 1088, { color: C.ink, width: 2 });
    const steps = [
      ["01", "NAPISZ", "tekst i autocomplete", C.purple],
      ["02", "SPRAWDŹ", "lint w tej samej chwili", C.red],
      ["03", "ZOBACZ", "graf i Flow questa", C.green],
      ["04", "PRZEJDŹ", "Play z wyborami", C.orange],
      ["05", "ZREVIEWUJ", "lens, diff i eksport", C.ink],
    ];
    steps.forEach(([num, label, body, color], i) => {
      const x = 72 + i * 242;
      addRect(slide, x + 64, 339, 22, 22, { geometry: "ellipse", fill: color });
      addText(slide, num, x, 242, 42, 28, { fontSize: 18, bold: true, color: C.faint });
      addText(slide, label, x, 278, 190, 40, { fontSize: 27, bold: true, color });
      addText(slide, body, x, 404, 190, 74, { fontSize: 22, color: C.ink });
    });
    addText(slide, "Każdy krok korzysta z tego samego tekstu — bez ręcznego przepisywania danych między narzędziami.", 164, 550, 952, 60, {
      fontSize: 26,
      color: C.muted,
      align: "center",
    });
    addFooter(slide, 12);
    addNotes(slide, [
      "Najważniejszym efektem nie jest pojedyncza funkcja, lecz skrócenie pętli pracy.",
      "Autor pisze, od razu dostaje kontrolę struktury, może zobaczyć mapę, przejść ścieżkę, zrobić review i wyeksportować wynik.",
      "Ponieważ każdy krok korzysta z tego samego źródła, znika koszt ręcznego utrzymywania zgodności między dokumentami.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html",
    ]);
  }

  // 13 — close: Codex Grid slide-26 silhouette, resolving the opening.
  {
    const slide = newSlide(presentation);
    addText(slide, "WNIOSEK", M, 42, 220, 36, { fontSize: 20, bold: true, color: C.green });
    addText(slide, "Quest staje się kodem,\nktóry nadal dobrze się czyta.", M, 172, 1060, 270, {
      fontSize: 72,
      bold: true,
      valign: "bottom",
    });
    addRule(slide, M, 488, 1196, { color: C.rule, width: 1 });
    const outcomes = [
      ["SZYBCIEJ PISAĆ", "bez opuszczania tekstu", C.purple],
      ["WCZEŚNIEJ WYKRYWAĆ", "zanim błąd trafi do integracji", C.red],
      ["ŚWIADOMIE TESTOWAĆ", "tak jak przejdzie to gracz", C.green],
    ];
    outcomes.forEach(([title, body, color], i) => {
      const x = 42 + i * 411;
      addText(slide, title, x, 532, 360, 34, { fontSize: 22, bold: true, color });
      addText(slide, body, x, 578, 360, 46, { fontSize: 23, color: C.muted });
    });
    addText(slide, "QUEST EDITOR · PYTANIA?", 980, 42, 258, 34, { fontSize: 18, color: C.faint, align: "right" });
    addNotes(slide, [
      "Wróć do pytania z początku: czy można zachować szybkość tekstu i zyskać kontrolę grafu? W tym projekcie odpowiedzią jest wykonywalny model wyprowadzany z tekstu.",
      "Quest Editor traktuje narrację jak kod tylko tam, gdzie pomaga to autorowi: w walidacji, wizualizacji, testowaniu i review.",
      "Zamknij trzema efektami: szybciej pisać, wcześniej wykrywać i świadomie testować. Następnie przejdź do pytań lub krótkiego demo na żywo.",
    ], [
      "/Users/torgerd/Documents/GitHub/quest_editor/index.html",
      "/Users/torgerd/Documents/GitHub/quest_editor/tests/snapshot.js",
    ]);
  }

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(path.join(RENDER_DIR, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(RENDER_DIR, `${stem}.layout.json`), await layout.text());
  }

  await writeBlob(path.join(RENDER_DIR, "deck-montage.webp"), await presentation.export({ format: "webp", montage: true, scale: 1 }));
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);
  console.log(FINAL_PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
