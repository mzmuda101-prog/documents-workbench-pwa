// Etap 0 pakietu AI: pomiar jednego modelu językowego na urządzeniu (WebGPU).
// Każdy model działa w świeżym workerze — po pomiarze strona go zamyka, więc pamięć wraca.
// Tekst testów nie opuszcza urządzenia; z sieci idą tylko wagi modelu z Hugging Face.

const TRANSFORMERS_URL = "https://esm.sh/@huggingface/transformers@4.3.1";
const CACHE_KEY = "dwb-llm-bench-v1";

let lib = null;
const now = () => performance.now();
const post = (type, payload = {}) => self.postMessage({ type, ...payload });

async function loadLib() {
  lib ||= await import(TRANSFORMERS_URL);
  lib.env.useBrowserCache = true;
  lib.env.cacheKey = CACHE_KEY;
  lib.env.allowLocalModels = false;
  return lib;
}

// ---------- testy (po polsku, pod funkcje S1 / D2 / D3 z pakietu) ----------

const SPEED_PROMPT = [
  "Poniżej jest fragment regulaminu. Streść go w trzech punktach.",
  "",
  "§ 1. Regulamin określa zasady korzystania z pracowni komputerowej. Z pracowni mogą korzystać",
  "pracownicy działu oraz osoby, którym kierownik wydał pisemną zgodę. Pracownia jest czynna od",
  "poniedziałku do piątku w godzinach od 7:30 do 15:30, z wyjątkiem dni ustawowo wolnych od pracy.",
  "§ 2. Każdy użytkownik odpowiada za powierzony sprzęt. Usterki należy zgłaszać niezwłocznie, nie",
  "później niż w ciągu jednego dnia roboczego, osobiście lub pocztą elektroniczną. Zabrania się",
  "samodzielnej naprawy sprzętu, instalowania oprogramowania bez zgody administratora oraz",
  "wynoszenia sprzętu poza budynek. § 3. Dane zapisane na komputerach w pracowni są usuwane w",
  "każdy piątek po zamknięciu. Użytkownik powinien zapisać swoje pliki na nośniku służbowym albo w",
  "folderze działu. § 4. Naruszenie regulaminu może skutkować odebraniem prawa do korzystania z",
  "pracowni na okres do trzech miesięcy. Decyzję podejmuje kierownik po wysłuchaniu użytkownika.",
].join("\n");

const TASKS = [
  {
    id: "s1-filtr",
    label: "S1 · pytanie → filtr (JSON)",
    max: 60,
    system:
      "Zamieniasz pytanie po polsku na filtr arkusza. Dozwolone tryby: @open (sprawa w trakcie), " +
      "@closed (zamknięta), @nostart (brak startu), @first (pierwszy cykl), @last (ostatni cykl), " +
      "@daysN (trwa dłużej niż N dni, np. @days45). Odpowiedz WYŁĄCZNIE obiektem JSON: " +
      '{"szukaj": "<tekst do znalezienia>", "tryby": ["..."]}',
    user: "Pokaż sprawy Kowalskiej, które są nadal otwarte i trwają dłużej niż 30 dni.",
    check(out) {
      const j = firstJson(out);
      if (!j) return { ok: false, note: "brak poprawnego JSON" };
      const modes = new Set((j.tryby || []).map((s) => String(s).trim().toLowerCase()));
      const textOk = /kowalsk/i.test(String(j.szukaj || ""));
      const modesOk = modes.has("@open") && modes.has("@days30") && modes.size === 2;
      return { ok: textOk && modesOk, note: `szukaj ${textOk ? "✓" : "✗"}, tryby ${modesOk ? "✓" : "✗"}` };
    },
  },
  {
    id: "wyciag",
    label: "Wyciąganie danych z umowy (JSON)",
    max: 70,
    system:
      'Wyciągasz dane z tekstu. Odpowiedz WYŁĄCZNIE obiektem JSON: {"od": "RRRR-MM-DD", "do": "RRRR-MM-DD", "kwota": liczba}',
    user:
      "Umowa zostaje zawarta na czas określony od 1 marca 2026 r. do 28 lutego 2027 r. Wynagrodzenie " +
      "wynosi 4 500 zł brutto miesięcznie, płatne do 10. dnia następnego miesiąca.",
    check(out) {
      const j = firstJson(out);
      if (!j) return { ok: false, note: "brak poprawnego JSON" };
      const amount = Number(String(j.kwota ?? "").replace(/[\s ]/g, "").replace(",", "."));
      const hits = [j.od === "2026-03-01", j.do === "2027-02-28", amount === 4500];
      return { ok: hits.every(Boolean), note: `od ${hits[0] ? "✓" : "✗"}, do ${hits[1] ? "✓" : "✗"}, kwota ${hits[2] ? "✓" : "✗"}` };
    },
  },
  {
    id: "d3-pytanie",
    label: "D3 · pytanie do dokumentu",
    max: 60,
    system: "Odpowiadasz krótko, wyłącznie na podstawie podanego tekstu. Jeśli tekst nie zawiera odpowiedzi, napisz: brak w tekście.",
    user:
      "Tekst: Reklamację można złożyć w ciągu 30 dni od odbioru towaru. Sprzedawca rozpatruje ją w " +
      "terminie 14 dni od otrzymania. Zwrot pieniędzy następuje tą samą metodą płatności.\n\n" +
      "Pytanie: Ile dni ma sprzedawca na rozpatrzenie reklamacji?",
    check(out) {
      const ok = /\b14\b/.test(out) && !/\b30 dni\b.*rozpatr/i.test(out);
      return { ok, note: ok ? "podał 14 dni" : "zła odpowiedź" };
    },
  },
  {
    id: "d2-formalnie",
    label: "D2 · przepisz formalnie (do oceny)",
    max: 60,
    system: "Przepisujesz tekst po polsku w stylu formalnym. Odpowiedz jednym zdaniem, bez komentarza.",
    user: "No to wyślijcie nam te papiery do piątku, bo inaczej nie zdążymy.",
    check(out) {
      const t = out.trim();
      // Sens musi zostać: prośba o przesłanie dokumentów do piątku, bez zmiany na „nie zdążyliśmy”.
      const ok = t.length > 20 && t.length < 260 && /piąt/i.test(t) && /prosz|prosimy|uprzejm|uprasza/i.test(t) &&
        /przesła|dostarcz|wysła|przekaza/i.test(t) && /dokument/i.test(t) && !/papiery|nie zdążyliśmy/i.test(t);
      return { ok, note: ok ? "wygląda sensownie — do oceny" : "podejrzane — do oceny" };
    },
  },
];

function firstJson(text) {
  const m = String(text).match(/\{[\s\S]*?\}/);
  if (!m) return null;
  try {
    return JSON.parse(m[0]);
  } catch {
    return null;
  }
}

// ---------- ładowanie ----------

async function loadModel(spec, dtype) {
  const L = await loadLib();
  const files = new Map(); // plik → { loaded, total, start, end }
  let firstByteAt = 0;
  let lastReport = 0;
  const progress_callback = (p) => {
    if (!p.file) return;
    const f = files.get(p.file) || { loaded: 0, total: 0, start: now(), end: 0 };
    if (p.status === "progress") {
      firstByteAt ||= now();
      f.loaded = p.loaded || f.loaded;
      f.total = p.total || f.total;
    } else if (p.status === "done") {
      f.end = now();
      if (f.total) f.loaded = f.total;
    }
    files.set(p.file, f);
    const t = now();
    if (t - lastReport > 250 || p.status === "done") {
      lastReport = t;
      let loaded = 0;
      let total = 0;
      for (const v of files.values()) {
        loaded += v.loaded;
        total += v.total;
      }
      post("progress", { loaded, total });
    }
  };

  const t0 = now();
  let model;
  let tokenizer;
  if (spec.kind === "vlm") {
    const processor = await L.AutoProcessor.from_pretrained(spec.id, { progress_callback });
    tokenizer = processor.tokenizer;
    post("phase", { phase: "kompilacja" });
    const opts = { device: "webgpu", progress_callback };
    // Gemma 4 „mobile” ma typy zapisane w config.json (q2f16); pozostałe — ten sam typ dla każdej części.
    if (dtype !== "q2f16") opts.dtype = { embed_tokens: dtype, decoder_model_merged: dtype, vision_encoder: dtype };
    model = await L.AutoModelForImageTextToText.from_pretrained(spec.id, opts);
  } else {
    tokenizer = await L.AutoTokenizer.from_pretrained(spec.id, { progress_callback });
    post("phase", { phase: "kompilacja" });
    model = await L.AutoModelForCausalLM.from_pretrained(spec.id, { dtype, device: "webgpu", progress_callback });
  }
  const tReady = now();

  let bytes = 0;
  let downloadEnd = 0;
  for (const v of files.values()) {
    bytes += v.total || v.loaded;
    downloadEnd = Math.max(downloadEnd, v.end);
  }
  // Czy wagi zostały w pamięci przeglądarki? Jeśli nie, każde otwarcie pobiera model od nowa.
  let cachedBytes = 0;
  try {
    const cache = await caches.open(CACHE_KEY);
    const urls = (await cache.keys()).map((r) => r.url);
    for (const [file, v] of files) {
      if (urls.some((u) => u.includes(spec.id) && u.endsWith("/" + file))) cachedBytes += v.total || v.loaded;
    }
  } catch {}
  const downloadMs = firstByteAt && downloadEnd ? downloadEnd - t0 : 0;
  return {
    model,
    tokenizer,
    stats: {
      bytes,
      cachedBytes,
      loadMs: tReady - t0,
      downloadMs,
      initMs: downloadEnd ? tReady - downloadEnd : tReady - t0,
    },
  };
}

// ---------- generowanie z pomiarem czasu ----------

async function generate(model, tokenizer, messages, maxNew) {
  const L = lib;
  const text = tokenizer.apply_chat_template(messages, {
    tokenize: false,
    add_generation_prompt: true,
    enable_thinking: false,
  });
  const inputs = tokenizer(text, { add_special_tokens: false });
  const promptTokens = inputs.input_ids.dims.at(-1);
  const stamps = [];
  const streamer = new L.TextStreamer(tokenizer, {
    skip_prompt: true,
    skip_special_tokens: true,
    callback_function: () => {},
    token_callback_function: () => stamps.push(now()),
  });
  const t0 = now();
  const out = await model.generate({ ...inputs, max_new_tokens: maxNew, do_sample: false, streamer });
  const tEnd = now();
  const seq = out.slice ? out.slice(null, [promptTokens, null]) : out;
  const decoded = tokenizer.batch_decode(seq, { skip_special_tokens: true })[0] || "";
  const newTokens = stamps.length;
  const ttft = stamps.length ? stamps[0] - t0 : tEnd - t0;
  const decodeTps = newTokens > 1 ? (newTokens - 1) / ((stamps.at(-1) - stamps[0]) / 1000) : 0;
  return {
    text: decoded.replace(/<think>[\s\S]*?<\/think>/g, "").trim(),
    promptTokens,
    newTokens,
    ttftMs: ttft,
    totalMs: tEnd - t0,
    prefillTps: promptTokens / (ttft / 1000),
    decodeTps,
  };
}

function msgs(system, user) {
  return [
    { role: "system", content: system },
    { role: "user", content: user },
  ];
}

// ---------- przebieg ----------

// Błąd w promise biblioteki (np. zerwane pobieranie dużego pliku) nie dochodzi do try/catch niżej
// i worker wisiałby w nieskończoność — zgłaszamy go stronie.
self.addEventListener("unhandledrejection", (e) => {
  post("fatal", { error: String(e.reason?.message || e.reason).slice(0, 500) });
});

self.onmessage = async (e) => {
  const { spec, dtype } = e.data;
  const result = { id: spec.id, dtype, startedAt: new Date().toISOString() };
  try {
    post("phase", { phase: "pobieranie" });
    const { model, tokenizer, stats } = await loadModel(spec, dtype);
    Object.assign(result, stats);

    post("phase", { phase: "rozgrzewka" });
    const w = await generate(model, tokenizer, msgs("Odpowiadasz krótko.", "Cześć!"), 4);
    result.warmupMs = w.totalMs;

    post("phase", { phase: "szybkość" });
    const s = await generate(model, tokenizer, msgs("Odpowiadasz po polsku.", SPEED_PROMPT), 96);
    result.speed = {
      promptTokens: s.promptTokens,
      newTokens: s.newTokens,
      ttftMs: s.ttftMs,
      prefillTps: s.prefillTps,
      decodeTps: s.decodeTps,
      sample: s.text.slice(0, 600),
    };

    result.tasks = [];
    for (const task of TASKS) {
      post("phase", { phase: `test: ${task.label}` });
      const g = await generate(model, tokenizer, msgs(task.system, task.user), task.max);
      const verdict = task.check(g.text);
      result.tasks.push({ id: task.id, label: task.label, ok: verdict.ok, note: verdict.note, output: g.text.slice(0, 400), ms: g.totalMs });
    }
    result.score = result.tasks.filter((t) => t.ok).length;
    result.status = "ok";
    try {
      await model.dispose();
    } catch {}
  } catch (err) {
    result.status = "błąd";
    result.error = String(err?.message || err).slice(0, 500);
  }
  result.finishedAt = new Date().toISOString();
  post("result", { result });
};
