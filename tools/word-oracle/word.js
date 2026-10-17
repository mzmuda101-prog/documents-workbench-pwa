// word.js — odpytywanie prawdziwego Worda (Mac, AppleScript) o to, co pokazuje dla pliku.
// Zawsze ten sam plik wejściowy (output/word-oracle/_in.docx): piaskownica Worda pyta o zgodę
// tylko raz na ścieżkę. Tylko testy/narzędzia — nie w dist/.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const DIR = path.resolve(__dirname, "../../output/word-oracle");
fs.mkdirSync(DIR, { recursive: true });
const IN = path.join(DIR, "_in.docx");
const esc = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');

function osa(script, timeout = 240000) {
  return execFileSync("osascript", ["-e", script], { timeout, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }).replace(/\n$/, "");
}

function closeAll() {
  try { osa('tell application "Microsoft Word" to close every document saving no', 30000); } catch (_) { /* nic */ }
}

// body = AppleScript wewnątrz `tell active document`; zwraca jego wynik (tekst)
function withDocument(bytes, body) {
  fs.writeFileSync(IN, bytes);
  try {
    return osa(`
      tell application "Microsoft Word"
        open file name "${esc(IN)}"
        set res to ""
        tell active document
          ${body}
        end tell
        try
          close active document saving no
        end try
        return res
      end tell`);
  } catch (e) {
    closeAll();
    throw new Error(String(e.stderr || e.message).trim().split("\n").pop());
  }
}

// numery list i tekst każdego akapitu: [{ label, text }]
const SEP = "\u0001";
function listStrings(bytes) {
  const out = withDocument(bytes, `
    set n to count of paragraphs
    repeat with i from 1 to n
      set pp to paragraph i
      set ls to ""
      try
        set ls to list string of list format of text object of pp
      end try
      if ls is missing value then set ls to ""
      set tx to content of text object of pp
      set res to res & ls & "${SEP}" & tx & linefeed
    end repeat`);
  return out.split("\n").filter(Boolean).map((l) => {
    const [label, text] = l.split(SEP);
    return { label, text: (text || "").replace(/[\r\u0007]+$/, "") };
  });
}

module.exports = { osa, withDocument, listStrings, closeAll, DIR };

// statystyka wyrazów Worda (jak okno „Statystyka wyrazów”, bez przypisów)
function statistics(bytes, notes = false) {
  const out = withDocument(bytes, `
    set res to ((compute statistics statistic statistic words include footnotes and endnotes ${notes}) as text) & "|" & ((compute statistics statistic statistic characters include footnotes and endnotes ${notes}) as text) & "|" & ((compute statistics statistic statistic characters with spaces include footnotes and endnotes ${notes}) as text) & "|" & ((compute statistics statistic statistic paragraphs include footnotes and endnotes ${notes}) as text) & "|" & ((compute statistics statistic statistic lines include footnotes and endnotes ${notes}) as text) & "|" & ((compute statistics statistic statistic pages include footnotes and endnotes ${notes}) as text)`);
  const [words, chars, charsWithSpaces, paragraphs, lines, pages] = out.split("|").map(Number);
  return { words, chars, charsWithSpaces, paragraphs, lines, pages };
}
module.exports.statistics = statistics;

// zmiana wielkości liter w Wordzie: [{ para (1…), from, to (znaki w akapicie, null = cały), mode }] → teksty akapitów
const CASE_ENUM = { lower: "lower case", upper: "upper case", title: "title word", sentence: "title sentence", toggle: "toggle case" };
function changeCase(bytes, ops) {
  const body = ops.map((o, i) => `
    set pr to text object of paragraph ${o.para}
    set ps to start of content of pr
    set pe to (end of content of pr) - 1
    set r${i} to create range it start (ps + ${o.from || 0}) end ${o.to == null ? "pe" : `(ps + ${o.to})`}
    set case of r${i} to ${CASE_ENUM[o.mode]}
    set res to res & (content of text object of paragraph ${o.para}) & "${SEP}"`).join("\n");
  return withDocument(bytes, body).split(SEP).slice(0, ops.length).map((s) => s.replace(/[\r\u0007]+$/, ""));
}
module.exports.changeCase = changeCase;
