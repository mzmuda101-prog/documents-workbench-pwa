// keys-table.js — TABELA ZACHOWAŃ KLAWISZY: rodzaj akapitu × klawisz → oczekiwany wynik (jak w Wordzie).
// Uruchamia ją scripts/keys-table-playwright.js. Jedno miejsce, w którym widać całą specyfikację —
// brakujące pole tabeli to błąd, którego nikt nie wymyślił (2026-11-08: Backspace w pustej linijce
// pod linią poziomą usuwał linię — nikt nie zapisał, co ma się stać).
//
// Zapis stanu: akapity rozdzielone „ ¶ ”, każdy `rodzaj:tekst` (zwykły — sam tekst), pusty = ∅,
// na końcu kursor `@akapit:znak`. Rodzaje: list (punkt listy), ind (wcięty), h1, quote (Cytat),
// rule (linia pozioma = pusty akapit z dolną krawędzią), line (tekst z dolną krawędzią).
// Klawisze: bs = Backspace na początku, del = Delete na końcu, enterStart / enterEnd = Enter na
// początku / końcu (w pustym akapicie: enter). „?” = jeszcze bez decyzji — raport pokazuje, co
// aplikacja robi dziś; po decyzji wpisać wynik.
//
// Zasady Worda wspólne dla tabeli:
//  - sklejenie: wynik ma formatowanie GÓRNEGO akapitu; górny pusty — znika, dolny zostaje ze swoim
//  - Enter na końcu nagłówka → następny akapit Normalny; w środku/na początku — ten sam styl
//  - Backspace na początku punktu listy / wciętego / pustej ramki najpierw zdejmuje tę cechę

module.exports = {
  contexts: {
    normal: { start: "Alfa ¶ Beta ¶ Gamma", at: 1 },
    empty: { start: "Alfa ¶ ∅ ¶ Gamma", at: 1 },
    list: { start: "Alfa ¶ list:Beta ¶ Gamma", at: 1 },
    listEmpty: { start: "Alfa ¶ list:∅ ¶ Gamma", at: 1 },
    ind: { start: "Alfa ¶ ind:Beta ¶ Gamma", at: 1 },
    h1: { start: "Alfa ¶ h1:Beta ¶ Gamma", at: 1 },
    quote: { start: "Alfa ¶ quote:Beta ¶ Gamma", at: 1 },
    quoteEmpty: { start: "Alfa ¶ quote:∅ ¶ Gamma", at: 1 },
    rule: { start: "Alfa ¶ rule:∅ ¶ ∅ ¶ Gamma", at: 1 },
    underRule: { start: "Alfa ¶ rule:∅ ¶ ∅ ¶ Gamma", at: 2 },
    line: { start: "Alfa ¶ line:Beta ¶ ∅ ¶ Gamma", at: 1 },
  },
  expect: {
    normal: {
      bs: "AlfaBeta ¶ Gamma @0:4",
      del: "Alfa ¶ BetaGamma @1:4",
      enterStart: "Alfa ¶ ∅ ¶ Beta ¶ Gamma @2:0",
      enterEnd: "Alfa ¶ Beta ¶ ∅ ¶ Gamma @2:0",
    },
    empty: {
      bs: "Alfa ¶ Gamma @0:4",
      del: "Alfa ¶ Gamma @1:0",
      enter: "Alfa ¶ ∅ ¶ ∅ ¶ Gamma @2:0",
    },
    list: {
      bs: "Alfa ¶ ind:Beta ¶ Gamma @1:0", // punktor znika, wcięcie tekstu zostaje
      del: "Alfa ¶ list:BetaGamma @1:4",
      enterStart: "Alfa ¶ list:∅ ¶ list:Beta ¶ Gamma @2:0",
      enterEnd: "Alfa ¶ list:Beta ¶ list:∅ ¶ Gamma @2:0",
    },
    listEmpty: {
      bs: "Alfa ¶ ind:∅ ¶ Gamma @1:0",
      del: "Alfa ¶ Gamma @1:0",
      enter: "Alfa ¶ ∅ ¶ Gamma @1:0", // koniec listy, od razu do marginesu
    },
    ind: {
      bs: "Alfa ¶ Beta ¶ Gamma @1:0", // wcięcie do zera, akapit zostaje
      del: "Alfa ¶ ind:BetaGamma @1:4",
      enterStart: "Alfa ¶ ind:∅ ¶ ind:Beta ¶ Gamma @2:0",
      enterEnd: "Alfa ¶ ind:Beta ¶ ind:∅ ¶ Gamma @2:0",
    },
    h1: {
      bs: "AlfaBeta ¶ Gamma @0:4",
      del: "Alfa ¶ h1:BetaGamma @1:4",
      enterStart: "Alfa ¶ h1:∅ ¶ h1:Beta ¶ Gamma @2:0",
      enterEnd: "Alfa ¶ h1:Beta ¶ ∅ ¶ Gamma @2:0",
    },
    quote: {
      bs: "AlfaBeta ¶ Gamma @0:4", // jak Word: od razu skleja (wcięcie ze stylu to wygląd Cytatu)
      del: "Alfa ¶ quote:BetaGamma @1:4",
      enterStart: "Alfa ¶ quote:∅ ¶ quote:Beta ¶ Gamma @2:0",
      enterEnd: "Alfa ¶ quote:Beta ¶ ∅ ¶ Gamma @2:0", // Word: Cytat ma „następny styl: Normalny”
    },
    quoteEmpty: {
      bs: "Alfa ¶ ∅ ¶ Gamma @1:0", // najpierw zwykły tekst, kolejny Backspace skleja
      del: "Alfa ¶ Gamma @1:0",
      enter: "Alfa ¶ ∅ ¶ Gamma @1:0", // wyjście z ramki jak z listy
    },
    rule: {
      bs: "Alfa ¶ ∅ ¶ Gamma @0:4", // linia wsiąka w akapit wyżej
      del: "Alfa ¶ rule:∅ ¶ Gamma @1:0", // pusta linijka pod linią znika, linia zostaje
      enter: "Alfa ¶ rule:∅ ¶ ∅ ¶ ∅ ¶ Gamma @2:0", // nowy akapit bez linii
    },
    underRule: {
      bs: "Alfa ¶ rule:∅ ¶ Gamma @1:0", // znika pusta linijka, nie linia (zgłoszenie 2026-11-08)
      del: "Alfa ¶ rule:∅ ¶ Gamma @2:0",
      enter: "Alfa ¶ rule:∅ ¶ ∅ ¶ ∅ ¶ Gamma @3:0",
    },
    line: {
      bs: "AlfaBeta ¶ ∅ ¶ Gamma @0:4", // jak w Wordzie: skleja, linię usuwa menu (decyzja 2026-11-08)
      del: "Alfa ¶ line:Beta ¶ Gamma @1:4",
      // Word: oba akapity z krawędzią, rysowane jako grupa z JEDNĄ linią na dole — u nas linia tylko
      // przy części z tekstem (wygląd jak w Wordzie): na początku — zostaje pod tekstem, na końcu —
      // nowa linijka bez linii (decyzja 2026-11-08)
      enterStart: "Alfa ¶ ∅ ¶ line:Beta ¶ ∅ ¶ Gamma @2:0",
      enterEnd: "Alfa ¶ line:Beta ¶ ∅ ¶ ∅ ¶ Gamma @2:0",
    },
  },
};
