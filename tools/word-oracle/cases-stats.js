// Statystyka wyrazów: przypadki tekstu (każdy = osobny dokument)
const { p } = require("./build-docx");
const tbl = (rows) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${rows[0].map(() => '<w:gridCol w:w="3000"/>').join("")}</w:tblGrid>${rows.map((r) => `<w:tr>${r.map((c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(c)}</w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
const raw = (xml) => ({ raw: xml });
module.exports = [
  { name: "proste", body: ["Ala ma kota.", "Kot ma Alę."] },
  { name: "myslniki", body: ["To jest — myślnik i - dywiz oraz – półpauza.", "a—b c–d e-f"] },
  { name: "liczby", body: ["e-mail, 3,5 kg, 10 000 zł, 1/2, x+y=z, 50%"] },
  { name: "nbsp", body: ["10 000 zł i tak", "a b"] },
  { name: "interpunkcja-sama", body: ["Słowo , słowo . ! ? ... … ( ) „ ” « »", "* • · # @ & § | /"] },
  { name: "nawiasy-cudzyslowy", body: ["„Cytat” (nawias) [kw] {kl} 'pojedyncze' \"proste\""] },
  { name: "ukosniki", body: ["i/lub słowo\\słowo http://example.com/a?b=c user@example.com"] },
  { name: "emoji", body: ["Hej 👋 świecie 🌍🙂 ok", "😀"] },
  { name: "puste-akapity", body: ["Jeden", "", "", "Dwa", " ", "\t"] },
  { name: "tabulatory", body: [raw('<w:r><w:t>a</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>b</w:t></w:r><w:r><w:tab/><w:t>c</w:t></w:r>')] },
  { name: "lamanie-wiersza", body: [raw('<w:r><w:t>raz</w:t><w:br/><w:t>dwa</w:t></w:r>')] },
  { name: "tabela", body: ["Przed tabelą.", tbl([["a b", "c"], ["d e f", ""]]), "Po tabeli."] },
  { name: "lista-numerowana", abstractNums: [{ levels: [{ fmt: "decimal", text: "%1." }] }], nums: [{ id: 1, abs: 0 }], body: [{ text: "pierwszy punkt", numId: 1 }, { text: "drugi", numId: 1 }] },
  { name: "ukryty-tekst", body: [raw('<w:r><w:t xml:space="preserve">widać </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>ukryte słowa</w:t></w:r>')] },
  { name: "skreslony-sledzony", body: [raw('<w:r><w:t xml:space="preserve">jest </w:t></w:r><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>usunięte słowo</w:delText></w:r></w:del><w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve"> wstawione</w:t></w:r></w:ins>')] },
  { name: "spacje-bez-preserve", body: [raw('<w:r><w:rPr><w:b/></w:rPr><w:t>Red bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t> italic underline</w:t></w:r><w:r><w:t> mono</w:t></w:r>'), raw('<w:r><w:t>A  b</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t>c\u00a0</w:t></w:r><w:r><w:t>d</w:t></w:r>'), raw('<w:r><w:t xml:space="preserve">z </w:t></w:r><w:r><w:t>preserve</w:t></w:r>')] },
  { name: "dlugi", body: Array.from({ length: 30 }, (_, i) => `Zdanie numer ${i + 1} ma kilka słów, przecinek; średnik: dwukropek — i koniec.`) },
];
