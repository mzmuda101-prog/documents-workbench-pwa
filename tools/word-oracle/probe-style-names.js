// probe-style-names.js — polskie nazwy stylów wbudowanych Worda (w pliku zawsze angielskie w:name)
const { buildDocx } = require("./build-docx");
const word = require("./word");
const NAMES = ["Quote", "Intense Quote", "Block Text", "Body Text", "Body Text 2", "Body Text 3", "Body Text Indent", "Body Text Indent 2", "Body Text Indent 3", "Body Text First Indent", "Body Text First Indent 2", "No Spacing", "caption", "footnote text", "endnote text", "header", "footer", "Normal (Web)", "Plain Text", "Signature", "Salutation", "Closing", "Date", "List", "List 2", "List Bullet", "List Bullet 2", "List Number", "List Number 2", "List Continue", "List Continue 2", "heading 4", "heading 5", "heading 6", "heading 7", "heading 8", "heading 9", "Title", "Subtitle", "Balloon Text", "annotation text", "annotation subject", "envelope address", "envelope return", "HTML Preformatted", "HTML Address", "index 1", "index heading", "table of figures", "table of authorities", "toa heading", "macro", "Message Header", "Note Heading", "E-mail Signature", "Normal Indent", "Bibliography", "TOC Heading", "toc 4", "Document Map", "Revision", "List Paragraph", "Intense Emphasis", "Emphasis", "Strong", "Subtle Emphasis", "Book Title", "Subtle Reference", "Intense Reference", "Hyperlink", "Code"];
(async () => {
  const styles = NAMES.map((n, i) => `<w:style w:type="paragraph" w:styleId="S${i}"><w:name w:val="${n}"/><w:basedOn w:val="Normal"/></w:style>`).join("");
  const bytes = await buildDocx({ styles, body: NAMES.map((n, i) => ({ text: n, style: `S${i}` })) });
  const out = word.withDocument(bytes, `
    repeat with i from 1 to count of paragraphs
      set res to res & (name local of (get style of paragraph i)) & "|"
    end repeat`);
  const pl = out.split("|");
  const map = {};
  NAMES.forEach((n, i) => { if (pl[i] && pl[i] !== n) map[n.toLowerCase()] = pl[i]; });
  console.log(JSON.stringify(map, null, 1));
})();
