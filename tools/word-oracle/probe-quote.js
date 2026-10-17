// probe-quote.js — cytaty i ramki: jaką nazwę stylu pokazuje Word i co robi gumka / styl Normalny
const { buildDocx } = require("./build-docx");
const word = require("./word");
const { QUOTE_STYLES, QUOTE_BODY } = require("./cases-quote");
(async () => {
  const bytes = await buildDocx({ styles: QUOTE_STYLES, body: QUOTE_BODY });
  const names = word.withDocument(bytes, `
    repeat with i from 1 to count of paragraphs
      set res to res & (name local of (get style of paragraph i)) & "|"
    end repeat`);
  console.log(names);
  const r = word.saveDocx(bytes, `
    select text object of paragraph 4
    clear formatting (get selection of application "Microsoft Word")
    set style of paragraph 5 to style normal
    select (create range it start (start of content of text object of paragraph 6) end (start of content of text object of paragraph 6))
    clear formatting (get selection of application "Microsoft Word")`);
  console.log(r.parts["word/document.xml"].match(/<w:body>[\s\S]*<\/w:body>/)[0].replace(/<w:p[ >]/g, "\n$&"));
})();
