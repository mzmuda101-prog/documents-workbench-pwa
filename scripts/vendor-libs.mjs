#!/usr/bin/env node
/** Bundle docx-preview for browser (global docx.renderAsync). */
import esbuild from "esbuild";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "lib", "docx-preview.bundle.js");
const ENTRY = path.join(ROOT, "node_modules", "docx-preview", "dist", "docx-preview.js");

async function main() {
  if (!fs.existsSync(ENTRY)) {
    console.warn("⚠️  docx-preview not installed — run npm install first");
    return;
  }
  await esbuild.build({
    entryPoints: [ENTRY],
    outfile: OUT,
    bundle: true,
    format: "iife",
    globalName: "docx",
    minify: true,
    legalComments: "none",
    target: "es2020",
  });
  patchBundle();
  const kb = (fs.statSync(OUT).size / 1024).toFixed(1);
  console.log(`✅  lib/docx-preview.bundle.js (${kb} KB)`);
}

// Łatki na zbudowanej bibliotece (nazwy po minifikacji się zmieniają — szukamy wzorcem).
// 1) „Od nowej strony” (w:pageBreakBefore) docx-preview bierze tylko ze STYLU akapitu;
//    bezpośrednie ustawienie akapitu (Ctrl+Enter w Documents Workbench, Word „Podział
//    wiersza i strony → Podział strony przed”) nie łamało strony w podglądzie.
function patchBundle() {
  let src = fs.readFileSync(OUT, "utf8");
  const re = /(\w+)\.type==(\w+)\.Paragraph&&this\.findStyle\(\1\.styleName\)\?\.paragraphProps\?\.pageBreakBefore/;
  if (!re.test(src)) throw new Error("łatka pageBreakBefore: wzorzec nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(re, (_, p, T) => `${p}.type==${T}.Paragraph&&(${p}.pageBreakBefore||this.findStyle(${p}.styleName)?.paragraphProps?.pageBreakBefore)`);
  console.log("  ✅  łatka: pageBreakBefore akapitu");

  // 2) Obraz „pływający” (wp:anchor, wrapNone) położony względem STRONY: docx-preview stawiał go
  //    względem akapitu, w którym siedzi kotwica, i ignorował „za tekstem”. PDF → DOCX kotwiczy
  //    mapki/skany/warstwę grafiki do strony — bez tej łatki lądowały w złym miejscu.
  const reBehind = /(\w+)\.boolAttr\((\w+),"behindDoc"\);let (\w+)=\{relative:"page",align:"left",offset:"0"\},(\w+)=\{relative:"page",align:"top",offset:"0"\};/;
  const mB = src.match(reBehind);
  if (!mB) throw new Error("łatka anchor: wzorzec behindDoc nie pasuje (nowa wersja docx-preview?)");
  const [, P, NODE, PX, PY] = mB;
  src = src.replace(reBehind, `var __dwbBehind=${P}.boolAttr(${NODE},"behindDoc");let ${PX}={relative:"page",align:"left",offset:"0"},${PY}={relative:"page",align:"top",offset:"0"};`);
  const reWrap = /(\w+)=="wrapNone"\?\((\w+)\.cssStyle\.display="block",\2\.cssStyle\.position="relative",/;
  if (!reWrap.test(src)) throw new Error("łatka anchor: wzorzec wrapNone nie pasuje");
  src = src.replace(reWrap, (_, O, E) => `${O}=="wrapNone"?(${E}.cssStyle.display="block",${E}.cssStyle.position=(${PX}.relative=="page"&&${PY}.relative=="page"?"absolute":"relative"),${PX}.relative=="page"&&${PY}.relative=="page"&&(${E}.cssStyle["z-index"]=__dwbBehind?"-1":"1",${E}.cssStyle["pointer-events"]="none"),`);
  console.log("  ✅  łatka: obraz zakotwiczony do strony (absolute, za tekstem)");

  // 3) Tabulatory: docx-preview rysuje każdy jako stałą spację (&emsp;), a rozkład wg pozycji
  //    tabulatorów ma tylko w trybie „experimental” (po 500 ms, na współrzędnych ekranu — psuje
  //    się przy powiększeniu). Zostawiamy na znaczniku pozycje z akapitu; rozkład robi
  //    app/docx-render-fixes.js (layoutTabStops) od razu po renderze.
  const reTab = /renderTab\((\w+)\)\{var (\w+)=this\.createElement\("span"\);if\(\2\.innerHTML="&emsp;",this\.options\.experimental\)\{\2\.className=this\.tabStopClass\(\);var (\w+)=([\w$]+)\(\1,([\w$]+)\.Paragraph\)\?\.tabs;/;
  if (!reTab.test(src)) throw new Error("łatka tabulatorów: wzorzec renderTab nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reTab, (m, T, E, R, F, D) => `renderTab(${T}){var ${E}=this.createElement("span");${E}.className="docx-tab";try{var __dwbStops=${F}(${T},${D}.Paragraph)?.tabs;__dwbStops&&__dwbStops.length&&(${E}.dataset.stops=JSON.stringify(__dwbStops.map(s=>[parseFloat(s.position),s.leader||"none",s.style||"left"])));this.defaultTabSize&&(${E}.dataset.dt=parseFloat(this.defaultTabSize))}catch(__e){}if(${E}.innerHTML="&emsp;",this.options.experimental){${E}.className+=" "+this.tabStopClass();var ${R}=${F}(${T},${D}.Paragraph)?.tabs;`);
  console.log("  ✅  łatka: tabulatory z pozycjami (dane dla layoutTabStops)");

  // 4) Odstępy między znakami w przebiegu (w:rPr/w:spacing — Word „Czcionka → Zaawansowane →
  //    Odstępy: zagęszczone/rozstrzelone”): docx-preview czytał spacing tylko z akapitu.
  //    PDF → DOCX zagęszcza tak tekst, gdy krój z PDF zastępuje Arial/Times (szerszy).
  const reSp = /case"spacing":(\w+)\.localName=="pPr"&&this\.parseSpacing\((\w+),(\w+)\);break;/;
  const mSp = src.match(reSp);
  if (!mSp) throw new Error("łatka spacing przebiegu: wzorzec nie pasuje (nowa wersja docx-preview?)");
  const XML = (src.slice(Math.max(0, mSp.index - 4000), mSp.index).match(/(\w+)\.lengthAttr\(\w+,"val",\w+\.FontSize\)/) || [])[1];
  if (!XML) throw new Error("łatka spacing przebiegu: brak lengthAttr w pobliżu");
  src = src.replace(reSp, (_, T, K, E) => `case"spacing":${T}.localName=="pPr"?this.parseSpacing(${K},${E}):${T}.localName=="rPr"&&(${E}["letter-spacing"]=${XML}.lengthAttr(${K},"val"));break;`);
  console.log("  ✅  łatka: odstępy między znakami w przebiegu (letter-spacing)");

  // 5) Przypisy: docx-preview numerował od 1 na KAŻDEJ stronie podglądu (Word — ciągle przez
  //    dokument) i nie zostawiał śladu, który przypis wskazuje odnośnik. Numer ciągły + na
  //    odnośniku i pozycji listy „footnote:ID” / „endnote:ID” — app/doc-notes.js robi z tego
  //    dymek z treścią, skok do przypisu, edycję tekstu przypisu i format numeru z ustawień.
  for (const kind of ["Footnote", "Endnote"]) {
    const reRef = new RegExp(`render${kind}Reference\\((\\w+)\\)\\{var (\\w+)=this\\.createElement\\("sup"\\);return this\\.current${kind}Ids\\.push\\(\\1\\.id\\),\\2\\.textContent=\`\\$\\{this\\.current${kind}Ids\\.length\\}\``);
    if (!reRef.test(src)) throw new Error(`łatka przypisów: wzorzec render${kind}Reference nie pasuje (nowa wersja docx-preview?)`);
    const low = kind.toLowerCase();
    src = src.replace(reRef, (_, el, sup) => `render${kind}Reference(${el}){var ${sup}=this.createElement("sup");return this.current${kind}Ids.push(${el}.id),this.__dwb${kind}N=(this.__dwb${kind}N||0)+1,${sup}.dataset.dwbNote="${low}:"+${el}.id,${sup}.dataset.dwbNoteNum=this.__dwb${kind}N,${sup}.textContent=\`\${this.__dwb${kind}N}\``);
  }
  const reNotes = /renderNotes\((\w+),(\w+),(\w+)\)\{var (\w+)=\1\.map\((\w+)=>\2\[\5\]\)\.filter\(\5=>\5\);if\(\4\.length>0\)\{var (\w+)=this\.createElement\("ol",null,this\.renderElements\(\4\)\);\3\.appendChild\(\6\)\}/;
  if (!reNotes.test(src)) throw new Error("łatka przypisów: wzorzec renderNotes nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reNotes, (_, ids, map, into, notes, x, ol) => `renderNotes(${ids},${map},${into}){var ${notes}=${ids}.map(${x}=>${map}[${x}]).filter(${x}=>${x});if(${notes}.length>0){var ${ol}=this.createElement("ol",null,this.renderElements(${notes}));var __k=${map}===this.footnoteMap?"footnote":"endnote";${ol}.className="dwb-notes dwb-notes-"+__k;${notes}.forEach((__n,__i)=>{var __li=${ol}.children[__i];__li&&(__li.dataset.dwbNote=__k+":"+__n.id)});${into}.appendChild(${ol})}`);
  console.log("  ✅  łatka: przypisy — numeracja ciągła, odnośnik i pozycja listy z id przypisu");

  // 6) Podział strony W ŚRODKU akapitu (tekst za <w:br w:type="page"/>): docx-preview dzieli
  //    akapit na dwa <p>. Dawniej akapity podglądu rozjeżdżały się od tego miejsca z plikiem
  //    (edycja dalszych trafiała w złe akapity albo przepadała). Druga połówka dostaje
  //    data-dwb-cont — collectPreviewParagraphElements ją pomija, cały akapit tylko do odczytu.
  const reSplit = /var (\w+)=(\w+)\.children,(\w+)=\{\.\.\.\2,children:\1\.slice\((\w+)\)\};/;
  if (!reSplit.test(src)) throw new Error("łatka podziału akapitu: wzorzec splitBySection nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reSplit, (_, kids, p, np, idx) => `var ${kids}=${p}.children,${np}={...${p},children:${kids}.slice(${idx}),__dwbCont:!0};`);
  const reRenderP = /renderParagraph\((\w+)\)\{var (\w+)=this\.renderContainer\(\1,"p"\);/;
  if (!reRenderP.test(src)) throw new Error("łatka podziału akapitu: wzorzec renderParagraph nie pasuje");
  src = src.replace(reRenderP, (_, el, p) => `renderParagraph(${el}){var ${p}=this.renderContainer(${el},"p");${el}.__dwbCont&&(${p}.dataset.dwbCont="1");`);
  console.log("  ✅  łatka: druga połówka akapitu podzielonego podziałem strony oznaczona");

  // 7) Wcięcie tabeli (w:tblInd w:w="…" w:type="dxa"): docx-preview czytał je jak wcięcie akapitu
  //    (atrybuty left/start) — tabela zawsze stała przy marginesie. W Wordzie (tryb 2013+) to
  //    odległość krawędzi tabeli od marginesu, także ujemna (tabela wysunięta w lewo).
  const reTblInd = /case"ind":case"tblInd":this\.parseIndentation\((\w+),(\w+)\);break;/;
  if (!reTblInd.test(src)) throw new Error("łatka tblInd: wzorzec nie pasuje (nowa wersja docx-preview?)");
  const mParser = src.match(/parseIndentation\(\w+,\w+\)\{var \w+=(\w+)\.lengthAttr\(/);
  if (!mParser) throw new Error("łatka tblInd: brak globalXmlParser w parseIndentation");
  const XP = mParser[1];
  src = src.replace(reTblInd, (_, el, st) => `case"ind":this.parseIndentation(${el},${st});break;case"tblInd":{var __ty=${XP}.attr(${el},"type");if(!__ty||__ty==="dxa"){var __w=${XP}.lengthAttr(${el},"w");__w&&(${st}["margin-inline-start"]=__w)}}break;`);
  console.log("  ✅  łatka: wcięcie tabeli (w:tblInd)");

  // 8) Wyrównanie strony w pionie (sectPr w:vAlign: top / center / both / bottom — Word:
  //    Ustawienia strony → Układ). docx-preview go nie czytał; strona dostaje data-dwb-v-align,
  //    resztę robi CSS (styles/app.css).
  const reTitlePg = /case"titlePg":(\w+)\.titlePage=(\w+)\.boolAttr\((\w+),"val",!0\);break;/;
  if (!reTitlePg.test(src)) throw new Error("łatka vAlign strony: wzorzec titlePg nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reTitlePg, (m, sec, xml, el) => `${m}case"vAlign":${sec}.vAlign=${xml}.attr(${el},"val");break;`);
  const rePage = /createPageElement\((\w+),(\w+)\)\{var (\w+)=this\.createElement\("section",\{className:\1\}\);/;
  if (!rePage.test(src)) throw new Error("łatka vAlign strony: wzorzec createPageElement nie pasuje");
  src = src.replace(rePage, (m, cls, props, el) => `${m}${props}&&${props}.vAlign&&${props}.vAlign!=="top"&&(${el}.dataset.dwbVAlign=${props}.vAlign);`);
  console.log("  ✅  łatka: wyrównanie strony w pionie (sectPr w:vAlign)");

  // 9) Listy dzielące jedną definicję (w:abstractNum) — Word tworzy nowy w:num przy każdym
  //    „zacznij od nowa” i przy wklejaniu list. docx-preview mapował abstractNum → JEDEN numId
  //    (ostatni), więc tylko ostatnia lista dostawała style (kropka/numer, wcięcie); pozostałe
  //    stały przy lewym marginesie bez znaczników (wykryte porównaniem z PDF-em z Worda,
  //    npm run word:compare, 2026-10-05: wykład Mateusza — 16 z 17 list bez kropek).
  //    Teraz: poziomy definicji dla KAŻDEGO w:num; w:lvlOverride (startOverride / własny w:lvl)
  //    uwzględnione. Licznik: lista z nadpisaniem liczy od nowa (własny licznik), listy bez
  //    nadpisania jednej definicji dzielą licznik — numeracja idzie dalej, jak w Wordzie.
  const reNum = /parseNumberingFile\((\w+)\)\{var (\w+)=\[\],(\w+)=\{\},(\w+)=\[\];for\(let (\w+) of (\w+)\.elements\(\1\)\)switch\(\5\.localName\)\{case"abstractNum":[\s\S]*?return \2\.forEach\(\w+=>\w+\.id=\3\[\w+\.id\]\),\2\}/;
  const mNum = src.match(reNum);
  if (!mNum) throw new Error("łatka list: wzorzec parseNumberingFile nie pasuje (nowa wersja docx-preview?)");
  const X = mNum[6];
  src = src.replace(reNum, (_, T) => `parseNumberingFile(${T}){var __abs={},__nums=[],__bul=[],__out=[];for(let __e of ${X}.elements(${T}))switch(__e.localName){case"abstractNum":__abs[${X}.attr(__e,"abstractNumId")]=this.parseAbstractNumbering(__e,__bul);break;case"numPicBullet":__bul.push(this.parseNumberingPicBullet(__e));break;case"num":{let __ov={},__has=!1;for(let __l of ${X}.elements(__e))if(__l.localName=="lvlOverride"){let __s=${X}.element(__l,"startOverride"),__v=${X}.element(__l,"lvl");__ov[${X}.intAttr(__l,"ilvl")]={start:__s?${X}.intAttr(__s,"val"):null,lvl:__v};__has=!0}__nums.push({id:${X}.attr(__e,"numId"),abs:${X}.elementAttr(__e,"abstractNumId","val"),ov:__ov,has:__has});break}}for(let __n of __nums)for(let __x of __abs[__n.abs]||[]){let __o=__n.ov[__x.level],__c=__o&&__o.lvl?this.parseNumberingLevel(__n.id,__o.lvl,__bul):{...__x,id:__n.id};__o&&__o.start!=null&&(__c.start=__o.start);__c.cid=__n.has?__n.id:"a"+__n.abs;__out.push(__c)}return __out}`);
  const reCnt = /let (\w+)=this\.numberingCounter\((\w+)\.id,\2\.level\),(\w+)=\1\+" "\+\(\2\.start-1\);/;
  if (!reCnt.test(src)) throw new Error("łatka list: wzorzec numberingCounter w renderNumbering nie pasuje");
  src = src.replace(reCnt, (_, J, K, KT) => `let ${J}=this.numberingCounter(${K}.cid??${K}.id,${K}.level),${KT}=${J}+" "+(${K}.start-1);`);
  const reLvlTxt = /this\.levelTextToContent\((\w+)\.levelText,\1\.suff,\1\.id,/;
  if (!reLvlTxt.test(src)) throw new Error("łatka list: wzorzec levelTextToContent nie pasuje");
  src = src.replace(reLvlTxt, (_, K) => `this.levelTextToContent(${K}.levelText,${K}.suff,${K}.cid??${K}.id,`);
  const rePush = /(\w+)\.push\((\w+)\),(\w+)\+=this\.styleToString\(`\$\{(\w+)\}:before`,\{content:this\.levelTextToContent/;
  if (!rePush.test(src)) throw new Error("łatka list: wzorzec counter-reset nie pasuje");
  src = src.replace(rePush, (_, O, KT, R, J) => `${O}.includes(${KT})||${O}.push(${KT}),${R}+=this.styleToString(\`\${${J}}:before\`,{content:this.levelTextToContent`);
  console.log("  ✅  łatka: listy dzielące definicję (każdy w:num ze stylami, lvlOverride, wspólny licznik)");

  // 10) Domyślny styl akapitu („Normalny”) a tekst. docx-preview dokleja selektor domyślny
  //     PRZED całą listą: „.docx p, p.docx_normalny span { krój… }” — część „span” dostaje tylko
  //     druga połówka. Krój/rozmiar Normalnego trafiały więc na sam akapit, a tekst w nim brał
  //     krój z docDefaults (reguła „.docx span”). Podanie z Normalnym „Liberation Serif 11 pt”
  //     przy docDefaults Aptos/Calibri 12 pt: u nas bezszeryfowe 12 pt, 2 strony zamiast 1
  //     (npm run word:compare, 2026-10-09). Teraz także „.docx :where(p) span” — specyficzność
  //     jak „.docx span” (później = wygrywa z docDefaults), a style znakowe (span.docx_X, też
  //     później) i style akapitów (p.docx_X span) dalej wygrywają z Normalnym, jak w Wordzie.
  const reDef = /(\w+)\[(\w+)\.target\]==\2&&\((\w+)=`\.\$\{this\.className\} \$\{\2\.target\}, `\+\3\)/;
  if (!reDef.test(src)) throw new Error("łatka stylu domyślnego: wzorzec renderStyles nie pasuje (nowa wersja docx-preview?)");
  const mLoop = src.match(/for\(let (\w+) of (\w+)\)\{var (\w+)=`\$\{(\w+)\.target\?\?""\}\.\$\{\4\.cssName\}`;\4\.target!=\1\.target&&\(\3\+=` \$\{\1\.target\}`\),/);
  if (!mLoop) throw new Error("łatka stylu domyślnego: brak pętli podstylów w renderStyles");
  const KT = mLoop[1];
  src = src.replace(reDef, (_, O, J, TT) => `${O}[${J}.target]==${J}&&(${TT}=\`.\${this.className} \${${J}.target}, \`+(${J}.target!=${KT}.target?\`.\${this.className} :where(\${${J}.target}) \${${KT}.target}, \`:"")+${TT})`);
  console.log("  ✅  łatka: domyślny styl akapitu działa też na tekst (krój/rozmiar Normalnego)");

  // 11) Sekcja „ciągła” (w:type continuous) — w Wordzie nowa sekcja zaczyna się NA TEJ SAMEJ
  //     stronie (np. środek strony w 2 kolumnach, inna numeracja wierszy). docx-preview przy
  //     ignoreLastRenderedPageBreak (domyślnie włączone) robił nową kartkę po KAŻDEJ sekcji:
  //     „ANALIZA POTRZEB…” (5 sekcji ciągłych) miała 5 stron zamiast 2 (word:compare, 2026-10-09).
  //     Teraz jak w Wordzie: nowa kartka przy jawnym podziale strony, przy sekcji innego typu niż
  //     ciągła/następna kolumna (brak w:type = nextPage) albo przy zmianie rozmiaru/orientacji.
  //     Typ czytamy z sekcji, która się ZACZYNA (w:type opisuje początek swojej sekcji).
  const reGroup = /groupByPageBreaks\((\w+)\)\{let (\w+)=\[\],(\w+),(\w+)=\[\2\];for\(let (\w+) of \1\)\2\.push\(\5\),\(this\.options\.ignoreLastRenderedPageBreak\|\|\5\.pageBreak\|\|this\.isPageBreakSection\(\3,\5\.sectProps\)\)&&\4\.push\(\2=\[\]\),\3=\5\.sectProps;return \4\.filter\((\w+)=>\6\.length>0\)\}/;
  if (!reGroup.test(src)) throw new Error("łatka sekcji ciągłych: wzorzec groupByPageBreaks nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reGroup, (_, T, E, _R, O) => `groupByPageBreaks(${T}){let ${E}=[],${O}=[${E}];for(let __i=0;__i<${T}.length;__i++){let __k=${T}[__i],__n=${T}[__i+1];${E}.push(__k);if(!__n)break;let __a=__k.sectProps,__b=__n.sectProps;(__k.pageBreak||__a!==__b&&(!/^(continuous|nextColumn)$/.test(__b?.type||"")||this.isPageBreakSection(__a,__b)))&&${O}.push(${E}=[])}return ${O}.filter(__x=>__x.length>0)}`);
  console.log("  ✅  łatka: sekcje ciągłe na tej samej stronie (nowa kartka tylko przy podziale / sekcji „następna strona”)");

  // 12) Znak końca akapitu (w:pPr/w:rPr): kolor i rozmiar. Kolorem Word rysuje punktor/numer listy,
  //     gdy poziom listy nie ma własnego koloru (czerwony akapit = czerwona kropka). docx-preview
  //     czyta pPr/rPr (paragraph.runProps: color, fontSize), ale go nie używa — znaczniki zawsze
  //     czarne („Punkt…”: szare akapity z czarnymi kwadracikami). Kolor trafia do zmiennej
  //     --dwb-mark-color akapitu; app.css daje ją TYLKO znacznikowi (::before) — tekst bez
  //     własnego koloru dalej „auto”, jak w Wordzie. Rozmiar (data-dwb-mark-size) = wysokość
  //     PUSTEGO akapitu w Wordzie (applyWordLineMetrics): „Punkt Marzeny…” — puste akapity stylu
  //     27 pt ze znakiem 15 pt, u nas o połowę wyższe.
  const reMarkRender = /(renderParagraph\((\w+)\)\{var (\w+)=this\.renderContainer\(\2,"p"\);)/;
  if (!reMarkRender.test(src)) throw new Error("łatka znaku akapitu: wzorzec renderParagraph nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reMarkRender, (_, head, T, E) => `${head}{let __rp=${T}.runProps;__rp?.color&&/^[0-9a-f]{6}$/i.test(__rp.color)&&${E}.style.setProperty("--dwb-mark-color","#"+__rp.color);__rp?.fontSize&&(${E}.dataset.dwbMarkSize=__rp.fontSize)}`);
  console.log("  ✅  łatka: kolor znaku końca akapitu dla punktorów/numerów listy");

  // 13) „Nie dodawaj odstępu między akapitami tego samego stylu” (w:contextualSpacing) — jest w
  //     stylu „Akapit z listą” prawie każdego pliku z Worda. docx-preview go nie czytał: każdy
  //     punkt listy dostawał pełny odstęp po (8 pt) — listy 1,5–2× wyższe niż w Wordzie, strona
  //     więcej („Punkt…”, word:compare 2026-10-09). Parser (wspólny dla stylów i akapitów; style
  //     dziedziczą po basedOn) zapamiętuje cechę, akapit dostaje data-dwb-ctx, a odstępy zeruje
  //     applyContextualSpacing (docx-render-fixes.js) — tylko między akapitami tego samego stylu.
  const reCtxParse = /case"keepNext":(\w+)\.keepNext=(\w+)\.boolAttr\((\w+),"val",!0\);break;/;
  if (!reCtxParse.test(src)) throw new Error("łatka contextualSpacing: wzorzec parseParagraphProperty nie pasuje (nowa wersja docx-preview?)");
  src = src.replace(reCtxParse, (m0, T, E, S) => `${m0}case"contextualSpacing":${T}.contextualSpacing=${E}.boolAttr(${S},"val",!0);break;`);
  const reCtxRender = /(renderParagraph\((\w+)\)\{var (\w+)=this\.renderContainer\(\2,"p"\);[^]*?let (\w+)=this\.findStyle\(\2\.styleName\);)/;
  if (!reCtxRender.test(src)) throw new Error("łatka contextualSpacing: wzorzec renderParagraph nie pasuje");
  src = src.replace(reCtxRender, (head, _h, T, E, R) => `${head}(${T}.contextualSpacing??${R}?.paragraphProps?.contextualSpacing)&&(${E}.dataset.dwbCtx="1");`);
  console.log("  ✅  łatka: odstępy między akapitami tego samego stylu (contextualSpacing)");
  fs.writeFileSync(OUT, src);
}

// pdf.js (konwersja PDF → DOCX): moduł główny + worker + wasm (obrazy JPEG2000/JBIG2, profile
// kolorów) + cmapy (PDF-y z tekstem CJK). Wszystko lokalnie — plik nie wychodzi z urządzenia.
function vendorPdfjs() {
  const SRC = path.join(ROOT, "node_modules", "pdfjs-dist");
  if (!fs.existsSync(SRC)) {
    console.warn("⚠️  pdfjs-dist not installed — run npm install first");
    return;
  }
  const OUT_DIR = path.join(ROOT, "lib", "pdfjs");
  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(path.join(OUT_DIR, "wasm"), { recursive: true });
  fs.mkdirSync(path.join(OUT_DIR, "cmaps"), { recursive: true });
  for (const f of ["pdf.min.mjs", "pdf.worker.min.mjs"]) fs.copyFileSync(path.join(SRC, "build", f), path.join(OUT_DIR, f));
  for (const f of ["jbig2.wasm", "openjpeg.wasm", "qcms_bg.wasm", "jbig2_nowasm_fallback.js", "openjpeg_nowasm_fallback.js"]) {
    const from = path.join(SRC, "wasm", f);
    if (fs.existsSync(from)) fs.copyFileSync(from, path.join(OUT_DIR, "wasm", f));
  }
  for (const f of fs.readdirSync(path.join(SRC, "cmaps"))) fs.copyFileSync(path.join(SRC, "cmaps", f), path.join(OUT_DIR, "cmaps", f));
  fs.copyFileSync(path.join(SRC, "LICENSE"), path.join(OUT_DIR, "LICENSE"));
  const ver = JSON.parse(fs.readFileSync(path.join(SRC, "package.json"), "utf8")).version;
  fs.writeFileSync(path.join(OUT_DIR, "VERSION"), ver + "\n");
  console.log(`✅  lib/pdfjs (pdf.js ${ver})`);
}

// tesseract.js (rozpoznawanie tekstu na skanach, PDF → DOCX): skrypt główny, worker, rdzeń wasm
// (SIMD + zwykły, wariant LSTM) i dane języka polskiego. Ładowane dopiero przy pierwszym skanie.
function vendorTesseract() {
  const T = path.join(ROOT, "node_modules", "tesseract.js");
  const CORE = path.join(ROOT, "node_modules", "tesseract.js-core");
  const POL = path.join(ROOT, "node_modules", "@tesseract.js-data", "pol", "4.0.0_best_int", "pol.traineddata.gz");
  if (!fs.existsSync(T) || !fs.existsSync(CORE) || !fs.existsSync(POL)) {
    console.warn("⚠️  tesseract.js / dane pol nie zainstalowane — run npm install first");
    return;
  }
  const OUT_DIR = path.join(ROOT, "lib", "tesseract");
  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(path.join(OUT_DIR, "lang"), { recursive: true });
  fs.copyFileSync(path.join(T, "dist", "tesseract.min.js"), path.join(OUT_DIR, "tesseract.min.js"));
  fs.copyFileSync(path.join(T, "dist", "worker.min.js"), path.join(OUT_DIR, "worker.min.js"));
  for (const f of ["tesseract-core-simd-lstm.js", "tesseract-core-simd-lstm.wasm", "tesseract-core-lstm.js", "tesseract-core-lstm.wasm"]) {
    // obok worker.min.js: rdzeń szuka swojego .wasm względem adresu workera
    fs.copyFileSync(path.join(CORE, f), path.join(OUT_DIR, f));
  }
  fs.copyFileSync(POL, path.join(OUT_DIR, "lang", "pol.traineddata.gz"));
  fs.copyFileSync(path.join(T, "LICENSE.md"), path.join(OUT_DIR, "LICENSE.md"));
  const ver = JSON.parse(fs.readFileSync(path.join(T, "package.json"), "utf8")).version;
  fs.writeFileSync(path.join(OUT_DIR, "VERSION"), `tesseract.js ${ver}, dane pol 4.0.0_best_int\n`);
  console.log(`✅  lib/tesseract (tesseract.js ${ver})`);
}

main().then(vendorPdfjs).then(vendorTesseract).catch((err) => {
  console.error("❌  vendor-libs failed:", err.message || err);
  process.exit(1);
});
