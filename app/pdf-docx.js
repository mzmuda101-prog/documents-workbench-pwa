// PDF → DOCX, krok 3: struktura stron (pdf-layout.js) → plik .docx (OOXML, JSZip).
// Zasady: każda strona PDF zaczyna nową stronę w Wordzie, odstępy pionowe liczone tak, by
// wiersze wypadały tam, gdzie w PDF; tabele z siatką i cieniowaniem; kolumny/karty jako tabela
// bez krawędzi; obrazy w tekście albo zakotwiczone na stronie (nad/pod tekstem).
(function (root) {
  "use strict";
  const NS = (root.DWPdf = root.DWPdf || {});

  const TW = (pt) => Math.round(pt * 20); // twipy
  const EMU = (pt) => Math.round(pt * 12700);
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  // XML 1.0 nie dopuszcza znaków sterujących (poza \t \n \r) i samotnych surogatów.
  const clean = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, "").replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, "");

  // Model wiersza Worda przy interlinii „dokładnie”: wysokość L, linia bazowa ~ BASE·L od góry.
  // (Skalibrowane na Wordzie: tekst siedzi tak, że zejście mieści się na dole wiersza.)
  const BASE = 0.8;
  const MIN_LINE = 1.0; // pt — minimalna wysokość wiersza dla bardzo małych tekstów

  function lineHeightFor(block) {
    if (block.lineGap && block.nLines > 1) return Math.max(MIN_LINE, block.lineGap);
    return Math.max(MIN_LINE, block.size * 1.17);
  }

  // Górna krawędź „pudełka” akapitu w modelu Worda (żeby pierwsza linia bazowa trafiła w PDF).
  function paraBoxTop(block) {
    const L = lineHeightFor(block);
    return block.firstBaseline - BASE * L;
  }
  function paraBoxBottom(block) {
    const L = lineHeightFor(block);
    return paraBoxTop(block) + L * block.nLines;
  }

  function blockTop(b) {
    if (b.type === "para") return paraBoxTop(b);
    if (b.type === "table") return b.y0;
    if (b.type === "columns") return Math.min(...b.cells.map((c) => (c.blocks.length ? blockTop(c.blocks[0]) : c.top)));
    if (b.type === "image") return b.top;
    if (b.type === "rule") return b.y - 1;
    return b.top || 0;
  }
  function blockBottom(b) {
    if (b.type === "para") return paraBoxBottom(b);
    if (b.type === "table") return b.y1;
    if (b.type === "columns") return Math.max(...b.cells.map((c) => (c.blocks.length ? blockBottom(c.blocks[c.blocks.length - 1]) : c.top)));
    if (b.type === "image") return b.bottom;
    if (b.type === "rule") return b.y + 1;
    return b.bottom || 0;
  }

  class DocxWriter {
    constructor(opts = {}) {
      this.media = []; // { name, bytes, ext }
      this.rels = []; // { id, type, target, external }
      this.relSeq = 10;
      this.docPrId = 1;
      this.headingSizes = opts.headingSizes || [];
      this.bodyFont = opts.bodyFont || "Calibri";
      this.bodySize = opts.bodySize || 11;
      this.lang = opts.lang || "pl-PL";
    }

    addRel(type, target, external) {
      const id = "rId" + this.relSeq++;
      this.rels.push({ id, type, target, external: !!external });
      return id;
    }

    addImage(img) {
      const name = `image${this.media.length + 1}.${img.ext}`;
      this.media.push({ name, bytes: img.bytes, ext: img.ext });
      return this.addRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/" + name);
    }

    // ---------------- runy
    runProps(t) {
      const p = [];
      if (t.font) p.push(`<w:rFonts w:ascii="${esc(t.font)}" w:hAnsi="${esc(t.font)}" w:cs="${esc(t.font)}" w:eastAsia="${esc(t.font)}"/>`);
      if (t.bold) p.push("<w:b/><w:bCs/>");
      if (t.italic) p.push("<w:i/><w:iCs/>");
      if (t.strike) p.push("<w:strike/>");
      if (t.color) p.push(`<w:color w:val="${t.color.toUpperCase()}"/>`);
      if (t.hScale && t.hScale >= 1 && t.hScale <= 600) p.push(`<w:w w:val="${t.hScale}"/>`);
      if (t.size) {
        const hp = Math.max(2, Math.round(t.size * 2));
        p.push(`<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`);
      }
      if (t.underline) p.push('<w:u w:val="single"/>');
      if (t.vert) p.push(`<w:vertAlign w:val="${t.vert}"/>`);
      return p.length ? `<w:rPr>${p.join("")}</w:rPr>` : "";
    }

    runsXml(tokens) {
      let out = "";
      let i = 0;
      while (i < tokens.length) {
        const t = tokens[i];
        if (t.type === "tab") {
          // tabulator dziedziczy styl poprzedniego tekstu (wysokość wiersza)
          const prev = tokens.slice(0, i).reverse().find((x) => x.type === "text") || tokens.slice(i).find((x) => x.type === "text") || {};
          out += `<w:r>${this.runProps({ ...prev, underline: false, strike: false })}<w:tab/></w:r>`;
          i++;
          continue;
        }
        if (t.link) {
          const id = this.addRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", t.link, true);
          let inner = "";
          while (i < tokens.length && tokens[i].type === "text" && tokens[i].link === t.link) {
            inner += this.textRun(tokens[i]);
            i++;
          }
          out += `<w:hyperlink r:id="${id}" w:history="1">${inner}</w:hyperlink>`;
          continue;
        }
        out += this.textRun(t);
        i++;
      }
      return out;
    }

    textRun(t) {
      const text = clean(t.text);
      if (!text) return "";
      return `<w:r>${this.runProps(t)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
    }

    // ---------------- akapity
    paraXml(b, ctx) {
      const L = lineHeightFor(b);
      const before = Math.max(0, ctx.spaceBefore || 0);
      const pPr = [];
      if (b.headingLevel) pPr.push(`<w:pStyle w:val="Heading${b.headingLevel}"/>`);
      if (ctx.pageBreakBefore) pPr.push("<w:pageBreakBefore/>");
      pPr.push("<w:widowControl w:val=\"0\"/>");
      if (b.tabs && b.tabs.length) {
        pPr.push("<w:tabs>" + b.tabs.map((t) => `<w:tab w:val="${t.align === "right" ? "right" : "left"}"${t.leader ? ` w:leader="${t.leader}"` : ""} w:pos="${Math.max(0, TW(t.pos))}"/>`).join("") + "</w:tabs>");
      }
      pPr.push(`<w:spacing w:before="${TW(before)}" w:after="0" w:line="${TW(L)}" w:lineRule="exact"/>`);
      const ind = [];
      if (b.indLeft > 0.4) ind.push(`w:left="${TW(b.indLeft)}"`);
      if (Math.abs(b.indRight) > 0.4) ind.push(`w:right="${TW(b.indRight)}"`);
      if (b.firstLine > 0.4) ind.push(`w:firstLine="${TW(b.firstLine)}"`);
      else if (b.firstLine < -0.4) ind.push(`w:hanging="${TW(-b.firstLine)}"`);
      if (ind.length) pPr.push(`<w:ind ${ind.join(" ")}/>`);
      if (b.align && b.align !== "left") pPr.push(`<w:jc w:val="${b.align}"/>`);
      // znacznik akapitu ma rozmiar tekstu (inaczej pusty koniec akapitu zmienia wysokość wiersza)
      // Kolejność w pPr jest ścisła (schemat OOXML): … jc, rPr, sectPr — Word inaczej zgłasza błąd.
      const firstText = b.tokens.find((t) => t.type === "text");
      if (firstText) pPr.push(this.runProps({ font: firstText.font, size: firstText.size }));
      if (ctx.sectPr) pPr.push(ctx.sectPr);
      return `<w:p><w:pPr>${pPr.join("")}</w:pPr>${ctx.anchors || ""}${this.runsXml(b.tokens)}</w:p>`;
    }

    emptyPara(ctx, heightPt) {
      const pPr = [];
      if (ctx.pageBreakBefore) pPr.push("<w:pageBreakBefore/>");
      const h = Math.max(MIN_LINE, heightPt || 1);
      pPr.push(`<w:spacing w:before="${TW(Math.max(0, ctx.spaceBefore || 0))}" w:after="0" w:line="${TW(h)}" w:lineRule="exact"/>`);
      pPr.push(`<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>`);
      if (ctx.sectPr) pPr.push(ctx.sectPr);
      return `<w:p><w:pPr>${pPr.join("")}</w:pPr>${ctx.anchors || ""}</w:p>`;
    }

    ruleXml(b, ctx, box) {
      const pPr = [];
      if (ctx.pageBreakBefore) pPr.push("<w:pageBreakBefore/>");
      const sz = Math.max(2, Math.min(96, Math.round((b.w || 0.75) * 8)));
      pPr.push(`<w:pBdr><w:bottom w:val="single" w:sz="${sz}" w:space="0" w:color="${(b.color || "000000").toUpperCase()}"/></w:pBdr>`);
      pPr.push(`<w:spacing w:before="${TW(Math.max(0, ctx.spaceBefore || 0))}" w:after="0" w:line="20" w:lineRule="exact"/>`);
      const left = Math.max(0, b.x0 - box.x0), right = Math.max(0, box.x1 - b.x1);
      pPr.push(`<w:ind w:left="${TW(left)}" w:right="${TW(right)}"/>`);
      pPr.push(`<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>`);
      if (ctx.sectPr) pPr.push(ctx.sectPr);
      return `<w:p><w:pPr>${pPr.join("")}</w:pPr>${ctx.anchors || ""}</w:p>`;
    }

    // ---------------- obrazy
    drawingInline(img, wPt, hPt) {
      const rid = this.addImage(img);
      const id = this.docPrId++;
      const cx = EMU(wPt), cy = EMU(hPt);
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Obraz ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    }

    drawingAnchor(img, f, z) {
      const rid = this.addImage(img);
      const id = this.docPrId++;
      const cx = EMU(f.x1 - f.x0), cy = EMU(f.bottom - f.top);
      return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}" behindDoc="${f.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${EMU(f.x0)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${EMU(f.top)}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Obraz ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    }

    imageParaXml(b, ctx, box) {
      const img = ctx.images.get(b.atom.img);
      const w = b.atom.x1 - b.atom.x0, h = b.atom.bottom - b.atom.top;
      const pPr = [];
      if (ctx.pageBreakBefore) pPr.push("<w:pageBreakBefore/>");
      // Interlinia „pojedyncza” + czcionka 1 pt: wiersz = wysokość obrazu (przy „dokładnie”
      // przeglądarka dokładała pod obrazem pół wiersza zapasu).
      pPr.push(`<w:spacing w:before="${TW(Math.max(0, ctx.spaceBefore || 0))}" w:after="0" w:line="240" w:lineRule="auto"/>`);
      const left = Math.max(0, b.atom.x0 - box.x0);
      if (left > 0.4) pPr.push(`<w:ind w:left="${TW(left)}"/>`);
      pPr.push(`<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>`);
      if (ctx.sectPr) pPr.push(ctx.sectPr);
      const run = img ? this.drawingInline(img, w, h).replace("<w:r><w:drawing>", `<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:drawing>`) : "";
      return `<w:p><w:pPr>${pPr.join("")}</w:pPr>${ctx.anchors || ""}${run}</w:p>`;
    }

    // ---------------- bloki w obszarze
    // box: { x0, x1, top } — obszar (strona w marginesach, komórka, kolumna)
    blocksXml(blocks, box, ctx) {
      let xml = "";
      let cursor = box.top; // dolna krawędź poprzedniego bloku (model Worda)
      const n = blocks.length;
      for (let i = 0; i < n; i++) {
        const b = blocks[i];
        const first = i === 0;
        const local = {
          images: ctx.images,
          pageBreakBefore: first && ctx.pageBreakBefore,
          anchors: first ? ctx.anchors || "" : "",
          sectPr: i === n - 1 ? ctx.sectPr || "" : "",
        };
        const top = blockTop(b);
        local.spaceBefore = top - cursor;
        // Odstęp „przed” na górze nowej strony Word/LibreOffice pomijają — zamiast niego pusty
        // akapit o dokładnej wysokości (podział strony idzie z nim).
        if (local.pageBreakBefore && local.spaceBefore > 0.5 && b.type !== "table" && b.type !== "columns") {
          xml += this.emptyPara({ pageBreakBefore: true, anchors: local.anchors, spaceBefore: 0 }, local.spaceBefore);
          local.pageBreakBefore = false;
          local.anchors = "";
          local.spaceBefore = 0;
        }
        if (b.type === "para") xml += this.paraXml(b, local);
        else if (b.type === "image") xml += this.imageParaXml(b, local, box);
        else if (b.type === "rule") xml += this.ruleXml(b, local, box);
        else if (b.type === "table" || b.type === "columns") {
          // Tabela nie ma odstępu „przed” — wstawiamy pusty akapit-odstęp (Word i tak wymaga
          // akapitu między sąsiednimi tabelami, inaczej je scala).
          const gap = top - cursor;
          const prevIsTable = i > 0 && (blocks[i - 1].type === "table" || blocks[i - 1].type === "columns");
          if (gap > 1.2 || prevIsTable || local.pageBreakBefore || local.anchors) {
            const h = Math.max(MIN_LINE, gap);
            xml += this.emptyPara({ pageBreakBefore: local.pageBreakBefore, anchors: local.anchors, spaceBefore: 0 }, h);
            local.pageBreakBefore = false;
            local.anchors = "";
          }
          xml += b.type === "table" ? this.tableXml(b, box, ctx) : this.columnsXml(b, box, ctx);
          if (local.sectPr) xml += this.emptyPara({ sectPr: local.sectPr }, 1);
        }
        cursor = blockBottom(b);
        if (b.type === "table" || b.type === "columns") cursor = Math.max(cursor, top);
      }
      if (!n) xml += this.emptyPara({ pageBreakBefore: ctx.pageBreakBefore, anchors: ctx.anchors, sectPr: ctx.sectPr, spaceBefore: 0 }, 1);
      return xml;
    }

    borderXml(side, e) {
      if (!e) return `<w:${side} w:val="nil"/>`;
      const sz = Math.max(2, Math.min(96, Math.round(e.w * 8)));
      return `<w:${side} w:val="single" w:sz="${sz}" w:space="0" w:color="${(e.color || "000000").toUpperCase()}"/>`;
    }

    tableXml(t, box, ctx) {
      const X = t.X, Y = t.Y;
      const nC = X.length - 1, nR = Y.length - 1;
      const grid = [];
      for (let c = 0; c < nC; c++) grid.push(`<w:gridCol w:w="${TW(X[c + 1] - X[c])}"/>`);
      const pad = 2;
      let rows = "";
      for (let r = 0; r < nR; r++) {
        let cells = "";
        for (let c = 0; c < nC; ) {
          const cell = t.cellOf[r][c];
          const span = cell.c1 - cell.c0 + 1;
          const tcPr = [];
          tcPr.push(`<w:tcW w:w="${TW(X[cell.c1 + 1] - X[cell.c0])}" w:type="dxa"/>`);
          if (span > 1) tcPr.push(`<w:gridSpan w:val="${span}"/>`);
          if (cell.r1 > cell.r0) tcPr.push(r === cell.r0 ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>");
          const bd = cell.borders;
          tcPr.push(`<w:tcBorders>${this.borderXml("top", bd.top)}${this.borderXml("left", bd.left)}${this.borderXml("bottom", bd.bottom)}${this.borderXml("right", bd.right)}</w:tcBorders>`);
          if (cell.fill) tcPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${cell.fill.toUpperCase()}"/>`);
          let body;
          if (r === cell.r0) {
            body = this.blocksXml(cell.blocks || [], { x0: cell.x0 + pad, x1: cell.x1 - pad, top: cell.y0 }, { images: ctx.images });
          } else body = this.emptyPara({}, 1);
          if (!body.endsWith("</w:p>")) body += this.emptyPara({}, 1); // komórka musi kończyć się akapitem
          cells += `<w:tc><w:tcPr>${tcPr.join("")}</w:tcPr>${body}</w:tc>`;
          c += span;
        }
        const h = Y[r + 1] - Y[r];
        rows += `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${TW(h)}" w:hRule="atLeast"/></w:trPr>${cells}</w:tr>`;
      }
      const ind = t.x0 - box.x0;
      return `<w:tbl><w:tblPr><w:tblW w:w="${TW(X[nC] - X[0])}" w:type="dxa"/><w:tblInd w:w="${TW(ind)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${TW(pad)}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${TW(pad)}" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${grid.join("")}</w:tblGrid>${rows}</w:tbl>`;
    }

    columnsXml(b, box, ctx) {
      const cols = b.cells;
      const grid = cols.map((c) => `<w:gridCol w:w="${TW(c.x1 - c.x0)}"/>`).join("");
      const top = Math.min(...cols.map((c) => (c.blocks.length ? blockTop(c.blocks[0]) : c.top)));
      let cells = "";
      for (const c of cols) {
        let body = this.blocksXml(c.blocks, { x0: c.x0, x1: c.x1, top }, { images: ctx.images });
        if (!body.endsWith("</w:p>")) body += this.emptyPara({}, 1);
        cells += `<w:tc><w:tcPr><w:tcW w:w="${TW(c.x1 - c.x0)}" w:type="dxa"/><w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders></w:tcPr>${body}</w:tc>`;
      }
      const ind = cols[0].x0 - box.x0;
      return `<w:tbl><w:tblPr><w:tblStyle w:val="DWBLayout"/><w:tblW w:w="${TW(cols[cols.length - 1].x1 - cols[0].x0)}" w:type="dxa"/><w:tblInd w:w="${TW(ind)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>${cells}</w:tr></w:tbl>`;
    }

    // ---------------- części pakietu
    sectPrXml(sec) {
      const orient = sec.width > sec.height ? ' w:orient="landscape"' : "";
      let refs = "";
      if (sec.footerRel) refs += `<w:footerReference w:type="default" r:id="${sec.footerRel}"/>`;
      if (sec.headerRel) refs += `<w:headerReference w:type="default" r:id="${sec.headerRel}"/>`;
      return `<w:sectPr>${refs}<w:pgSz w:w="${TW(sec.width)}" w:h="${TW(sec.height)}"${orient}/><w:pgMar w:top="${TW(sec.mt)}" w:right="${TW(sec.mr)}" w:bottom="${TW(sec.mb)}" w:left="${TW(sec.ml)}" w:header="${TW(sec.headerDist || 18)}" w:footer="${TW(sec.footerDist || 18)}" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
    }

    stylesXml() {
      const f = esc(this.bodyFont);
      const hp = Math.round(this.bodySize * 2);
      let headings = "";
      for (let lvl = 1; lvl <= 6; lvl++) {
        headings += `<w:style w:type="paragraph" w:styleId="Heading${lvl}"><w:name w:val="heading ${lvl}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="${lvl - 1}"/></w:pPr></w:style>`;
      }
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}" w:eastAsia="${f}"/><w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/><w:lang w:val="${this.lang}" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${headings}<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:customStyle="1" w:styleId="DWBLayout"><w:name w:val="Układ z PDF (kolumny)"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="99"/><w:tblPr><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders></w:tblPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`;
    }

    settingsXml() {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat></w:settings>`;
    }

    footerXml(parts) {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${parts}</w:ftr>`;
    }
    headerXml(parts) {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${parts}</w:hdr>`;
    }
  }

  const DOC_NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

  /**
   * Złożenie pakietu .docx.
   * @param pages  [{ width, height, blocks, floats, margins: {ml, mr, mt, mb}, footer?, header? }]
   * @param images Map(obraz PDF → { bytes, ext })
   * @param meta   { title, bodyFont, bodySize }
   */
  async function buildDocx(pages, images, meta, JSZipCtor) {
    const w = new DocxWriter(meta);
    let body = "";
    // sekcje = ciągi stron o tym samym rozmiarze i marginesach
    const sections = [];
    for (const p of pages) {
      const last = sections[sections.length - 1];
      const key = [Math.round(p.width), Math.round(p.height), Math.round(p.margins.ml), Math.round(p.margins.mr), Math.round(p.margins.mt), Math.round(p.margins.mb), p.footerKey || "", p.headerKey || ""].join("|");
      if (last && last.key === key) last.pages.push(p);
      else sections.push({ key, pages: [p], width: p.width, height: p.height, ...p.margins, footer: p.footer, header: p.header });
    }
    let z = 1;
    sections.forEach((sec, si) => {
      if (sec.footer) {
        const id = w.addRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", `footer${si + 1}.xml`);
        sec.footerRel = id;
        sec.footerName = `footer${si + 1}.xml`;
      }
      if (sec.header) {
        const id = w.addRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", `header${si + 1}.xml`);
        sec.headerRel = id;
        sec.headerName = `header${si + 1}.xml`;
      }
      const isLastSection = si === sections.length - 1;
      sec.pages.forEach((p, pi) => {
        const box = { x0: sec.ml, x1: sec.width - sec.mr, top: sec.mt };
        let anchors = "";
        for (const f of p.floats || []) {
          const img = images.get(f.img);
          if (img) anchors += w.drawingAnchor(img, f, (f.behind ? 1000 : 250000) + z++);
        }
        const lastPageOfSection = pi === sec.pages.length - 1;
        const sectPr = lastPageOfSection && !isLastSection ? w.sectPrXml(sec) : "";
        body += w.blocksXml(p.blocks, box, {
          images,
          pageBreakBefore: pi > 0, // pierwsza strona sekcji zaczyna się od podziału sekcji
          anchors,
          sectPr,
        });
      });
    });
    const finalSect = sections.length ? w.sectPrXml(sections[sections.length - 1]) : "";
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DOC_NS}><w:body>${body}${finalSect}</w:body></w:document>`;

    const zip = new JSZipCtor();
    const hasPng = w.media.some((m) => m.ext === "png"), hasJpg = w.media.some((m) => m.ext === "jpeg");
    let overrides = `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;
    if ((meta.fonts || []).length) overrides += `<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`;
    for (const sec of sections) {
      if (sec.footerName) overrides += `<Override PartName="/word/${sec.footerName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`;
      if (sec.headerName) overrides += `<Override PartName="/word/${sec.headerName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
    }
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${hasPng ? '<Default Extension="png" ContentType="image/png"/>' : ""}${hasJpg ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : ""}${(meta.fonts || []).length ? '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>' : ""}${overrides}</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    const rels = [
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`,
      ...((meta.fonts || []).length ? [`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`] : []),
      ...w.rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`),
    ];
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);
    zip.file("word/document.xml", documentXml);
    zip.file("word/styles.xml", w.stylesXml());
    let settings = w.settingsXml();
    // Osadzone czcionki (pdf-fonts.js): fontTable + zaciemnione .odttf
    const fonts = meta.fonts || [];
    if (fonts.length) {
      const byFam = new Map();
      fonts.forEach((f, i) => {
        if (!byFam.has(f.family)) byFam.set(f.family, []);
        byFam.get(f.family).push({ ...f, file: `font${i + 1}.odttf`, rid: `rId${i + 1}` });
      });
      let ft = "";
      const frels = [];
      for (const [fam, list] of byFam) {
        ft += `<w:font w:name="${esc(fam)}"><w:charset w:val="EE"/><w:pitch w:val="variable"/>${list.map((f) => `<w:embed${f.style} r:id="${f.rid}" w:fontKey="${f.key}"/>`).join("")}</w:font>`;
        for (const f of list) {
          frels.push(`<Relationship Id="${f.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${f.file}"/>`);
          zip.file("word/fonts/" + f.file, f.bytes);
        }
      }
      zip.file("word/fontTable.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${ft}</w:fonts>`);
      zip.file("word/_rels/fontTable.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${frels.join("")}</Relationships>`);
      settings = settings.replace("<w:zoom w:percent=\"100\"/>", "<w:zoom w:percent=\"100\"/><w:embedTrueTypeFonts/><w:saveSubsetFonts/>");
    }
    zip.file("word/settings.xml", settings);
    for (const sec of sections) {
      if (sec.footerName) zip.file("word/" + sec.footerName, w.footerXml(sec.footer));
      if (sec.headerName) zip.file("word/" + sec.headerName, w.headerXml(sec.header));
    }
    for (const m of w.media) zip.file("word/media/" + m.name, m.bytes);
    const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${meta.title ? `<dc:title>${esc(clean(meta.title))}</dc:title>` : ""}${meta.author ? `<dc:creator>${esc(clean(meta.author))}</dc:creator>` : ""}<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Documents Workbench (PDF → DOCX)</Application><Pages>${pages.length}</Pages></Properties>`);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  NS.buildDocx = buildDocx;
  NS.blockTop = blockTop;
  NS.blockBottom = blockBottom;
  NS.lineHeightFor = lineHeightFor;
  NS.DOCX_BASE = BASE;
})(typeof globalThis !== "undefined" ? globalThis : window);
