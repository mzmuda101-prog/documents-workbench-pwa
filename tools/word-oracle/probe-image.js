const { buildDocx } = require("./build-docx");
const word = require("./word");
const JSZip = require("jszip");
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64");
const drawing = '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Obraz 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="a.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
(async () => {
  for (const [k, body] of Object.entries({ samObraz: [{ raw: drawing }], obrazWTekscie: [{ raw: `<w:r><w:t>ab</w:t></w:r>${drawing}<w:r><w:t>cd</w:t></w:r>` }] })) {
    let bytes = await buildDocx({ body });
    const zip = await JSZip.loadAsync(bytes);
    zip.file("word/media/a.png", PNG);
    let rels = await zip.file("word/_rels/document.xml.rels").async("string");
    zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/></Relationships>'));
    let ct = await zip.file("[Content_Types].xml").async("string");
    zip.file("[Content_Types].xml", ct.replace("<Default Extension=\"xml\"", '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"'));
    bytes = await zip.generateAsync({ type: "nodebuffer" });
    require("fs").writeFileSync(`output/word-oracle/img-${k}.docx`, bytes);
    console.log(k, JSON.stringify(word.statistics(bytes)));
  }
})();
