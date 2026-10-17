const { buildDocx } = require("./build-docx");
const app = require("./app");
(async () => {
  const bytes = await buildDocx({
    abstractNums: [{ levels: [{ fmt: "lowerLetter", text: "%1)" }] }],
    nums: [{ id: 1, abs: 0 }],
    body: ["Wstęp", ...Array.from({ length: 28 }, (_, i) => ({ text: `poz ${i + 1}`, numId: 1 }))],
  });
  const { browser, page } = await app.openApp();
  await app.loadBytes(page, bytes);
  console.log(JSON.stringify(await app.listLabels(page)));
  await browser.close();
})();
