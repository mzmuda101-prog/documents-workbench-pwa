// find-replace-playwright.js — workbench: skan DOM, podgląd XML, zamiana jednego i wszystkich.

const path = require("path");
const {
  DOCX_FIXTURE,
  fixtureExists,
  createTestPage,
  bootApp,
  loadDocxFile,
  assertNoErrors,
} = require("./docx-test-helpers");

const DOCX_ARG = process.argv[2];

async function run() {
  const filePath = DOCX_ARG ? path.resolve(DOCX_ARG) : DOCX_FIXTURE;
  if (!fixtureExists(filePath)) {
    console.warn(`⚠️  Pominięto — brak pliku: ${filePath}`);
    process.exit(0);
  }

  const { browser, page, errors } = await createTestPage();
  await bootApp(page);
  await loadDocxFile(page, filePath, 90000);

  const result = await page.evaluate(async () => {
    await ensureDocLibs(false);
    const texts = await extractParagraphTextsFromDocx(originalFileBytes);
    const joined = texts.join(" ");
    const words = joined.match(/[\p{L}\p{N}]{5,}/gu) || [];
    const needle = words.find((w) => (joined.match(new RegExp(w, "g")) || []).length >= 2) || words[0];
    if (!needle) return { ok: false, step: "needle", msg: "Brak słowa do testu" };

    const markerOne = "[[FR_ONE]]";
    const markerAll = "[[FR_ALL]]";

    const q = document.getElementById("searchQuery");
    const r = document.getElementById("frReplace");
    if (!q || !r) return { ok: false, step: "ui", msg: "Brak pól find/replace" };

    q.value = needle;
    r.value = markerOne;
    await runDocumentSearch();
    const domHits = document.querySelectorAll(".search-hit").length;
    if (!domHits) return { ok: false, step: "scan", msg: `Brak trafień DOM dla: ${needle}` };

    const preview = countReplacePreview(texts, { op: "replace", find: needle, replace: markerOne, regex: false, scope: "all" });
    if (!preview.hits) return { ok: false, step: "preview", msg: "Podgląd XML — 0 trafień" };
    if (!document.querySelector(".fr-preview-item")) {
      return { ok: false, step: "preview-ui", msg: "Brak listy podglądu" };
    }

    const oneCount = await applyDocumentEdit(
      { op: "replace", find: needle, replace: markerOne, regex: false, scope: "all" },
      { maxReplacements: 1, skipReplacements: 0 }
    );
    if (!oneCount) return { ok: false, step: "replace-one", msg: "Zamiana jednego — 0" };

    const afterOne = await extractParagraphTextsFromDocx(originalFileBytes);
    const oneInXml = afterOne.join(" ").split(markerOne).length - 1;
    if (oneInXml !== 1) return { ok: false, step: "replace-one-verify", msg: `Oczekiwano 1× ${markerOne}, jest ${oneInXml}` };

    q.value = needle;
    r.value = markerAll;
    const allCount = await applyDocumentEdit(
      { op: "replace", find: needle, replace: markerAll, regex: false, scope: "all" }
    );
    if (!allCount) return { ok: false, step: "replace-all", msg: "Zamiana wszystkich — 0" };

    const afterAll = await extractParagraphTextsFromDocx(originalFileBytes);
    const joinedAfter = afterAll.join(" ");
    if (joinedAfter.includes(needle)) {
      return { ok: false, step: "replace-all-verify", msg: `Pozostało: ${needle}` };
    }
    if (!joinedAfter.includes(markerOne) || !joinedAfter.includes(markerAll)) {
      return { ok: false, step: "replace-all-markers", msg: "Brak markerów po zamianie wszystkich" };
    }

    return { ok: true, needle, domHits, previewHits: preview.hits, oneCount, allCount };
  });

  if (!result.ok) throw new Error(`find-replace failed at ${result.step}: ${result.msg}`);
  assertNoErrors(errors, "find-replace");
  await browser.close();
  console.log(
    `✅  find-replace-playwright passed (“${result.needle}”, DOM=${result.domHits}, preview=${result.previewHits}, one=${result.oneCount}, all=${result.allCount})`
  );
}

run().catch((err) => {
  console.error("❌  find-replace-playwright failed:", err.message || err);
  process.exit(1);
});
