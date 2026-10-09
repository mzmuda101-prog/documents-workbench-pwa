// Word's ¶ toggle, kept deliberately visual-only: it never changes DOCX bytes.
const DWB_FORMAT_MARKS_KEY = "dwb-show-formatting-marks-v1";
const formatMarksBtn = document.getElementById("formatMarksBtn");

function syncFormatMarks(on) {
  const active = !!on && !!originalFileBytes;
  docCanvasEl?.classList.toggle("show-formatting-marks", active);
  formatMarksBtn?.classList.toggle("is-on", active);
  formatMarksBtn?.setAttribute("aria-pressed", String(active));
  if (formatMarksBtn) formatMarksBtn.disabled = !originalFileBytes;
}

function restoreFormatMarks() {
  let on = false;
  try { on = localStorage.getItem(DWB_FORMAT_MARKS_KEY) === "1"; } catch (_) { /* private mode */ }
  syncFormatMarks(on);
}

formatMarksBtn?.addEventListener("click", () => {
  if (!originalFileBytes) return;
  const next = !docCanvasEl?.classList.contains("show-formatting-marks");
  try { localStorage.setItem(DWB_FORMAT_MARKS_KEY, next ? "1" : "0"); } catch (_) { /* still works this session */ }
  syncFormatMarks(next);
});
