// Panel „Recenzja” — śledzone zmiany, komentarze i przypisy z pliku Worda.
// Lista ze skokiem do miejsca, filtr autora, ✓ / ✗ przy każdej zmianie, zbiorcze
// Akceptuj / Odrzuć (dla wybranego autora) i usuwanie komentarzy. Logika: docx-revisions.js.

const rvSummaryEl = document.getElementById("rvSummary");
const rvAuthorEl = document.getElementById("rvAuthor");
const rvAcceptAllBtn = document.getElementById("rvAcceptAllBtn");
const rvRejectAllBtn = document.getElementById("rvRejectAllBtn");
const rvChangesEl = document.getElementById("rvChanges");
const rvCommentsEl = document.getElementById("rvComments");
const rvNotesEl = document.getElementById("rvNotes");
const rvRemoveCommentsBtn = document.getElementById("rvRemoveCommentsBtn");

const RV_LIST_LIMIT = 150;
let reviewScan = null;
let rvShowAll = false;

const RV_KIND_KEY = {
  ins: "reviewKindIns", del: "reviewKindDel", moveFrom: "reviewKindMoveFrom", moveTo: "reviewKindMoveTo",
  format: "reviewKindFormat", paraFormat: "reviewKindParaFormat", otherFormat: "reviewKindOtherFormat",
  paraIns: "reviewKindParaIns", paraDel: "reviewKindParaDel", rowIns: "reviewKindRowIns", rowDel: "reviewKindRowDel",
};
const RV_PART_KEY = { header: "reviewInHeader", footer: "reviewInFooter", footnote: "reviewInFootnote", endnote: "reviewInEndnote" };

function rvDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleDateString(I18N[currentLang].locale, { day: "numeric", month: "short", year: "numeric" });
}

function rvShort(text, max = 110) {
  const s = String(text || "").replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function rvEl(tag, className, text) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text != null) el.textContent = text;
  return el;
}

function rvFilteredChanges() {
  const author = rvAuthorEl?.value || "";
  return (reviewScan?.changes || []).filter((c) => !author || c.author === author);
}

function rvJump(paraIndex) {
  if (!Number.isFinite(paraIndex)) return;
  const el = collectPreviewParagraphElements(docCanvasEl?.querySelector(".docx-preview-host"))[paraIndex];
  if (!el) return;
  jumpToStructureItem({ el, id: "review" }, { silentSelect: true });
  if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
}

async function rvApply(edit, confirmKey, vars) {
  if (!originalFileBytes) return;
  if (confirmKey && !confirm(t(confirmKey, vars))) return;
  const n = await applyDocumentEdit({ op: "revisions", ...edit });
  toast(n > 0 ? t(edit.action === "removeComments" ? "reviewCommentsRemoved" : edit.action === "accept" ? "reviewAccepted" : "reviewRejected", { count: n }) : t("editNothing"), n > 0 ? "success" : "info");
  await runReviewScan();
}

function rvChangeRow(c) {
  const row = rvEl("div", `rv-row rv-${c.kind}`);
  const head = rvEl("div", "rv-head");
  head.append(rvEl("span", "rv-chip", t(RV_KIND_KEY[c.kind] || "reviewKindOtherFormat")));
  const meta = [c.author, rvDate(c.date), RV_PART_KEY[c.part] ? t(RV_PART_KEY[c.part]) : ""].filter(Boolean).join(" · ");
  head.append(rvEl("span", "rv-meta", meta));
  const actions = rvEl("span", "rv-actions");
  const ok = rvEl("button", "rv-act rv-accept", "✓");
  ok.type = "button";
  ok.setAttribute("aria-label", t("reviewAcceptOne"));
  ok.dataset.hint = "";
  ok.dataset.hintPl = "Akceptuj tę zmianę";
  ok.dataset.hintEn = "Accept this change";
  ok.addEventListener("click", (e) => { e.stopPropagation(); rvApply({ action: "accept", ids: [c.id] }); });
  const no = rvEl("button", "rv-act rv-reject", "✗");
  no.type = "button";
  no.setAttribute("aria-label", t("reviewRejectOne"));
  no.dataset.hint = "";
  no.dataset.hintPl = "Odrzuć tę zmianę";
  no.dataset.hintEn = "Reject this change";
  no.addEventListener("click", (e) => { e.stopPropagation(); rvApply({ action: "reject", ids: [c.id] }); });
  actions.append(ok, no);
  head.append(actions);
  row.append(head);
  const body = rvShort(c.text) || rvShort(c.context, 80);
  if (body) row.append(rvEl("div", c.text ? "rv-text" : "rv-text rv-context", body));
  if (Number.isFinite(c.paraIndex)) {
    row.classList.add("is-jump");
    row.tabIndex = 0;
    row.addEventListener("click", () => rvJump(c.paraIndex));
    row.addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target === row) rvJump(c.paraIndex); });
  }
  return row;
}

function rvCommentRow(c) {
  const row = rvEl("div", "rv-row rv-comment");
  const head = rvEl("div", "rv-head");
  head.append(rvEl("span", "rv-chip", c.initials || c.author.slice(0, 2).toUpperCase()));
  head.append(rvEl("span", "rv-meta", [c.author, rvDate(c.date), c.done ? t("reviewResolved") : ""].filter(Boolean).join(" · ")));
  const actions = rvEl("span", "rv-actions");
  const del = rvEl("button", "rv-act rv-reject", "✗");
  del.type = "button";
  del.setAttribute("aria-label", t("reviewDeleteComment"));
  del.dataset.hint = "";
  del.dataset.hintPl = "Usuń komentarz (z odpowiedziami)";
  del.dataset.hintEn = "Delete comment (with replies)";
  del.addEventListener("click", (e) => { e.stopPropagation(); rvApply({ action: "removeComments", ids: [c.id, ...c.replies.map((r) => r.id)] }); });
  actions.append(del);
  head.append(actions);
  row.append(head);
  if (c.anchor) row.append(rvEl("div", "rv-anchor", `„${rvShort(c.anchor, 80)}”`));
  row.append(rvEl("div", "rv-text", c.text));
  c.replies.forEach((r) => {
    const reply = rvEl("div", "rv-reply");
    reply.append(rvEl("span", "rv-meta", [r.author, rvDate(r.date)].filter(Boolean).join(" · ")), rvEl("div", "rv-text", r.text));
    row.append(reply);
  });
  if (Number.isFinite(c.paraIndex)) {
    row.classList.add("is-jump");
    row.tabIndex = 0;
    row.addEventListener("click", () => rvJump(c.paraIndex));
    row.addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target === row) rvJump(c.paraIndex); });
  }
  return row;
}

function rvNoteRow(n) {
  const row = rvEl("div", "rv-row rv-note");
  const head = rvEl("div", "rv-head");
  head.append(rvEl("span", "rv-chip", `${t(n.type === "footnote" ? "reviewFootnote" : "reviewEndnote")} ${n.id}`));
  row.append(head, rvEl("div", "rv-text", rvShort(n.text, 160)));
  if (Number.isFinite(n.paraIndex)) {
    row.classList.add("is-jump");
    row.tabIndex = 0;
    row.addEventListener("click", () => rvJump(n.paraIndex));
  }
  return row;
}

function rvEmpty(key) {
  return rvEl("p", "hint", t(key));
}

function renderReviewPanel() {
  if (!rvChangesEl) return;
  rvChangesEl.replaceChildren();
  rvCommentsEl.replaceChildren();
  rvNotesEl.replaceChildren();
  if (!originalFileBytes || !reviewScan) {
    if (rvSummaryEl) rvSummaryEl.textContent = originalFileBytes ? "" : t("noFileToSave");
    [rvAcceptAllBtn, rvRejectAllBtn, rvRemoveCommentsBtn].forEach((b) => { if (b) b.disabled = true; });
    return;
  }
  const { changes, comments, notes, authors } = reviewScan;
  const replies = comments.reduce((n, c) => n + c.replies.length, 0);
  if (rvSummaryEl) {
    rvSummaryEl.textContent = t("reviewSummary", { changes: changes.length, authors: authors.size, comments: comments.length + replies, notes: notes.length });
  }

  // autorzy: zachowaj wybór, jeśli autor nadal ma zmiany
  if (rvAuthorEl) {
    const keep = rvAuthorEl.value;
    rvAuthorEl.replaceChildren(new Option(t("reviewAllAuthors", { count: changes.length }), ""));
    [...authors.entries()].sort((a, b) => b[1] - a[1]).forEach(([name, n]) => rvAuthorEl.add(new Option(`${name} (${n})`, name)));
    rvAuthorEl.value = authors.has(keep) ? keep : "";
    rvAuthorEl.closest(".field")?.classList.toggle("hidden", authors.size < 2);
  }

  const list = rvFilteredChanges();
  if (!list.length) rvChangesEl.append(rvEmpty("reviewNoChanges"));
  (rvShowAll ? list : list.slice(0, RV_LIST_LIMIT)).forEach((c) => rvChangesEl.append(rvChangeRow(c)));
  if (!rvShowAll && list.length > RV_LIST_LIMIT) {
    const more = rvEl("button", "btn", t("reviewShowMore", { count: list.length - RV_LIST_LIMIT }));
    more.type = "button";
    more.addEventListener("click", () => { rvShowAll = true; renderReviewPanel(); });
    rvChangesEl.append(more);
  }
  if (!comments.length) rvCommentsEl.append(rvEmpty("reviewNoComments"));
  comments.forEach((c) => rvCommentsEl.append(rvCommentRow(c)));
  if (!notes.length) rvNotesEl.append(rvEmpty("reviewNoNotes"));
  notes.forEach((n) => rvNotesEl.append(rvNoteRow(n)));

  if (rvAcceptAllBtn) rvAcceptAllBtn.disabled = !list.length;
  if (rvRejectAllBtn) rvRejectAllBtn.disabled = !list.length;
  if (rvRemoveCommentsBtn) rvRemoveCommentsBtn.disabled = !comments.length;
  const label = rvAuthorEl?.value ? "reviewAcceptAuthor" : "reviewAcceptAll";
  if (rvAcceptAllBtn) rvAcceptAllBtn.textContent = t(label, { author: rvAuthorEl?.value || "" });
  if (rvRejectAllBtn) rvRejectAllBtn.textContent = t(rvAuthorEl?.value ? "reviewRejectAuthor" : "reviewRejectAll", { author: rvAuthorEl?.value || "" });
}

async function runReviewScan() {
  rvShowAll = false;
  reviewScan = originalFileBytes ? await scanDocxRevisions(originalFileBytes) : null;
  renderReviewPanel();
}

function rvBulk(action) {
  const author = rvAuthorEl?.value || "";
  const list = rvFilteredChanges();
  if (!list.length) return;
  const ids = author ? list.map((c) => c.id) : null;
  rvApply({ action, ids }, action === "accept" ? "reviewAcceptConfirm" : "reviewRejectConfirm", { count: list.length });
}

rvAcceptAllBtn?.addEventListener("click", () => rvBulk("accept"));
rvRejectAllBtn?.addEventListener("click", () => rvBulk("reject"));
rvRemoveCommentsBtn?.addEventListener("click", () => {
  const n = (reviewScan?.comments || []).reduce((s, c) => s + 1 + c.replies.length, 0);
  rvApply({ action: "removeComments" }, "reviewRemoveCommentsConfirm", { count: n });
});
rvAuthorEl?.addEventListener("change", renderReviewPanel);
