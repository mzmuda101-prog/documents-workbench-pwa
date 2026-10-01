// Panel „Formularz” — pola z Worda (kontrolki zawartości i stare pola formularza) jako lista
// z właściwymi przyciskami: pole wyboru, lista, data, tekst. Zmiana od razu trafia do pliku
// (jeden krok Cofnij na zmianę). Logika odczytu/zapisu i klikanie w dokumencie: docx-forms.js.

const ffSummaryEl = document.getElementById("ffSummary");
const ffProtectionEl = document.getElementById("ffProtection");
const ffListEl = document.getElementById("ffList");
const ffShadeEl = document.getElementById("ffShade");

const FF_KIND_KEY = { checkbox: "formsKindCheckbox", dropdown: "formsKindDropdown", combo: "formsKindCombo", date: "formsKindDate", text: "formsKindText" };

function ffEl(tag, className, text) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text != null) el.textContent = text;
  return el;
}

function ffControl(f) {
  const id = `ff-${f.key}`;
  if (f.kind === "checkbox") {
    const label = ffEl("label", "field checkbox ff-check");
    const box = ffEl("input");
    box.type = "checkbox";
    box.id = id;
    box.checked = !!f.value;
    box.disabled = f.locked;
    box.addEventListener("change", () => fillFormFields({ [f.key]: box.checked }));
    label.append(box, ffEl("span", "", formFieldName(f)));
    return label;
  }
  let input;
  if (f.kind === "dropdown") {
    input = ffEl("select");
    if (!f.value || f.placeholder) input.add(new Option(f.placeholder ? f.display : "—", "", true, true));
    f.options.forEach((o) => input.add(new Option(o.text, o.value, false, !f.placeholder && o.value === f.value)));
  } else {
    input = ffEl(f.multiLine ? "textarea" : "input");
    if (f.kind === "date") input.type = "date";
    else if (!f.multiLine) input.type = "text";
    input.value = f.kind === "date" ? f.value || "" : f.placeholder ? "" : f.kind === "combo" ? f.display : f.value;
    if (f.kind !== "date" && f.placeholder) input.placeholder = f.display;
    if (f.maxLength) input.maxLength = f.maxLength;
    if (f.kind === "combo" && f.options.length) {
      const list = ffEl("datalist");
      list.id = `${id}-list`;
      f.options.forEach((o) => list.append(new Option(o.text)));
      input.setAttribute("list", list.id);
      input._list = list;
    }
    input.autocomplete = "off";
  }
  input.id = id;
  input.disabled = f.locked;
  input.setAttribute("aria-label", formFieldName(f));
  input.addEventListener("change", () => {
    if (f.kind === "dropdown" && !input.value) return;
    if (f.kind === "date" && !input.value) return;
    fillFormFields({ [f.key]: input.value });
  });
  if (!f.multiLine && input.tagName === "INPUT") {
    input.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); input.blur(); } });
  }
  return input._list ? [input, input._list] : input;
}

function ffRow(f) {
  const row = ffEl("div", `rv-row ff-row${formIsEmpty(f) ? " is-empty" : ""}`);
  row.dataset.key = f.key;
  const head = ffEl("div", "rv-head");
  head.append(ffEl("span", "rv-chip", t(FF_KIND_KEY[f.kind])));
  const meta = [
    f.source === "legacy" ? t("formsLegacy") : "",
    f.copies ? t("formsCopies", { count: f.copies }) : "",
    f.locked ? t("formsLockedShort") : "",
  ].filter(Boolean).join(" · ");
  head.append(ffEl("span", "rv-meta", meta));
  const actions = ffEl("span", "rv-actions");
  if (Number.isFinite(f.paraIndex)) {
    const go = ffEl("button", "rv-act", "↗");
    go.type = "button";
    go.setAttribute("aria-label", t("formsJump"));
    go.dataset.hint = "";
    go.dataset.hintPl = "Pokaż pole w dokumencie";
    go.dataset.hintEn = "Show the field in the document";
    go.addEventListener("click", () => {
      jumpToFormField(f.key);
      if (typeof closeMobileSidebarIfOpen === "function") closeMobileSidebarIfOpen();
    });
    actions.append(go);
  }
  head.append(actions);
  row.append(head);
  const ctrl = [ffControl(f)].flat();
  if (f.kind === "checkbox") {
    row.append(...ctrl);
    return row;
  }
  const wrap = ffEl("label", "field ff-field"); // ten sam wygląd pól co w reszcie panelu
  wrap.append(ffEl("span", "", formFieldName(f)), ...ctrl);
  row.append(wrap);
  return row;
}

function renderFormsPanel() {
  if (!ffListEl) return;
  // po przerysowaniu dokumentu lista powstaje od nowa — fokus wraca do tego samego pola
  const focusKey = document.activeElement?.closest?.(".ff-row")?.dataset.key;
  const scrollParent = ffListEl.closest(".sidebar-scroll") || ffListEl.closest(".sidebar");
  const keepScroll = scrollParent?.scrollTop;
  ffListEl.replaceChildren();
  formUi.panelBytes = formScan?.bytes || null;
  if (ffShadeEl) ffShadeEl.checked = formShadeOn();
  if (!originalFileBytes) {
    if (ffSummaryEl) ffSummaryEl.textContent = t("noFileToSave");
    ffProtectionEl?.classList.add("hidden");
    return;
  }
  const main = formMainFields();
  if (ffSummaryEl) {
    ffSummaryEl.textContent = main.length ? t("formsSummary", { count: main.length, empty: docFormCounts.empty }) : t("formsNone");
  }
  ffProtectionEl?.classList.toggle("hidden", formScan?.protection !== "forms");
  main.forEach((f) => ffListEl.append(ffRow(f)));
  if (focusKey) ffListEl.querySelector(`.ff-row[data-key="${focusKey}"] input, .ff-row[data-key="${focusKey}"] select, .ff-row[data-key="${focusKey}"] textarea`)?.focus({ preventScroll: true });
  if (scrollParent && keepScroll != null) scrollParent.scrollTop = keepScroll;
}

ffShadeEl?.addEventListener("change", () => setFormShade(ffShadeEl.checked));
