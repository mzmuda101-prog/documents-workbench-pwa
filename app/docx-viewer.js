// DOCX preview rendering via docx-preview (lazy-loaded global).

async function renderDocxPreview(bytes, container) {
  if (!container) return;
  container.replaceChildren();
  const wrapper = document.createElement("div");
  wrapper.className = "docx-preview-host";
  container.appendChild(wrapper);

  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await window.docx.renderAsync(ab, wrapper, null, {
    className: "docx",
    inWrapper: true,
    ignoreWidth: false,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
  });
  fixDocxBulletRendering(wrapper);
}
