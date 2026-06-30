// DOCX preview rendering via docx-preview (lazy-loaded global).

async function renderDocxPreview(bytes, container) {
  if (!container) return;
  container.replaceChildren();
  const wrapper = document.createElement("div");
  wrapper.className = "docx-preview-host";
  container.appendChild(wrapper);

  const mobileReflow = typeof shouldUseMobileReflow === "function" && shouldUseMobileReflow();
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await window.docx.renderAsync(ab, wrapper, null, {
    className: "docx",
    inWrapper: true,
    ignoreWidth: mobileReflow,
    ignoreHeight: mobileReflow,
    ignoreFonts: false,
    breakPages: !mobileReflow,
    renderHeaders: true,
    renderFooters: !mobileReflow,
    renderFootnotes: true,
    renderEndnotes: true,
  });
  fixDocxBulletRendering(wrapper);
  if (mobileReflow && typeof applyMobileReflowLayout === "function") applyMobileReflowLayout(wrapper);
}
