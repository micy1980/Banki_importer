/* core-dom.js — közös DOM helperek és dialog A11y.
 * Globálisan be van töltve az app.js ELŐTT, ezért az itt definiált
 * függvények bare névvel elérhetők a többi szkriptből.
 */

const el = id => document.getElementById(id);

const dialogTriggers = new WeakMap();

function focusableElements(root) {
  return [...root.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])')]
    .filter(item => item.offsetParent !== null || item === document.activeElement);
}

function openDialog(dialogId, trigger = document.activeElement) {
  const dialog = el(dialogId);
  if (!dialog) return;
  dialogTriggers.set(dialog, trigger);
  dialog.showModal();
  requestAnimationFrame(() => {
    const first = focusableElements(dialog)[0];
    (first || dialog).focus();
  });
}

function closeDialog(dialogId) {
  const dialog = el(dialogId);
  if (!dialog?.open) return;
  dialog.close();
}

function setupDialogA11y() {
  document.querySelectorAll("dialog").forEach(dialog => {
    if (dialog.dataset.a11yReady) return;
    dialog.dataset.a11yReady = "1";
    dialog.setAttribute("tabindex", "-1");
    dialog.addEventListener("keydown", event => {
      if (event.key !== "Tab") return;
      const items = focusableElements(dialog);
      if (!items.length) {
        event.preventDefault();
        dialog.focus();
        return;
      }
      const first = items[0];
      const last = items[items.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    });
    dialog.addEventListener("close", () => {
      const trigger = dialogTriggers.get(dialog);
      if (trigger && document.contains(trigger)) {
        trigger.focus();
      }
    });
  });
}

function escapeHtml(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function normalise(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function optionList(headers, selected) {
  let out = `<option value="">-- nincs oszlop --</option>`;
  for (const h of headers) {
    const safe = escapeHtml(h);
    out += `<option value="${safe}" ${h === selected ? "selected" : ""}>${safe}</option>`;
  }
  return out;
}

function emptyState(title, description) {
  return `<div class="empty-state" role="status"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(description)}</span></div>`;
}

function loadingRows(label = "Betöltés...") {
  return `
    <div class="loading-row" role="status" aria-live="polite" aria-label="${escapeHtml(label)}">
      <div class="skeleton medium"></div>
      <div class="skeleton"></div>
      <div class="skeleton short"></div>
    </div>
  `;
}

function setButtonLoading(buttonId, isLoading, labelWhenLoading = "Dolgozom...") {
  const button = el(buttonId);
  if (!button) return;
  if (isLoading) {
    if (!button.dataset.originalText) button.dataset.originalText = button.textContent;
    button.textContent = labelWhenLoading;
    button.dataset.loading = "true";
    button.disabled = true;
  } else {
    button.textContent = button.dataset.originalText || button.textContent;
    delete button.dataset.originalText;
    button.dataset.loading = "false";
    button.disabled = false;
  }
}

function renderListState(listId, rows, renderer, emptyTitle, emptyDescription) {
  const list = el(listId);
  if (!rows?.length) {
    list.innerHTML = emptyState(emptyTitle, emptyDescription);
    return;
  }
  list.innerHTML = rows.map(renderer).join("");
}

async function fetchJson(url, options = {}) {
  const res = await fetch(url, options);
  const contentType = res.headers.get("content-type") || "";
  const data = contentType.includes("application/json") ? await res.json() : { error: await res.text() };
  if (!res.ok) {
    throw new Error(data.error || "A kérés nem sikerült.");
  }
  return data;
}

// Expose helpers as globals (const declarations are script-scoped in classic <script>).
window.el = el;
window.focusableElements = focusableElements;
window.openDialog = openDialog;
window.closeDialog = closeDialog;
window.setupDialogA11y = setupDialogA11y;
window.escapeHtml = escapeHtml;
window.normalise = normalise;
window.optionList = optionList;
window.emptyState = emptyState;
window.loadingRows = loadingRows;
window.setButtonLoading = setButtonLoading;
window.renderListState = renderListState;
window.fetchJson = fetchJson;
