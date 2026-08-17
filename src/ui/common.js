export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function escapeAttr(value) {
  return escapeHtml(value).replace(/'/g, '&#39;');
}

export function cssEscape(value) {
  if (typeof CSS !== 'undefined' && CSS.escape) return CSS.escape(value);
  return String(value).replace(/["\\]/g, '\\$&');
}

export function formatCount(n) {
  return Number(n).toLocaleString('en-GB');
}

export function qs(root, selector) {
  return root.querySelector(selector);
}

export function setStatusLines(el, lines) {
  if (!el) return;
  el.innerHTML =
    '<div class="status-lines">' +
    lines
      .map(function (line) {
        return '<div>' + escapeHtml(line) + '</div>';
      })
      .join('') +
    '</div>';
}

export function openConfirm(title, message, onConfirm, confirmLabel) {
  var modal = document.getElementById('confirm-modal');
  document.getElementById('confirm-title').textContent = title;
  document.getElementById('confirm-message').textContent = message;
  var okBtn = document.getElementById('confirm-ok');
  var cancelBtn = document.getElementById('confirm-cancel');
  okBtn.textContent = confirmLabel || 'Confirm';
  modal.classList.remove('hidden');

  var cleanup = function () {
    okBtn.onclick = null;
    cancelBtn.onclick = null;
    modal.querySelectorAll('[data-role="confirm-close"]').forEach(function (el) {
      el.onclick = null;
    });
  };

  var close = function () {
    cleanup();
    closeConfirm();
  };

  okBtn.onclick = function () {
    close();
    onConfirm();
  };
  cancelBtn.onclick = close;
  modal.querySelectorAll('[data-role="confirm-close"]').forEach(function (el) {
    el.onclick = close;
  });
}

export function closeConfirm() {
  document.getElementById('confirm-modal').classList.add('hidden');
}
