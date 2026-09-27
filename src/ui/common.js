export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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
