// Applies the saved theme before the page draws (dark by default).
(function () {
  var t = 'dark';
  try { t = localStorage.getItem('haem-rota-theme') || 'dark'; } catch (e) { /* storage blocked */ }
  if (t !== 'light' && t !== 'blue') t = 'dark';
  document.documentElement.setAttribute('data-theme', t);
})();
