// Privacy-friendly analytics by Plausible.
// External file (not inline) to stay CSP-friendly. The Plausible loader
// itself is pulled in via a separate <script async> tag in page.html.
window.plausible =
  window.plausible ||
  function () {
    (plausible.q = plausible.q || []).push(arguments);
  };
plausible.init =
  plausible.init ||
  function (i) {
    plausible.o = i || {};
  };
plausible.init();
