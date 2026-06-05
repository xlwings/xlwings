// Keep the selected item visible in Furo's left sidebar across navigations.
//
// Furo reloads the whole page on navigation, so the sidebar re-renders at
// scrollTop 0. This restores the previous scroll position (same-session nav)
// and otherwise brings the current item into view. scroll-behavior is forced to
// "auto" while we scroll so it snaps instantly — Furo's default
// `scroll-behavior: smooth` is what made it visibly animate ("spin").
//
// Loaded via a plain <script src> placed right after the sidebar markup in
// page.html, so it runs during parse (before first paint) and needs no inline
// code — i.e. it's CSP-friendly (no 'unsafe-inline').
(function () {
  var KEY = "furo-sidebar-scroll";
  var scroller = document.querySelector(".sidebar-scroll");
  if (!scroller) return;

  // Center the current item if it isn't fully in view. Uses the link <a> (not
  // the wrapping <li>, which can be tall and throws off the math) and measured
  // rects, so it doesn't depend on offsetParent.
  function ensureVisible() {
    var prev = scroller.style.scrollBehavior;
    scroller.style.scrollBehavior = "auto";

    var link = scroller.querySelector(".current-page > .reference");
    if (!link) {
      // No current entry in the sidebar tree (e.g. the landing page reached via
      // the logo) -> show the top of the nav, not a restored mid-scroll.
      scroller.scrollTop = 0;
    } else {
      var sRect = scroller.getBoundingClientRect();
      var lRect = link.getBoundingClientRect();
      // Position of the link within the full scroll content.
      var linkInContent = scroller.scrollTop + (lRect.top - sRect.top);
      if (linkInContent < scroller.clientHeight / 2) {
        // Items near the top of the tree belong at the very top, not a restored
        // mid-scroll position.
        scroller.scrollTop = 0;
      } else if (lRect.top < sRect.top || lRect.bottom > sRect.bottom) {
        // Otherwise center it only if it isn't already fully in view.
        scroller.scrollTop +=
          lRect.top - sRect.top - (scroller.clientHeight - lRect.height) / 2;
      }
    }

    scroller.style.scrollBehavior = prev;
  }

  // Early, before first paint: restore the last position (same-session nav).
  var prev = scroller.style.scrollBehavior;
  scroller.style.scrollBehavior = "auto";
  var saved = sessionStorage.getItem(KEY);
  if (saved !== null) scroller.scrollTop = parseInt(saved, 10);
  scroller.style.scrollBehavior = prev;

  ensureVisible();
  // Re-run once layout is fully settled (collapsible trees, fonts, etc.) so deep
  // links / fresh loads land on the right spot, not short of it.
  window.addEventListener("load", ensureVisible);

  scroller.addEventListener("scroll", function () {
    sessionStorage.setItem(KEY, scroller.scrollTop);
  });
})();
