/*
 * Travel Desk — the mark that goes in the generated email's footer.
 *
 * Embedded rather than linked. A remote <img> would make the coordinator's
 * mail client fetch from this add-in's web host every time the form is
 * opened, which is a third host the privacy policy does not claim, is
 * blocked by default in most clients anyway, and looks exactly like a
 * tracking pixel to anyone reviewing automated mail. The bytes travel with
 * the message as an inline attachment instead, so opening the form reaches
 * nothing.
 *
 * This is assets/icon-32.png. Regenerate with:
 *   node tools/make-logo.js
 */
(function (root) {
  "use strict";

  var api = {
    /** base64 PNG, 32x32 */
    BYTES:
    "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",
    CONTENT_TYPE: "image/png",
    FILENAME: "travel-desk.png",
    /** The cid: the email body refers to. */
    CID: "traveldesklogo",
  };

  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.TravelLogo = api; }
})(typeof self !== "undefined" ? self : this);
