/* Nine Four Format: core transform. Pure function over an HTML string so it
   can run in Outlook (browser DOMParser) and in tests (jsdom). */
(function (root) {
  var BODY_FONT = "Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif";
  var BODY_SIZE = "11pt";
  var SIG_FONT = "'Calibri Light', 'Calibri Light_EmbeddedFont', 'Calibri Light_MSFontService', Calibri, sans-serif";
  var SIG_SIZE = "10pt";
  var SIG_KEY = "KurtRamirezGeneralPartnerNINEFOURVENTURES";

  // Where the reply/forward quote or an Outlook-inserted signature begins.
  var STOP_SELECTOR = [
    "#appendonsend",
    "[id*='divRplyFwdMsg']",
    "#mail-editor-reference-message-container",
    "[id='Signature']",
    "[id='x_Signature']"
  ].join(",");

  function squash(s) { return (s || "").replace(/\s+/g, ""); }

  function isBlank(node) {
    if (node.nodeType === 3) return !node.textContent.trim();
    if (node.nodeType === 8) return true;
    if (node.nodeType !== 1) return false;
    if (node.querySelector && node.querySelector("img,table,hr")) return false;
    return !node.textContent.replace(/ /g, " ").trim();
  }

  function styleBody(el) {
    el.style.setProperty("font-family", BODY_FONT);
    el.style.setProperty("font-size", BODY_SIZE);
  }

  function transform(html, parse) {
    var doc = parse(html);
    var body = doc.body;

    var stop = body.querySelector(STOP_SELECTOR);
    var outlookSig = !!(stop && /Signature/.test(stop.id || ""));

    // Pull out everything above the quote/signature as the new message text.
    var range = doc.createRange();
    range.setStart(body, 0);
    if (stop) range.setEndBefore(stop); else range.setEnd(body, body.childNodes.length);
    var frag = range.extractContents();

    var wrap = doc.createElement("div");
    wrap.appendChild(frag);

    // Remove a signature this button added on an earlier click.
    var all = wrap.querySelectorAll("*");
    for (var i = all.length - 1; i >= 0; i--) {
      var el = all[i];
      if (squash(el.textContent) === SIG_KEY && el.parentNode) {
        var inner = false;
        for (var c = 0; c < el.children.length; c++) {
          if (squash(el.children[c].textContent) === SIG_KEY) { inner = true; break; }
        }
        if (!inner) { el.parentNode.removeChild(el); break; }
      }
    }

    // Trim trailing empty lines.
    while (wrap.lastChild && isBlank(wrap.lastChild)) wrap.removeChild(wrap.lastChild);

    // Paragraphs: zero margins, explicit blank line between them (how Outlook types).
    var ps = wrap.querySelectorAll("p");
    for (var j = 0; j < ps.length; j++) {
      var p = ps[j];
      p.style.setProperty("margin", "0");
      var next = p.nextElementSibling;
      if (next && next.tagName === "P") {
        var spacer = doc.createElement("div");
        spacer.appendChild(doc.createElement("br"));
        p.parentNode.insertBefore(spacer, next);
      }
    }

    // Font on every element, so no stylesheet rule can override it.
    var els = wrap.querySelectorAll("*");
    for (var k = 0; k < els.length; k++) {
      var e = els[k];
      if (e.tagName === "FONT") { e.removeAttribute("face"); e.removeAttribute("size"); }
      if (e.tagName === "BR") continue;
      styleBody(e);
    }
    styleBody(wrap);
    wrap.style.setProperty("color", "rgb(0, 0, 0)");

    if (!outlookSig) {
      var gap = doc.createElement("div");
      styleBody(gap);
      gap.appendChild(doc.createElement("br"));
      wrap.appendChild(gap);

      var sig = doc.createElement("div");
      sig.style.setProperty("font-family", SIG_FONT);
      sig.style.setProperty("font-size", SIG_SIZE);
      sig.style.setProperty("color", "rgb(0, 0, 0)");
      sig.innerHTML = "Kurt Ramirez<br>General Partner<br><b>NINE FOUR</b> VENTURES";
      wrap.appendChild(sig);
    }

    // When replying, keep a blank line between the signature and the quote.
    if (stop) {
      var after = doc.createElement("div");
      after.appendChild(doc.createElement("br"));
      wrap.appendChild(after);
    }

    body.insertBefore(wrap, body.firstChild);
    return "<html>" + doc.documentElement.innerHTML + "</html>";
  }

  root.NineFourFormat = { transform: transform, SIG_KEY: SIG_KEY };
  if (typeof module !== "undefined") module.exports = root.NineFourFormat;
})(typeof window !== "undefined" ? window : globalThis);
