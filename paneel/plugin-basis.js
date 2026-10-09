// Kleine eigen lader voor een Euro-Office-plugin (in plaats van plugins.js van onlyoffice.github.io: geen
// code van buiten de eigen server, niets van een Amerikaans CDN). Zie docs/werkplek-document-ai.md.
//
// Hoe het werkt (pluginmodel van de editor, sdkjs/common/plugins/plugin_base.js):
//   1. deze pagina staat in een iframe van de editor; na het laden lezen we config.json (guid) en melden we
//      ons bij de editor: {type: 'initialize', guid} naar window.parent;
//   2. de editor antwoordt met {type: 'plugin_init', data: <code>}: dat is de plugin-runtime van de editor
//      zelf. Die voeren we uit; hij vult window.Asc.plugin aan (callCommand, executeMethod, init, …) en zet
//      window.plugin_onMessage;
//   3. daarna gaan alle berichten van de editor naar window.plugin_onMessage.
// Alleen berichten van window.parent (de editor) tellen; andere vensters negeren we.
(function (w) {
  'use strict';

  w.Asc = w.Asc || {};
  w.Asc.plugin = w.Asc.plugin || {};
  w.Asc.scope = w.Asc.scope || {};
  w.Asc.plugin.tr = function (t) { return t; };
  w.Asc.plugin.ie_channel = null;
  w.Asc.plugin.ie_channel_check = function () {};

  var gestart = false;

  function naarEditor(bericht) {
    w.parent.postMessage(JSON.stringify(bericht), '*');
  }

  function opBericht(e) {
    if (e.source !== w.parent) return;
    if (w.plugin_onMessage) return w.plugin_onMessage(e);
    if (gestart || typeof e.data !== 'string') return;
    var b;
    try { b = JSON.parse(e.data); } catch (x) { return; }
    if (!b || b.type !== 'plugin_init' || typeof b.data !== 'string') return;
    gestart = true;
    // De runtime van de editor zelf (zelfde server); zonder deze stap werkt geen enkele plugin.
    // eslint-disable-next-line no-eval
    (0, eval)(b.data);
  }
  w.addEventListener('message', opBericht, false);

  function meldAan() {
    var x = new XMLHttpRequest();
    x.open('GET', './config.json', true);
    x.responseType = 'json';
    x.onload = function () {
      var c = x.response;
      if (typeof c === 'string') { try { c = JSON.parse(c); } catch (e) { c = null; } }
      if (!c || !c.guid) return;
      for (var k in c) if (Object.prototype.hasOwnProperty.call(c, k)) w.Asc.plugin[k] = c[k];
      naarEditor({ type: 'initialize', guid: c.guid });
    };
    x.send();
  }
  if (document.readyState === 'complete') meldAan();
  else w.addEventListener('load', meldAan);
})(window);
