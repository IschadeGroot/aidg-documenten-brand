/*
 * AIDG-koppeling voor de AI van AIDG Documenten (het AI-tabblad in de editors
 * en de AI-agent in het startscherm).
 *
 * De installatie zet dit script vóór de eigen scripts van die twee pagina's.
 * Het leest de verbinding die de AIDG-app per gebruiker neerzet
 * (data/sdkjs-plugins/aidg-koppeling/verbinding.json in het eigen profiel) en
 * zet "AIDG" als aanbieder en model in de instellingen van de AI. Zo werkt de
 * AI zoals online: via het eigen AIDG-account en het gekozen model, zonder dat
 * iemand iets hoeft in te stellen.
 *
 * Heeft de gebruiker zelf een andere aanbieder gekozen, dan blijft die staan;
 * AIDG staat dan wel in de lijst. Gaat er iets mis (geen app, geen bestand),
 * dan doet dit script niets en werkt de pagina zoals altijd.
 */
(function () {
  'use strict';
  try {
    if (!window.AscDesktopEditor || !window.AscDesktopEditor.GetInstallPlugins) return;
    var lijsten = JSON.parse(window.AscDesktopEditor.GetInstallPlugins());
    var eigen = lijsten && lijsten[1] && lijsten[1].url;
    if (!eigen) return;
    // file:///C:/Users/... → onlyoffice://plugin/C:/Users/... (zelfde herkomst als deze pagina)
    var pad = String(eigen).replace(/^file:\/\/\//, '').split(' ').join('%20');
    if (pad.charAt(pad.length - 1) !== '/') pad += '/';
    var xhr = new XMLHttpRequest();
    // Geen ?t=… erachter: dan geeft de office een leeg bestand terug (gemeten, 9.4.0).
    xhr.open('GET', 'onlyoffice://plugin/' + pad + 'aidg-koppeling/verbinding.json', false);
    xhr.send(null);
    if (xhr.status !== 200 && xhr.status !== 0) return;
    var v = JSON.parse(xhr.responseText || 'null');
    if (!v || v.versie !== 1 || !v.baseUrl || !v.sleutel) return;

    var NAAM = v.naam || 'AIDG';
    var aanbieder = { name: NAAM, type: 'openaicompatible', baseUrl: v.baseUrl, key: v.sleutel };
    var model = { id: v.model || 'aidg', name: v.modelNaam || NAAM, provider: NAAM };

    var lees = function (k) {
      try {
        return JSON.parse(window.localStorage.getItem(k) || 'null');
      } catch (e) {
        return null;
      }
    };
    var schrijf = function (k, w) {
      window.localStorage.setItem(k, JSON.stringify(w));
    };

    // De lijst met aanbieders: AIDG erin, of bijgewerkt (poort en sleutel kunnen veranderen).
    var lijst = lees('providers');
    if (!Array.isArray(lijst)) lijst = [];
    var gevonden = false;
    for (var i = 0; i < lijst.length; i++) {
      if (lijst[i] && String(lijst[i].name).toLowerCase() === NAAM.toLowerCase()) {
        lijst[i] = aanbieder;
        gevonden = true;
      }
    }
    if (!gevonden) lijst.unshift(aanbieder);
    schrijf('providers', lijst);

    // Huidige aanbieder en model: AIDG, tenzij de gebruiker bewust iets anders koos.
    var huidig = lees('current-provider');
    if (!huidig || String(huidig.name).toLowerCase() === NAAM.toLowerCase()) {
      schrijf('current-provider', aanbieder);
      var m = lees('current-model');
      if (!m || m.provider === NAAM || !huidig) schrijf('current-model', model);
    }
  } catch (e) {
    /* nooit de AI zelf breken */
  }
})();
