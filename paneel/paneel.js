// AI-collega in AIDG Documenten (desktop). Tegenhanger van het paneel AIDG in AIDG Werkplek (online),
// maar de chat is die van de AIDG-app op deze pc (eigenaarsbesluit 09-10-2026: "deze resultaten in de
// desktop chat"):
//
// 1. De app zet in het eigen profiel verbinding.json neer (aidg-koppeling/), met `app: {poort, sleutel}`:
//    waar de app luistert. Geen app (of nog niet gestart)? Dan zegt dit paneel dat, en kijkt elke paar
//    seconden opnieuw.
// 2. Verbinden met ws://127.0.0.1:<poort>/office/verbinden, zelfde protocol als online (aanmelden,
//    welkom, opdracht/antwoord, ping/pong) plus: vraag → het gesprek in de app, chat ← het antwoord.
// 3. Opdrachten van de app (de AI gebruikt word_*, excel_*, powerpoint_*) één voor één uitvoeren met
//    opdrachten.js (callCommand), precies zoals online.
//
// Nooit: inhoud van het document of de sleutel in de console.
(function () {
  'use strict';

  var VERSIE = 1;
  var PING_MS = 25000;
  var GEEN_PONG_MS = 60000;
  var COMMANDO_MS = 45000;
  var OPNIEUW_MS = 4000;
  var OPSLAG_VENSTER = 'aidg-desktop-venster';

  var W = window.AidgWerkplek;
  var $ = function (id) { return document.getElementById(id); };
  var staat = {
    programma: null,
    document: '',
    ws: null,
    verbonden: false,
    klok: null,
    pingKlok: null,
    laatstePong: 0,
    stoppen: false,
    rij: Promise.resolve(),
    uitvoerder: null,
    aiBericht: null, // het antwoord dat nu binnenkomt
  };
  var vensterId = maakId();

  function maakId() {
    try {
      var b = window.sessionStorage.getItem(OPSLAG_VENSTER);
      if (b && /^[A-Za-z0-9_-]{8,64}$/.test(b)) return b;
    } catch (e) { /* geen opslag */ }
    var r = new Uint8Array(12);
    window.crypto.getRandomValues(r);
    var id = 'd' + Array.prototype.map.call(r, function (x) { return ('0' + x.toString(16)).slice(-2); }).join('');
    try { window.sessionStorage.setItem(OPSLAG_VENSTER, id); } catch (e) { /* geen opslag */ }
    return id;
  }

  // ------------------------------------------------------------------ verbinding met de app
  /** verbinding.json van de AIDG-app (zelfde plek en manier als aidg-koppeling.js), of null. */
  function leesVerbinding() {
    try {
      if (!window.AscDesktopEditor || !window.AscDesktopEditor.GetInstallPlugins) return null;
      var lijsten = JSON.parse(window.AscDesktopEditor.GetInstallPlugins());
      var eigen = lijsten && lijsten[1] && lijsten[1].url;
      if (!eigen) return null;
      var pad = String(eigen).replace(/^file:\/\/\//, '').split(' ').join('%20');
      if (pad.charAt(pad.length - 1) !== '/') pad += '/';
      var xhr = new XMLHttpRequest();
      // Geen ?query erachter: dan geeft de office een leeg bestand terug (gemeten, 9.4.0).
      xhr.open('GET', 'onlyoffice://plugin/' + pad + 'aidg-koppeling/verbinding.json', false);
      xhr.send(null);
      if (xhr.status !== 200 && xhr.status !== 0) return null;
      var v = JSON.parse(xhr.responseText || 'null');
      return v && v.app && v.app.poort && v.app.sleutel ? v.app : null;
    } catch (e) {
      return null;
    }
  }

  // ------------------------------------------------------------------ weergave
  function zetStatus(tekst, soort) {
    $('status').textContent = tekst;
    $('stip').className = 'stip' + (soort ? ' ' + soort : '');
  }
  function toonMelding(tekst, knoppen) {
    var vak = $('melding');
    if (!tekst) { vak.hidden = true; return; }
    $('melding-tekst').textContent = tekst;
    var k = $('melding-knoppen');
    k.textContent = '';
    (knoppen || []).forEach(function (kn) {
      var el = document.createElement('button');
      el.type = 'button';
      el.className = 'knop' + (kn.tweede ? ' tweede' : '');
      el.textContent = kn.tekst;
      el.addEventListener('click', kn.doe);
      k.appendChild(el);
    });
    vak.hidden = false;
  }
  function invoerAan(aan) {
    $('vraag').disabled = !aan;
    $('stuur').disabled = !aan;
  }
  function bericht(soort, tekst) {
    $('welkom').hidden = true;
    var el = document.createElement('div');
    el.className = 'bericht ' + soort;
    el.textContent = tekst;
    $('gesprek').appendChild(el);
    $('gesprek').scrollTop = $('gesprek').scrollHeight;
    return el;
  }
  var VOORBEELDEN = {
    Word: ['Wat staat er in dit document?', 'Zet onder de laatste alinea een korte afsluiting.', 'Maak de tweede alinea formeler.'],
    Excel: ['Lees dit blad.', 'Zet in de eerste lege kolom een totaal per rij.', 'Maak er een grafiek van.'],
    PowerPoint: ['Lees de presentatie.', 'Maak een dia met drie punten over de planning.'],
  };
  var WAT = {
    word_lezen: 'document gelezen', word_selectie_lezen: 'selectie gelezen', word_invoegen: 'tekst ingevoegd',
    word_vervangen: 'tekst vervangen', word_opmerking: 'opmerking geplaatst', word_wijzigingen_bijhouden: 'wijzigingen bijhouden',
    excel_werkbladen: 'werkbladen bekeken', excel_lezen: 'cellen gelezen', excel_selectie_lezen: 'selectie gelezen',
    excel_schrijven: 'cellen geschreven', excel_opmaak: 'opgemaakt', excel_tabel_maken: 'tabel gemaakt',
    excel_grafiek_maken: 'grafiek gemaakt', excel_werkblad_toevoegen: 'werkblad toegevoegd',
    powerpoint_lezen: "dia's gelezen", powerpoint_dia_toevoegen: 'dia toegevoegd', powerpoint_tekst_vervangen: 'tekst vervangen',
  };
  function toonLaatst(tool, gelukt) {
    var nu = new Date();
    var tijd = ('0' + nu.getHours()).slice(-2) + ':' + ('0' + nu.getMinutes()).slice(-2);
    $('laatst').textContent = 'Laatst (' + tijd + '): ' + (WAT[tool] || tool.replace(/_/g, ' ')) + (gelukt ? '' : ' — lukte niet');
    $('laatst').hidden = false;
  }

  // ------------------------------------------------------------------ editor
  function roep(functie, scope) {
    return new Promise(function (resolve, reject) {
      var klaar = false;
      var klok = window.setTimeout(function () {
        if (klaar) return;
        klaar = true;
        reject(new Error('De editor reageerde niet binnen ' + COMMANDO_MS / 1000 + ' seconden (staat er een dialoogvenster open?).'));
      }, COMMANDO_MS);
      window.Asc.scope = JSON.parse(JSON.stringify(scope || {}));
      window.Asc.plugin.callCommand(functie, false, true, function (uit) {
        if (klaar) return;
        klaar = true;
        window.clearTimeout(klok);
        resolve(uit);
      });
    });
  }

  // ------------------------------------------------------------------ verbinden
  function stuur(b) {
    if (staat.ws && staat.ws.readyState === 1) staat.ws.send(JSON.stringify(b));
  }

  function later() {
    staat.ws = null;
    staat.verbonden = false;
    invoerAan(false);
    window.clearInterval(staat.pingKlok);
    if (!staat.stoppen) staat.klok = window.setTimeout(verbind, OPNIEUW_MS);
  }

  function verbind() {
    if (staat.stoppen || staat.ws) return;
    window.clearTimeout(staat.klok);
    var app = leesVerbinding();
    if (!app) {
      zetStatus('De AIDG-app draait niet', 'fout');
      toonMelding('Start de AIDG-app op deze pc; dit paneel verbindt dan vanzelf. De AI-knoppen bovenin (AI-tabblad) werken ook zonder.', []);
      staat.klok = window.setTimeout(verbind, OPNIEUW_MS);
      return;
    }
    zetStatus('Verbinden met de AIDG-app…', 'bezig');
    var ws;
    try {
      ws = new WebSocket('ws://127.0.0.1:' + Number(app.poort) + '/office/verbinden');
    } catch (e) {
      return later();
    }
    staat.ws = ws;
    ws.onopen = function () {
      ws.send(JSON.stringify({
        type: 'aanmelden',
        versie: VERSIE,
        venster: vensterId,
        host: staat.programma,
        document: staat.document,
        platform: 'AIDG Documenten',
        bewijs: String(app.sleutel),
      }));
    };
    ws.onmessage = function (e) {
      var b;
      try { b = JSON.parse(e.data); } catch (x) { return; }
      if (!b || ws !== staat.ws) return;
      staat.laatstePong = Date.now();
      if (b.type === 'welkom') {
        staat.verbonden = true;
        toonMelding('');
        zetStatus('Verbonden met de AIDG-app', 'goed');
        invoerAan(true);
        $('nieuw').hidden = false;
        startPing();
      } else if (b.type === 'opdracht') {
        staat.rij = staat.rij.then(function () { return voerUit(ws, b); });
      } else if (b.type === 'chat') {
        toonAntwoord(String(b.tekst || ''), b.bezig === true);
      }
    };
    ws.onclose = function (e) {
      if (ws !== staat.ws) return;
      if (e.code === 4409) {
        // Zelfde venster-id elders (gedupliceerd tabblad): eigen id en opnieuw.
        try { window.sessionStorage.removeItem(OPSLAG_VENSTER); } catch (x) { /* geen opslag */ }
        vensterId = maakId();
      }
      if (e.code === 4401) zetStatus('De AIDG-app herkende dit paneel niet; opnieuw…', 'fout');
      else zetStatus('Verbinding met de AIDG-app weg; opnieuw…', 'fout');
      later();
    };
  }

  function startPing() {
    window.clearInterval(staat.pingKlok);
    staat.laatstePong = Date.now();
    staat.pingKlok = window.setInterval(function () {
      if (!staat.ws) return;
      if (Date.now() - staat.laatstePong > GEEN_PONG_MS) {
        try { staat.ws.close(); } catch (e) { /* al dicht */ }
        return;
      }
      stuur({ type: 'ping' });
    }, PING_MS);
  }

  function voerUit(ws, b) {
    zetStatus('De AI werkt in je document…', 'bezig');
    var tool = String(b.tool);
    return Promise.resolve()
      .then(function () { return staat.uitvoerder(tool, b.args || {}); })
      .then(function (tekst) {
        toonLaatst(tool, true);
        return { type: 'antwoord', id: b.id, ok: true, tekst: String(tekst) };
      }, function (e) {
        toonLaatst(tool, false);
        return { type: 'antwoord', id: b.id, ok: false, fout: W.foutTekst(e) };
      })
      .then(function (antwoord) {
        if (ws.readyState === 1) ws.send(JSON.stringify(antwoord));
        if (staat.verbonden) zetStatus('Verbonden met de AIDG-app', 'goed');
      });
  }

  // ------------------------------------------------------------------ chat
  /** Eenvoudige opmaak van het antwoord (vet, cursief, code, opsommingen). Eerst alles escapen. */
  function opmaak(tekst) {
    var veilig = String(tekst)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    var regels = veilig.split('\n');
    var uit = [];
    var inLijst = false;
    regels.forEach(function (r) {
      var m = /^\s*(?:[-*•]|\d+[.)])\s+(.*)$/.exec(r);
      if (m) {
        if (!inLijst) { uit.push('<ul>'); inLijst = true; }
        uit.push('<li>' + m[1] + '</li>');
        return;
      }
      if (inLijst) { uit.push('</ul>'); inLijst = false; }
      var kop = /^#{1,4}\s+(.*)$/.exec(r);
      uit.push(kop ? '<b>' + kop[1] + '</b><br>' : r + '<br>');
    });
    if (inLijst) uit.push('</ul>');
    return uit.join('')
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
      .replace(/(^|[\s(])[*_]([^*_\s][^*_]*)[*_](?=[\s).,!?:;]|$)/g, '$1<i>$2</i>')
      .replace(/(<br>)+$/, '');
  }

  function toonAntwoord(tekst, bezig) {
    if (!staat.aiBericht) staat.aiBericht = bericht('ai', '');
    if (tekst) staat.aiBericht.innerHTML = opmaak(tekst);
    else staat.aiBericht.textContent = bezig ? 'Even denken' : '';
    staat.aiBericht.classList.toggle('bezig', bezig);
    $('gesprek').scrollTop = $('gesprek').scrollHeight;
    if (!bezig) staat.aiBericht = null;
  }

  function vraag(tekst) {
    tekst = String(tekst || '').trim();
    if (!tekst || !staat.verbonden) return;
    bericht('wij', tekst);
    staat.aiBericht = null;
    toonAntwoord('', true);
    stuur({ type: 'vraag', tekst: tekst });
  }

  function nieuwGesprek() {
    var g = $('gesprek');
    Array.prototype.slice.call(g.querySelectorAll('.bericht')).forEach(function (el) { g.removeChild(el); });
    $('welkom').hidden = false;
    staat.aiBericht = null;
    stuur({ type: 'nieuw' });
  }

  // ------------------------------------------------------------------ start
  window.Asc.plugin.init = function () {
    var info = window.Asc.plugin.info || {};
    staat.programma = W.PROGRAMMA[info.editorType] || null;
    staat.document = String(info.documentTitle || '').slice(0, 200);
    if (staat.document) $('docnaam').textContent = staat.document;
    if (!staat.programma) {
      zetStatus('De AI kan (nog) niet in dit soort document werken.', 'fout');
      return;
    }
    var vb = $('voorbeelden');
    VOORBEELDEN[staat.programma].forEach(function (t) {
      var k = document.createElement('button');
      k.type = 'button';
      k.className = 'voorbeeld';
      k.textContent = t;
      k.addEventListener('click', function () { vraag(t); });
      vb.appendChild(k);
    });
    $('vraag').placeholder = 'Vraag over ' + (staat.document || 'dit document');
    staat.uitvoerder = W.maakUitvoerder({ programma: staat.programma, roep: roep });
    $('invoer').addEventListener('submit', function (e) {
      e.preventDefault();
      var t = $('vraag').value;
      $('vraag').value = '';
      vraag(t);
    });
    $('vraag').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        $('invoer').requestSubmit ? $('invoer').requestSubmit() : $('stuur').click();
      }
    });
    $('nieuw').addEventListener('click', nieuwGesprek);
    $('in-app').addEventListener('click', function () { stuur({ type: 'toon_in_app' }); });
    verbind();
  };

  window.Asc.plugin.onThemeChanged = function (thema) {
    document.body.classList.toggle('donker', Boolean(thema && /dark/i.test(String(thema.type || ''))));
  };

  window.Asc.plugin.button = function () {
    stop();
    window.Asc.plugin.executeCommand('close', '');
  };

  function stop() {
    staat.stoppen = true;
    window.clearTimeout(staat.klok);
    window.clearInterval(staat.pingKlok);
    try { if (staat.ws) staat.ws.close(1000, 'venster dicht'); } catch (e) { /* al dicht */ }
  }
  window.addEventListener('pagehide', stop);
})();
