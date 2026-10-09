// Opdrachten van de chat uitvoeren in het open document in AIDG Werkplek (Euro-Office), met de Document
// Builder API van de editor. Tegenhanger van office/opdrachten.js (Office.js); zelfde tools, zelfde
// argumenten, zelfde soort antwoorden, zodat de bestaande MCP-tools van office-aidg (sso/office.mjs) hier
// ongewijzigd werken. Zie docs/werkplek-document-ai.md.
//
// HOE. Elke opdracht draait als functie IN DE EDITOR (Asc.plugin.callCommand): de editor zet de functie om
// naar tekst en voert hem daar uit, met alleen Asc.scope als invoer. Zo'n functie mag dus GEEN variabelen of
// hulpfuncties van hierbuiten gebruiken; alles wat hij nodig heeft, staat in de functie zelf of in de scope.
// Hij geeft {ok: true, tekst} of {ok: true, data} terug, of {fout: '…'} met een Nederlandse uitleg voor de
// AI. Elke callCommand is één stap voor Ongedaan maken, en de wijziging gaat meteen naar iedereen die het
// document open heeft (samen bewerken).
//
// WIJZIGINGEN BIJHOUDEN (Word). Zoals de Office-invoegtoepassing: vóór de eerste bewerking in dit venster
// zetten we Wijzigingen bijhouden aan (als het uit stond). Staat het aan, dan gaan de wijzigingen op naam
// van "AIDG" (SetAssistantTrackRevisions), zodat iedereen ziet wat de AI deed en het kan accepteren of
// weigeren. Eén keer per venster: zet de gebruiker of de AI het daarna uit, dan blijft het uit.
//
// Werkt in de browser (window.AidgWerkplek) en in Node voor de tests (module.exports, nep-editor).
(function (wortel, maak) {
  if (typeof module === 'object' && module.exports) module.exports = maak();
  else wortel.AidgWerkplek = maak();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var MAX_TEKENS = 50000;
  var MAX_CELLEN = 4000; // lezen
  var MAX_KOLOMMEN = 60;
  var MAX_SCHRIJVEN = 20000; // cellen per keer
  var TABEL_RIJEN = 50;
  var AUTEUR = 'AIDG';
  var GRAFIEK = { kolom: 'bar', staaf: 'horizontalBar', lijn: 'line', taart: 'pie', vlak: 'area', spreiding: 'scatter' };
  var KLEUREN = {
    zwart: '#000000', black: '#000000', wit: '#FFFFFF', white: '#FFFFFF', rood: '#C00000', red: '#FF0000',
    groen: '#00B050', green: '#008000', blauw: '#0070C0', blue: '#0000FF', geel: '#FFFF00', yellow: '#FFFF00',
    oranje: '#FFC000', orange: '#FFA500', grijs: '#808080', gray: '#808080', grey: '#808080', paars: '#7030A0',
    purple: '#800080', aidg: '#153041',
  };

  /** Editor (info.editorType) → programma zoals de relay het kent. */
  var PROGRAMMA = { word: 'Word', cell: 'Excel', slide: 'PowerPoint' };

  /** Eigen fout met een uitleg die zo naar de AI mag. */
  function melding(tekst) {
    var e = new Error(tekst);
    e.aidgMelding = true;
    return e;
  }

  function foutTekst(e) {
    if (!e) return 'Dit lukte niet in AIDG Werkplek.';
    if (e.aidgMelding) return e.message;
    return String(e.message || e).slice(0, 300) || 'Dit lukte niet in AIDG Werkplek.';
  }

  function geheel(w, standaard) {
    var n = Number(w);
    return Number.isFinite(n) ? Math.floor(n) : standaard;
  }

  // ------------------------------------------------------------------ cellen en bereiken (Excel)
  /** Kolomletters: 1 → A, 28 → AB. */
  function kolomLetters(n) {
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /** "Blad1!A1:B2" → {blad: 'Blad1', adres: 'A1:B2'}; zonder blad: {blad: '', adres}. */
  function splitsBereik(bereik) {
    var b = String(bereik || '').trim();
    var i = b.lastIndexOf('!');
    if (i < 0) return { blad: '', adres: b.replace(/\$/g, '') };
    return { blad: b.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'"), adres: b.slice(i + 1).replace(/\$/g, '') };
  }

  /** "B2" → {rij: 2, kolom: 2}; "B2:D5" → ook {rij2, kolom2}. Geen geldig adres → null. */
  function leesAdres(adres) {
    var m = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/i.exec(String(adres || '').replace(/\$/g, '').trim());
    if (!m) return null;
    var kol = function (l) {
      var k = 0;
      for (var j = 0; j < l.length; j++) k = k * 26 + (l.toUpperCase().charCodeAt(j) - 64);
      return k;
    };
    var a = { rij: Number(m[2]), kolom: kol(m[1]) };
    a.rij2 = m[3] ? Number(m[4]) : a.rij;
    a.kolom2 = m[3] ? kol(m[3]) : a.kolom;
    if (a.rij2 < a.rij) { var t = a.rij; a.rij = a.rij2; a.rij2 = t; }
    if (a.kolom2 < a.kolom) { var u = a.kolom; a.kolom = a.kolom2; a.kolom2 = u; }
    return a;
  }

  function adresVan(rij, kolom, rij2, kolom2) {
    var a = kolomLetters(kolom) + rij;
    return rij2 === rij && kolom2 === kolom ? a : a + ':' + kolomLetters(kolom2) + rij2;
  }

  function celTekst(waarde, formule) {
    var w = waarde === null || waarde === undefined ? '' : String(waarde);
    if (typeof formule === 'string' && formule.charAt(0) === '=') return formule + ' → ' + w;
    return w;
  }

  /** Cellen als tekst, zoals in office/opdrachten.js: kolomletters bovenaan, rijnummers ervoor. */
  function bereikTekst(d) {
    var regels = [];
    var kop = ['   '];
    var kolommen = d.waarden.length ? d.waarden[0].length : 0;
    for (var k = 0; k < kolommen; k++) kop.push(kolomLetters(d.kolom + k));
    regels.push(kop.join(' | '));
    var tekens = 0;
    for (var r = 0; r < d.waarden.length; r++) {
      var cellen = [String(d.rij + r)];
      for (var c = 0; c < kolommen; c++) cellen.push(celTekst(d.waarden[r][c], d.formules[r][c]).replace(/\s*\n\s*/g, ' ⏎ '));
      var regel = cellen.join(' | ');
      tekens += regel.length;
      if (tekens > MAX_TEKENS) {
        regels.push('[… ingekort na rij ' + (d.rij + r - 1) + ': lees verder met een kleiner bereik]');
        break;
      }
      regels.push(regel);
    }
    var maat = d.totaalRijen + ' rijen × ' + d.totaalKolommen + ' kolommen';
    var ingekort = d.totaalRijen > d.waarden.length || d.totaalKolommen > kolommen
      ? ' (alleen de eerste ' + d.waarden.length + ' rijen × ' + kolommen + ' kolommen getoond; lees de rest met een kleiner bereik)'
      : '';
    return (d.blad ? d.blad + '!' : '') + d.adres + ': ' + maat + ingekort + '\n' + regels.join('\n');
  }

  /**
   * Formule in de Engelse schrijfwijze (komma tussen argumenten, punt als decimaalteken, zo vraagt de tool
   * het) → de schrijfwijze van de editor. In het Nederlands verwacht Euro-Office ; tussen argumenten en een
   * komma als decimaalteken: "=IF(A1>1,5,0.5)" → "=IF(A1>1;5;0,5)". Tekst tussen "…" en bladnamen tussen
   * '…' blijven ongemoeid; in een matrixconstante {…} is de komma het kolomscheidingsteken.
   */
  function lokaleFormule(f, sep) {
    if (!sep || (sep.argument === ',' && sep.decimaal === '.')) return f;
    var uit = '';
    var accolades = 0;
    for (var i = 0; i < f.length; i++) {
      var c = f.charAt(i);
      if (c === '"' || c === "'") {
        var j = i + 1;
        while (j < f.length) {
          if (f.charAt(j) === c) {
            if (f.charAt(j + 1) === c) { j += 2; continue; }
            break;
          }
          j++;
        }
        uit += f.slice(i, j + 1);
        i = j;
        continue;
      }
      if (c === '{') accolades++;
      else if (c === '}') accolades = Math.max(0, accolades - 1);
      if (c === ',') uit += accolades ? sep.kolom || sep.argument : sep.argument;
      else if (c === '.' && /\d/.test(f.charAt(i - 1)) && /\d/.test(f.charAt(i + 1))) uit += sep.decimaal;
      else uit += c;
    }
    return uit;
  }

  /** #RRGGBB of een kleurnaam → [r, g, b]; onbekend → null. */
  function rgb(kleur) {
    var k = String(kleur || '').trim().toLowerCase();
    if (KLEUREN[k]) k = KLEUREN[k].toLowerCase();
    var m = /^#?([0-9a-f]{6})$/.exec(k) || /^#?([0-9a-f]{3})$/.exec(k);
    if (!m) return null;
    var h = m[1].length === 3 ? m[1].replace(/./g, '$&$&') : m[1];
    return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  }

  // ------------------------------------------------------------------ tekst en HTML → blokken (Word)
  /**
   * Eenvoudige HTML (zoals word_invoegen hem toestaat: h1-h3, p, b/strong, i/em, u, br, ul/ol/li, table)
   * → blokken die de editor-functie zelf opbouwt. Geen DOM nodig (werkt ook in Node). Onbekende tags: alleen
   * de tekst telt. Script en style worden overgeslagen.
   *   blok: {soort: 'p'|'h1'|'h2'|'h3'|'li', stukken: [{t, b, i, u} | {br: true}], lijst, genummerd, niveau}
   *       | {soort: 'tabel', rijen: [[tekst]], kop: true/false}
   */
  function htmlNaarBlokken(html) {
    var ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', euro: '€' };
    var ontsnap = function (t) {
      return t.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, function (heel, n) {
        if (n.charAt(0) === '#') {
          var c = n.charAt(1).toLowerCase() === 'x' ? parseInt(n.slice(2), 16) : parseInt(n.slice(1), 10);
          return c > 0 && c < 0x110000 ? String.fromCodePoint(c) : '';
        }
        return Object.prototype.hasOwnProperty.call(ENT, n.toLowerCase()) ? ENT[n.toLowerCase()] : heel;
      });
    };
    var blokken = [];
    var huidig = null;
    var opmaak = { b: 0, i: 0, u: 0 };
    var lijsten = []; // stapel {genummerd, id}
    var lijstTeller = 0;
    var tabel = null; // {rijen, rij, cel, kop}
    var overslaan = 0;

    function nieuwBlok(soort) {
      sluitBlok();
      huidig = { soort: soort, stukken: [] };
      if (soort === 'li') {
        var l = lijsten[lijsten.length - 1] || { genummerd: false, id: 'los' };
        huidig.lijst = l.id;
        huidig.genummerd = l.genummerd;
        huidig.niveau = Math.max(0, lijsten.length - 1);
      }
    }
    function sluitBlok() {
      if (huidig && huidig.stukken.some(function (s) { return s.br || /\S/.test(s.t); })) {
        // Spaties aan begin en eind van een alinea weg (zoals een browser).
        var st = huidig.stukken;
        if (st[0] && !st[0].br) st[0].t = st[0].t.replace(/^\s+/, '');
        var l = st[st.length - 1];
        if (l && !l.br) l.t = l.t.replace(/\s+$/, '');
        huidig.stukken = st.filter(function (s) { return s.br || s.t; });
        blokken.push(huidig);
      }
      huidig = null;
    }
    function tekst(t) {
      if (overslaan) return;
      t = ontsnap(t.replace(/\s+/g, ' '));
      if (!t) return;
      if (tabel && tabel.cel !== null) { tabel.cel += t; return; }
      if (!huidig) {
        if (!/\S/.test(t)) return;
        nieuwBlok('p');
      }
      huidig.stukken.push({ t: t, b: opmaak.b > 0, i: opmaak.i > 0, u: opmaak.u > 0 });
    }

    var re = /<!--[\s\S]*?-->|<(\/?)([a-z][a-z0-9]*)\b[^>]*?(\/?)>|([^<]+)|</gi;
    var m;
    while ((m = re.exec(String(html || '')))) {
      if (m[4] !== undefined) { tekst(m[4]); continue; }
      if (!m[2]) { if (m[0] === '<') tekst('<'); continue; }
      var tag = m[2].toLowerCase();
      var sluit = m[1] === '/';
      if (tag === 'script' || tag === 'style') { overslaan += sluit ? -1 : 1; overslaan = Math.max(0, overslaan); continue; }
      if (overslaan) continue;
      switch (tag) {
        case 'b': case 'strong': opmaak.b += sluit ? -1 : 1; break;
        case 'i': case 'em': opmaak.i += sluit ? -1 : 1; break;
        case 'u': opmaak.u += sluit ? -1 : 1; break;
        case 'br':
          if (tabel && tabel.cel !== null) tabel.cel += '\n';
          else { if (!huidig) nieuwBlok('p'); huidig.stukken.push({ br: true }); }
          break;
        case 'p': case 'div': case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
          if (tabel) break;
          if (sluit) sluitBlok();
          else nieuwBlok(/^h[1-3]$/.test(tag) ? tag : /^h[4-6]$/.test(tag) ? 'h3' : 'p');
          break;
        case 'ul': case 'ol':
          if (tabel) break;
          sluitBlok();
          if (sluit) lijsten.pop();
          else lijsten.push({ genummerd: tag === 'ol', id: 'l' + ++lijstTeller });
          break;
        case 'li':
          if (tabel) break;
          if (sluit) sluitBlok();
          else nieuwBlok('li');
          break;
        case 'table':
          sluitBlok();
          if (!sluit) tabel = { rijen: [], rij: null, cel: null, kop: false };
          else if (tabel) {
            if (tabel.rij) tabel.rijen.push(tabel.rij);
            if (tabel.rijen.length) blokken.push({ soort: 'tabel', rijen: tabel.rijen, kop: tabel.kop });
            tabel = null;
          }
          break;
        case 'tr':
          if (!tabel) break;
          if (tabel.cel !== null) { tabel.rij.push(tabel.cel.trim()); tabel.cel = null; }
          if (tabel.rij) tabel.rijen.push(tabel.rij);
          tabel.rij = sluit ? null : [];
          break;
        case 'td': case 'th':
          if (!tabel) break;
          if (!tabel.rij) tabel.rij = [];
          if (tabel.cel !== null) { tabel.rij.push(tabel.cel.trim()); tabel.cel = null; }
          if (!sluit) {
            tabel.cel = '';
            if (tag === 'th' && !tabel.rijen.length) tabel.kop = true;
          }
          break;
        default:
          break;
      }
    }
    if (tabel) {
      if (tabel.cel !== null) tabel.rij.push(tabel.cel.trim());
      if (tabel.rij) tabel.rijen.push(tabel.rij);
      if (tabel.rijen.length) blokken.push({ soort: 'tabel', rijen: tabel.rijen, kop: tabel.kop });
    }
    sluitBlok();
    return blokken;
  }

  /** Gewone tekst: elke regel een alinea (lege regels blijven lege alinea's, zoals in Office). */
  function tekstNaarBlokken(tekst) {
    return String(tekst).replace(/\r\n?/g, '\n').split('\n').map(function (r) {
      return { soort: 'p', stukken: r ? [{ t: r }] : [] };
    });
  }

  // ------------------------------------------------------------------ IN DE EDITOR: Word
  // Let op: deze functies draaien in de editor (callCommand) en zien alleen Api en Asc.scope.
  var IN_EDITOR = {};

  IN_EDITOR.word_lezen = function () {
    var s = Asc.scope;
    try {
      var doc = Api.GetDocument();
      var ps = doc.GetAllParagraphs();
      var n = ps.length;
      var van = Math.max(1, s.van || 1);
      var tot = Math.min(n, Math.max(van, s.tot || n));
      var regels = [];
      var tekens = 0;
      for (var i = van; i <= tot; i++) {
        var p = ps[i - 1];
        var soort = '';
        var stijl = '';
        try { var st = p.GetParaPr().GetStyle(); stijl = st ? String(st.GetName() || '') : ''; } catch (e) { stijl = ''; }
        var m = /^(?:heading|kop)\s*(\d)/i.exec(stijl);
        if (m) soort = 'Kop ' + m[1];
        else if (/^(title|titel)$/i.test(stijl)) soort = 'Titel';
        else if (/^(subtitle|ondertitel)$/i.test(stijl)) soort = 'Ondertitel';
        try { if (p.GetNumbering()) soort = soort ? soort + ', opsomming' : 'opsomming'; } catch (e) { /* geen nummering */ }
        try { if (p.GetParentTable()) soort = soort ? soort + ', in tabel' : 'in tabel'; } catch (e) { /* niet in tabel */ }
        var t = String(p.GetText({ Numbering: false }) || '').replace(/[\r\n\t]+$/, '').replace(/\r\n?|\n/g, ' ↵ ');
        var regel = '[' + i + ']' + (soort ? ' (' + soort + ')' : '') + (t ? ' ' + t : '');
        tekens += regel.length + 1;
        if (tekens > s.max) {
          regels.push('[… gestopt na alinea ' + (i - 1) + ' van ' + n + ': lees verder met van_alinea ' + i + ']');
          break;
        }
        regels.push(regel);
      }
      var bijhouden = doc.IsTrackRevisions();
      var kop = 'Document met ' + n + " alinea's" + (van > 1 || tot < n ? ' (hier alinea ' + van + ' t/m ' + tot + ')' : '') +
        '; Wijzigingen bijhouden: ' + (bijhouden ? 'aan' : 'uit');
      var open = 0;
      try {
        var rapport = doc.GetReviewReport();
        for (var naam in rapport) if (Object.prototype.hasOwnProperty.call(rapport, naam) && Array.isArray(rapport[naam])) open += rapport[naam].length;
      } catch (e) { open = 0; }
      if (open) kop += '; ' + open + ' wijziging(en) wachten op accepteren of weigeren (tekst die als verwijderd gemarkeerd is, staat hieronder nog gewoon in de alinea)';
      var uit = kop + '\n' + regels.join('\n');
      var tabellen = doc.GetAllTables();
      if (tabellen.length && tekens < s.max) {
        var tt = ['', 'Tabellen (' + tabellen.length + '):'];
        for (var k = 0; k < tabellen.length; k++) {
          var tab = tabellen[k];
          var rijen = tab.GetRowsCount();
          tt.push('Tabel ' + (k + 1) + ' (' + rijen + ' rijen):');
          for (var r = 0; r < Math.min(rijen, s.tabelRijen); r++) {
            var rij = tab.GetRow(r);
            var cellen = [];
            for (var c = 0; c < rij.GetCellsCount(); c++) {
              cellen.push(String(rij.GetCell(c).GetContent().GetText() || '').replace(/[\r\n\t]+$/, '').replace(/\s*[\r\n]+\s*/g, ' '));
            }
            tt.push('| ' + cellen.join(' | ') + ' |');
          }
          if (rijen > s.tabelRijen) tt.push('[… nog ' + (rijen - s.tabelRijen) + ' rijen]');
        }
        uit += '\n' + tt.join('\n');
      }
      return { ok: true, tekst: uit };
    } catch (e) {
      return { fout: 'Lezen lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.word_selectie_lezen = function () {
    try {
      var r = Api.GetDocument().GetRangeBySelect();
      var t = r ? String(r.GetText() || '') : '';
      if (!t.replace(/[\r\n\s]+/g, '')) return { ok: true, tekst: 'Er is niets geselecteerd (alleen de cursor staat ergens). Vraag de gebruiker tekst te selecteren, of lees het document met word_lezen.' };
      return { ok: true, tekst: 'Geselecteerde tekst (' + t.length + ' tekens):\n' + t.slice(0, Asc.scope.max) };
    } catch (e) {
      return { fout: 'De selectie lezen lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.word_invoegen = function () {
    var s = Asc.scope;
    var doc = Api.GetDocument();
    var notitie = '';
    var ai = false;
    try {
      if (s.zetBijhoudenAan && !doc.IsTrackRevisions()) {
        doc.SetTrackRevisions(true);
        notitie = 'Wijzigingen bijhouden staat nu aan: iedereen ziet elke wijziging van de AI (op naam van ' + s.auteur + ') en kan die accepteren of weigeren (Samenwerking → Wijzigingen bijhouden).';
      }
      // Alleen de lege begin-alinea van een nieuw document: die vullen we liever dan er iets na te zetten.
      var leeg = doc.GetElementsCount() === 1 && doc.GetElement(0).GetClassType() === 'paragraph' &&
        !String(doc.GetElement(0).GetText() || '').replace(/[\r\n\s]+/g, '');

      // Vóór het opbouwen al op naam van AIDG: nieuwe tekst krijgt zijn auteur bij het aanmaken.
      ai = doc.IsTrackRevisions() && typeof doc.SetAssistantTrackRevisions === 'function';
      if (ai) doc.SetAssistantTrackRevisions(true, s.auteur);

      // Elementen opbouwen (alinea's, opsommingen, tabellen) uit de blokken van de plugin.
      var lijsten = {};
      var stijlen = {};
      var stijl = function (naam) {
        if (!(naam in stijlen)) { try { stijlen[naam] = doc.GetStyle(naam); } catch (e) { stijlen[naam] = null; } }
        return stijlen[naam];
      };
      var elementen = [];
      for (var i = 0; i < s.blokken.length; i++) {
        var b = s.blokken[i];
        if (b.soort === 'tabel') {
          var kolommen = 1;
          for (var r = 0; r < b.rijen.length; r++) kolommen = Math.max(kolommen, b.rijen[r].length);
          var t = Api.CreateTable(kolommen, b.rijen.length);
          var raster = stijl('Table Grid');
          if (raster) { try { t.SetStyle(raster); } catch (e) { /* standaardstijl */ } }
          try { t.SetWidth('percent', 100); } catch (e) { /* standaardbreedte */ }
          for (var rr = 0; rr < b.rijen.length; rr++) {
            for (var kk = 0; kk < b.rijen[rr].length; kk++) {
              var cp = t.GetCell(rr, kk).GetContent().GetElement(0);
              var celRun = Api.CreateRun();
              celRun.AddText(String(b.rijen[rr][kk]));
              if (b.kop && rr === 0) celRun.SetBold(true);
              cp.AddElement(celRun);
            }
          }
          elementen.push(t);
          continue;
        }
        var p = Api.CreateParagraph();
        if (/^h[1-3]$/.test(b.soort)) {
          var kopStijl = stijl('Heading ' + b.soort.charAt(1));
          if (kopStijl) p.SetStyle(kopStijl);
        }
        if (b.soort === 'li') {
          var sleutel = b.lijst + (b.genummerd ? '#' : '*');
          if (!lijsten[sleutel]) lijsten[sleutel] = doc.CreateNumbering(b.genummerd ? 'numbered' : 'bullet');
          p.SetNumbering(lijsten[sleutel].GetLevel(Math.min(8, b.niveau || 0)));
        }
        for (var j = 0; j < b.stukken.length; j++) {
          var st = b.stukken[j];
          var run = Api.CreateRun();
          if (st.br) run.AddLineBreak();
          else {
            run.AddText(st.t);
            if (st.b) run.SetBold(true);
            if (st.i) run.SetItalic(true);
            if (st.u) run.SetUnderline(true);
          }
          p.AddElement(run);
        }
        elementen.push(p);
      }
      if (!elementen.length) return { fout: 'Er is niets om in te voegen (lege tekst).' };

      var waar;
      if (s.positie === 'selectie') {
        var inline = elementen.length === 1 && elementen[0].GetClassType() === 'paragraph';
        if (!doc.InsertContent(elementen, inline)) return { fout: 'Invoegen op de plek van de cursor lukte niet (staat de cursor in het document?). Probeer positie einde of na_alinea.' };
        waar = 'op de plek van de selectie (of de cursor)';
      } else if (s.positie === 'begin') {
        for (var a = 0; a < elementen.length; a++) doc.AddElement(a, elementen[a]);
        if (leeg) doc.GetElement(elementen.length).Delete();
        waar = 'aan het begin van het document';
      } else if (s.positie === 'na_alinea') {
        var ps = doc.GetAllParagraphs();
        if (s.alinea < 1 || s.alinea > ps.length) return { fout: 'Alinea ' + s.alinea + ' bestaat niet: het document heeft ' + ps.length + " alinea's. Lees het opnieuw met word_lezen." };
        var doel = ps[s.alinea - 1];
        var inTabel = Boolean(doel.GetParentTable());
        var heeftTabel = elementen.some(function (el) { return el.GetClassType() === 'table'; });
        if (inTabel || heeftTabel) {
          if (inTabel && heeftTabel) return { fout: 'Een tabel kan niet in een tabelcel worden ingevoegd. Kies een alinea buiten de tabel.' };
          if (heeftTabel) {
            var pos = doel.GetPosInParent();
            for (var q = 0; q < elementen.length; q++) doc.AddElement(pos + 1 + q, elementen[q]);
          } else {
            var na = doel;
            for (var w = 0; w < elementen.length; w++) na = na.InsertParagraph(elementen[w], 'after', true);
          }
        } else {
          var vorige = doel;
          for (var x = 0; x < elementen.length; x++) vorige = vorige.InsertParagraph(elementen[x], 'after', true);
        }
        waar = 'na alinea ' + s.alinea;
      } else {
        for (var e2 = 0; e2 < elementen.length; e2++) doc.Push(elementen[e2]);
        if (leeg) doc.GetElement(0).Delete();
        waar = 'aan het eind van het document';
      }
      return { ok: true, tekst: 'Ingevoegd ' + waar + ' (' + elementen.length + ' alinea\'s/tabellen).' + (notitie ? '\n' + notitie : '') };
    } catch (e) {
      return { fout: 'Invoegen lukte niet: ' + (e && e.message) };
    } finally {
      if (ai) { try { doc.SetAssistantTrackRevisions(false); } catch (e) { /* laat staan */ } }
    }
  };

  IN_EDITOR.word_vervangen = function () {
    var s = Asc.scope;
    var doc = Api.GetDocument();
    var notitie = '';
    var ai = false;
    try {
      if (s.zetBijhoudenAan && !doc.IsTrackRevisions()) {
        doc.SetTrackRevisions(true);
        notitie = 'Wijzigingen bijhouden staat nu aan: iedereen ziet elke wijziging van de AI (op naam van ' + s.auteur + ') en kan die accepteren of weigeren.';
      }
      ai = doc.IsTrackRevisions() && typeof doc.SetAssistantTrackRevisions === 'function';
      var uit;
      if (s.alinea) {
        var ps = doc.GetAllParagraphs();
        if (s.alinea < 1 || s.alinea > ps.length) return { fout: 'Alinea ' + s.alinea + ' bestaat niet: het document heeft ' + ps.length + " alinea's. Lees het opnieuw met word_lezen." };
        var p = ps[s.alinea - 1];
        var regels = String(s.door).replace(/\r\n?/g, '\n').split('\n');
        if (ai) doc.SetAssistantTrackRevisions(true, s.auteur);
        // De nieuwe tekst achteraan in dezelfde alinea (alinea-opmaak blijft), daarna de oude stukken (runs)
        // als verwijderd markeren, van achter naar voren. Tekens tellen kan niet: een bereik telt ook de
        // grenzen tussen runs mee. Met Wijzigingen bijhouden ziet iedereen wat er wegging.
        var oude = [];
        for (var e = 0; e < p.GetElementsCount(); e++) {
          var el = p.GetElement(e);
          if (el && typeof el.GetRange === 'function') {
            var r = el.GetRange();
            if (r) oude.push(r);
          }
        }
        if (regels[0]) {
          var nieuw = Api.CreateRun();
          nieuw.AddText(regels[0]);
          p.AddElement(nieuw);
        }
        for (var o = oude.length - 1; o >= 0; o--) oude[o].Delete();
        var vorige = p;
        for (var i = 1; i < regels.length; i++) vorige = vorige.InsertParagraph(regels[i], 'after', true);
        uit = 'Alinea ' + s.alinea + ' vervangen' + (regels.length > 1 ? ' (nu ' + regels.length + " alinea's)" : '') + '.';
      } else {
        var gevonden = doc.Search(s.zoeken, s.hoofdletters);
        var n = gevonden.length;
        if (!n) return { fout: 'Niet gevonden: "' + s.zoeken.slice(0, 80) + '" staat niet in het document (let op spaties, leestekens en hoofdletters). Lees het document opnieuw.' };
        var door = String(s.door).replace(/\r\n?|\n/g, '\u000b'); // regeleinde binnen de alinea
        if (ai) doc.SetAssistantTrackRevisions(true, s.auteur);
        if (s.alle) {
          doc.SearchAndReplace({ searchString: s.zoeken, replaceString: door, matchCase: s.hoofdletters });
          uit = n + ' keer vervangen.';
        } else {
          gevonden[0].Delete();
          if (door) gevonden[0].AddText(door, 'before');
          uit = '1 keer vervangen' + (n > 1 ? ' (de eerste van de ' + n + ' plekken)' : '') + '.';
        }
      }
      return { ok: true, tekst: uit + (notitie ? '\n' + notitie : '') };
    } catch (e) {
      return { fout: 'Vervangen lukte niet: ' + (e && e.message) };
    } finally {
      if (ai) { try { doc.SetAssistantTrackRevisions(false); } catch (e) { /* laat staan */ } }
    }
  };

  IN_EDITOR.word_opmerking = function () {
    var s = Asc.scope;
    try {
      var doc = Api.GetDocument();
      var bereik;
      if (s.zoeken) {
        var gevonden = doc.Search(s.zoeken, false);
        if (!gevonden.length) return { fout: 'Niet gevonden: "' + s.zoeken.slice(0, 80) + '" staat niet in het document.' };
        bereik = gevonden[0];
      } else {
        bereik = doc.GetRangeBySelect();
        if (!bereik || !String(bereik.GetText() || '').replace(/[\r\n\s]+/g, '')) {
          return { fout: 'Er is niets geselecteerd. Geef zoeken mee (de tekst waar de opmerking bij hoort) of vraag de gebruiker tekst te selecteren.' };
        }
      }
      if (!bereik.AddComment(s.opmerking, s.auteur)) return { fout: 'De opmerking plaatsen lukte niet (is het document alleen-lezen?).' };
      return { ok: true, tekst: 'Opmerking geplaatst' + (s.zoeken ? ' bij "' + s.zoeken.slice(0, 80) + '"' : ' bij de selectie') + ' (op naam van ' + s.auteur + ').' };
    } catch (e) {
      return { fout: 'De opmerking plaatsen lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.word_wijzigingen_bijhouden = function () {
    var s = Asc.scope;
    try {
      var doc = Api.GetDocument();
      if (typeof s.aan === 'boolean') doc.SetTrackRevisions(s.aan);
      var aan = doc.IsTrackRevisions();
      return { ok: true, tekst: 'Wijzigingen bijhouden staat ' + (aan ? 'aan' : 'uit') + '.' + (aan ? ' Iedereen in het document kan wijzigingen accepteren of weigeren via Samenwerking → Wijzigingen bijhouden.' : '') };
    } catch (e) {
      return { fout: 'Wijzigingen bijhouden aan- of uitzetten lukte niet: ' + (e && e.message) };
    }
  };

  // ------------------------------------------------------------------ IN DE EDITOR: Excel
  /** Een bereik (of de selectie) lezen: zichtbare tekst en formules per cel; geeft {fout} of de gegevens. */
  IN_EDITOR.excel_lezen = function () {
    var s = Asc.scope;
    try {
      var blad;
      var adres = s.adres;
      if (s.selectie) {
        var sel = Api.GetSelection();
        if (!sel) return { fout: 'Er is niets geselecteerd.' };
        blad = sel.GetWorksheet();
        adres = String(sel.GetAddress(false, false) || '').replace(/\$/g, '');
      } else {
        blad = s.blad ? Api.GetSheet(s.blad) : Api.GetActiveSheet();
        if (!blad) return { fout: 'Werkblad "' + s.blad + '" bestaat niet. Kijk met excel_werkbladen welke bladen er zijn.' };
        if (!adres) {
          var gebruikt = blad.GetUsedRange();
          adres = gebruikt ? String(gebruikt.GetAddress(false, false) || 'A1').replace(/\$/g, '') : 'A1';
        }
      }
      adres = adres.split(',')[0];
      var m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(adres);
      if (!m) return { fout: 'Ongeldig bereik: ' + adres + '. Gebruik bijvoorbeeld A1:D10.' };
      var kol = function (l) { var k = 0; for (var j = 0; j < l.length; j++) k = k * 26 + (l.toUpperCase().charCodeAt(j) - 64); return k; };
      var r1 = Number(m[2]);
      var k1 = kol(m[1]);
      var r2 = m[3] ? Number(m[4]) : r1;
      var k2 = m[3] ? kol(m[3]) : k1;
      var rijen = r2 - r1 + 1;
      var kolommen = k2 - k1 + 1;
      var leesKolommen = Math.min(kolommen, s.maxKolommen);
      var leesRijen = Math.min(rijen, Math.max(1, Math.floor(s.maxCellen / leesKolommen)));
      var waarden = [];
      var formules = [];
      for (var r = 0; r < leesRijen; r++) {
        var w = [];
        var f = [];
        for (var c = 0; c < leesKolommen; c++) {
          var cel = blad.GetRangeByNumber(r1 - 1 + r, k1 - 1 + c);
          var tekst = cel.GetText();
          var formule = cel.GetFormula();
          w.push(tekst === undefined || tekst === null ? '' : String(tekst));
          f.push(typeof formule === 'string' && /^=/.test(formule) ? '=' + formule.replace(/^=\s*/, '') : null);
        }
        waarden.push(w);
        formules.push(f);
      }
      return { ok: true, data: { blad: blad.GetName(), adres: adres, rij: r1, kolom: k1, waarden: waarden, formules: formules, totaalRijen: rijen, totaalKolommen: kolommen } };
    } catch (e) {
      return { fout: 'Lezen lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.excel_werkbladen = function () {
    try {
      var actief = Api.GetActiveSheet().GetName();
      var regels = Api.GetSheets().map(function (b) {
        var delen = [b.GetName() + (b.GetName() === actief ? ' (actief)' : '') + (b.GetVisible() === false ? ' (verborgen)' : '')];
        var u = b.GetUsedRange();
        var adres = u ? String(u.GetAddress(false, false) || '').replace(/\$/g, '') : '';
        var leeg = !adres || (adres === 'A1' && !String(b.GetRange('A1').GetText() || ''));
        delen.push(leeg ? 'leeg' : 'gebruikt: ' + adres);
        return '- ' + delen.join('; ');
      });
      return { ok: true, tekst: 'Werkbladen (' + regels.length + '):\n' + regels.join('\n') };
    } catch (e) {
      return { fout: 'De werkbladen lezen lukte niet: ' + (e && e.message) };
    }
  };

  /** Scheidingstekens van de editor (Nederlands: ; tussen argumenten en een komma als decimaalteken). */
  IN_EDITOR.excel_scheidingstekens = function () {
    try {
      var fs = AscCommon.FormulaSeparators;
      return { ok: true, data: { argument: fs.functionArgumentSeparator || ',', decimaal: fs.digitSeparator || '.', kolom: fs.arrayColSeparator || ',' } };
    } catch (e) {
      return { ok: true, data: { argument: ',', decimaal: '.', kolom: ',' } };
    }
  };

  IN_EDITOR.excel_schrijven = function () {
    var s = Asc.scope;
    try {
      var blad = s.blad ? Api.GetSheet(s.blad) : Api.GetActiveSheet();
      if (!blad) return { fout: 'Werkblad "' + s.blad + '" bestaat niet. Kijk met excel_werkbladen welke bladen er zijn.' };
      var overschreven = 0;
      var formules = false;
      for (var r = 0; r < s.waarden.length; r++) {
        for (var k = 0; k < s.waarden[r].length; k++) {
          var v = s.waarden[r][k];
          if (v === null || v === undefined) continue; // deze cel niet aanraken (al omgezet naar tekst in de editor-taal)
          var cel = blad.GetRangeByNumber(s.rij - 1 + r, s.kolom - 1 + k);
          var oud = cel.GetValue();
          if (oud !== '' && oud !== null && oud !== undefined) overschreven++;
          if (v.charAt(0) === '=') formules = true;
          cel.SetValue(v);
        }
      }
      return { ok: true, data: { blad: blad.GetName(), overschreven: overschreven, formules: formules } };
    } catch (e) {
      return { fout: 'Schrijven lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.excel_opmaak = function () {
    var s = Asc.scope;
    try {
      var blad = s.blad ? Api.GetSheet(s.blad) : Api.GetActiveSheet();
      if (!blad) return { fout: 'Werkblad "' + s.blad + '" bestaat niet.' };
      var bereik = blad.GetRange(s.adres);
      if (!bereik) return { fout: 'Ongeldig bereik: ' + s.adres + '.' };
      var gedaan = [];
      if (typeof s.vet === 'boolean') { bereik.SetBold(s.vet); gedaan.push(s.vet ? 'vet' : 'niet vet'); }
      if (typeof s.cursief === 'boolean') { bereik.SetItalic(s.cursief); gedaan.push(s.cursief ? 'cursief' : 'niet cursief'); }
      if (s.tekstkleur) { bereik.SetFontColor(Api.CreateColorFromRGB(s.tekstkleur[0], s.tekstkleur[1], s.tekstkleur[2])); gedaan.push('tekstkleur'); }
      if (s.geenOpvulling) { bereik.SetFillColor('No Fill'); gedaan.push('geen opvulkleur'); }
      else if (s.opvulkleur) { bereik.SetFillColor(Api.CreateColorFromRGB(s.opvulkleur[0], s.opvulkleur[1], s.opvulkleur[2])); gedaan.push('opvulkleur'); }
      if (s.getalnotatie) { bereik.SetNumberFormat(s.getalnotatie); gedaan.push('getalnotatie ' + s.getalnotatie); }
      if (s.kolommenPassend) { bereik.AutoFit(false, true); gedaan.push('kolombreedte passend'); }
      return { ok: true, tekst: 'Opgemaakt (' + blad.GetName() + '!' + s.adres + '): ' + gedaan.join(', ') + '.' };
    } catch (e) {
      return { fout: 'Opmaken lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.excel_tabel_maken = function () {
    var s = Asc.scope;
    try {
      var blad = s.blad ? Api.GetSheet(s.blad) : Api.GetActiveSheet();
      if (!blad) return { fout: 'Werkblad "' + s.blad + '" bestaat niet.' };
      blad.FormatAsTable(s.adres);
      return { ok: true, tekst: 'Tabel gemaakt van ' + blad.GetName() + '!' + s.adres + ' (eerste rij als kopjes, met filterknoppen).' };
    } catch (e) {
      return { fout: 'Een tabel maken lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.excel_grafiek_maken = function () {
    var s = Asc.scope;
    try {
      var blad = s.blad ? Api.GetSheet(s.blad) : Api.GetActiveSheet();
      if (!blad) return { fout: 'Werkblad "' + s.blad + '" bestaat niet.' };
      var naam = blad.GetName().replace(/'/g, "''");
      var cm = 360000; // EMU per centimeter
      var g = blad.AddChart("'" + naam + "'!" + s.absoluut, s.inRijen, s.soort, 2, 15 * cm, 8 * cm, s.naastKolom, 0, s.vanafRij, 0);
      if (!g) return { fout: 'Een grafiek maken lukte niet (is het blad beveiligd?).' };
      if (s.titel) { try { g.SetTitle(s.titel, 13); } catch (e) { /* zonder titel */ } }
      return { ok: true, tekst: 'Grafiek gemaakt van ' + blad.GetName() + '!' + s.adres + (s.titel ? ' met titel "' + s.titel + '"' : '') + '. De gebruiker kan hem verslepen en aanpassen.' };
    } catch (e) {
      return { fout: 'Een grafiek maken lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.excel_werkblad_toevoegen = function () {
    var s = Asc.scope;
    try {
      if (Api.GetSheet(s.naam)) return { fout: 'Er is al een werkblad "' + s.naam + '".' };
      var vorige = Api.GetActiveSheet();
      Api.AddSheet(s.naam);
      if (!s.activeren && vorige) vorige.SetActive();
      return { ok: true, tekst: 'Werkblad "' + s.naam + '" toegevoegd' + (s.activeren ? ' en geopend' : '') + '.' };
    } catch (e) {
      return { fout: 'Een werkblad toevoegen lukte niet: ' + (e && e.message) };
    }
  };

  // ------------------------------------------------------------------ IN DE EDITOR: PowerPoint
  IN_EDITOR.powerpoint_lezen = function () {
    var s = Asc.scope;
    try {
      var pres = Api.GetPresentation();
      var aantal = pres.GetSlidesCount();
      var regels = ['Presentatie met ' + aantal + " dia's:"];
      var tekens = 0;
      for (var d = 0; d < aantal; d++) {
        regels.push('Dia ' + (d + 1) + ':');
        var vormen = pres.GetSlideByIndex(d).GetAllShapes();
        var iets = false;
        for (var v = 0; v < vormen.length; v++) {
          var inhoud = vormen[v].GetDocContent();
          if (!inhoud) continue;
          var t = String(inhoud.GetText() || '').replace(/\r\n?/g, '\n').replace(/^\n+|\n+$/g, '');
          if (!t.trim()) continue;
          var ph = null;
          try { ph = vormen[v].GetPlaceholder(); } catch (e) { ph = null; }
          var soort = ph ? String(ph.GetType() || '') : '';
          var naam = '';
          try { naam = vormen[v].GetName ? String(vormen[v].GetName() || '') : ''; } catch (e) { naam = ''; }
          var label = /^(title|ctrTitle)$/i.test(soort) ? 'titel' : naam || soort || 'vorm ' + (v + 1);
          var r = '  - [' + label + '] ' + t.replace(/\n/g, '\n    ');
          tekens += r.length;
          regels.push(r);
          iets = true;
        }
        if (!iets) regels.push('  (geen tekst)');
        if (tekens > s.max) {
          regels.push('[… gestopt na dia ' + (d + 1) + ' van ' + aantal + ']');
          break;
        }
      }
      return { ok: true, tekst: regels.join('\n') };
    } catch (e) {
      return { fout: "De dia's lezen lukte niet: " + (e && e.message) };
    }
  };

  IN_EDITOR.powerpoint_dia_toevoegen = function () {
    var s = Asc.scope;
    try {
      var pres = Api.GetPresentation();
      var meester = pres.GetMaster(0);
      var indeling = meester ? meester.GetLayoutByType('obj') : null;
      var dia = Api.CreateSlide();
      if (indeling) dia.ApplyLayout(indeling);
      pres.AddSlide(dia);
      var vormen = dia.GetAllShapes();
      var soortVan = function (v) { try { var ph = v.GetPlaceholder(); return ph ? String(ph.GetType() || '') : ''; } catch (e) { return ''; } };
      var titel = null;
      var lijf = null;
      for (var i = 0; i < vormen.length; i++) {
        var st = soortVan(vormen[i]);
        if (!titel && /^(title|ctrTitle)$/i.test(st)) titel = vormen[i];
        else if (!lijf && st && !/^(title|ctrTitle|subTitle|dt|date|ftr|footer|sldNum)$/i.test(st)) lijf = vormen[i];
      }
      var vul = function (vorm, regels) {
        var inhoud = vorm.GetDocContent();
        inhoud.RemoveAllElements();
        for (var r = 0; r < regels.length; r++) {
          var p = r === 0 && inhoud.GetElementsCount() ? inhoud.GetElement(0) : null;
          if (!p) { p = Api.CreateParagraph(); inhoud.Push(p); }
          p.AddText(regels[r]);
        }
      };
      var noot = '';
      if (titel) vul(titel, [s.titel]);
      else noot += ' De titel kon niet op deze dia (geen titelvak in de indeling).';
      if (s.punten.length) {
        if (lijf) vul(lijf, s.punten);
        else noot += ' De opsommingstekens konden niet op deze dia (geen tekstvak in de indeling).';
      }
      return { ok: true, tekst: 'Dia ' + pres.GetSlidesCount() + ' toegevoegd' + (indeling ? ' (indeling "Titel en object")' : '') + ' met titel "' + s.titel + '"' + (s.punten.length ? ' en ' + s.punten.length + ' punten' : '') + '.' + noot };
    } catch (e) {
      return { fout: 'Een dia toevoegen lukte niet: ' + (e && e.message) };
    }
  };

  IN_EDITOR.powerpoint_tekst_vervangen = function () {
    var s = Asc.scope;
    try {
      var pres = Api.GetPresentation();
      var aantal = pres.GetSlidesCount();
      if (s.dia && (s.dia < 1 || s.dia > aantal)) return { fout: 'Dia ' + s.dia + ' bestaat niet: de presentatie heeft ' + aantal + " dia's." };
      var keer = 0;
      var alineas = 0;
      for (var d = 0; d < aantal; d++) {
        if (s.dia && d !== s.dia - 1) continue;
        var vormen = pres.GetSlideByIndex(d).GetAllShapes();
        for (var v = 0; v < vormen.length; v++) {
          var inhoud = vormen[v].GetDocContent();
          if (!inhoud) continue;
          var n = inhoud.GetElementsCount();
          for (var i = 0; i < n; i++) {
            var p = inhoud.GetElement(i);
            if (!p || p.GetClassType() !== 'paragraph') continue;
            var t = String(p.GetText() || '').replace(/[\r\n]+$/, '');
            var stukken = t.split(s.zoeken);
            if (stukken.length < 2) continue;
            p.RemoveAllElements();
            p.AddText(stukken.join(s.door));
            keer += stukken.length - 1;
            alineas++;
          }
        }
      }
      if (!keer) return { fout: 'Niet gevonden: "' + s.zoeken.slice(0, 80) + '" staat niet in ' + (s.dia ? 'dia ' + s.dia : 'de presentatie') + ' (hoofdlettergevoelig).' };
      return { ok: true, tekst: keer + ' keer vervangen in ' + alineas + ' alinea(\'s). Let op: een alinea waarin iets vervangen is, heeft nu één opmaak (die van het begin).' };
    } catch (e) {
      return { fout: 'Vervangen lukte niet: ' + (e && e.message) };
    }
  };

  // ------------------------------------------------------------------ uitvoerder (in de plugin)
  /**
   * @param o.roep       (functie, scope) → Promise<resultaat van de functie in de editor> (callCommand)
   * @param o.programma  'Word' | 'Excel' | 'PowerPoint'
   * @returns async (tool, args) → tekst (of gooit een fout met uitleg)
   */
  function maakUitvoerder(o) {
    var roep = o.roep;
    var staat = { bijhoudenGeregeld: false };

    async function inEditor(naam, scope) {
      var uit = await roep(IN_EDITOR[naam], scope);
      if (!uit || typeof uit !== 'object') {
        throw melding('De editor gaf geen antwoord. Misschien is het document alleen te lezen, of staat er een dialoogvenster open. Vraag de gebruiker dat te controleren en probeer het opnieuw.');
      }
      if (uit.fout) throw melding(String(uit.fout));
      return uit;
    }
    function moet(programma) {
      if (o.programma !== programma) throw melding('Dit venster is geen ' + programma + '-document maar ' + o.programma + '.');
    }
    function bereikArg(bereik, bladArg) {
      var sb = splitsBereik(bereik);
      var a = leesAdres(sb.adres);
      if (!a) throw melding('Ongeldig bereik: "' + String(bereik || '').slice(0, 40) + '". Gebruik bijvoorbeeld A1:D10 (of Blad1!A1:D10).');
      return { blad: sb.blad || (bladArg ? String(bladArg) : ''), adres: adresVan(a.rij, a.kolom, a.rij2, a.kolom2), a: a };
    }
    function scheidingstekens() {
      if (staat.sep) return Promise.resolve(staat.sep);
      return inEditor('excel_scheidingstekens', {}).then(function (u) { staat.sep = u.data; return u.data; });
    }
    function bijhoudenScope() {
      var zet = !staat.bijhoudenGeregeld;
      staat.bijhoudenGeregeld = true;
      return zet;
    }

    var tools = {
      word_lezen: function (a) {
        moet('Word');
        return inEditor('word_lezen', { van: geheel(a.van_alinea, 1), tot: geheel(a.tot_alinea, 0), max: MAX_TEKENS, tabelRijen: TABEL_RIJEN }).then(function (u) { return u.tekst; });
      },
      word_selectie_lezen: function () {
        moet('Word');
        return inEditor('word_selectie_lezen', { max: MAX_TEKENS }).then(function (u) { return u.tekst; });
      },
      word_invoegen: function (a) {
        moet('Word');
        var tekst = typeof a.tekst === 'string' ? a.tekst : null;
        var html = typeof a.html === 'string' ? a.html : null;
        if ((tekst === null) === (html === null)) throw melding('Geef tekst óf html (één van de twee).');
        var positie = a.positie || 'einde';
        if (['selectie', 'begin', 'einde', 'na_alinea'].indexOf(positie) < 0) throw melding('Onbekende positie: ' + positie + '. Kies selectie, begin, einde of na_alinea.');
        if (positie === 'na_alinea' && !a.alinea) throw melding('Bij na_alinea hoort een alineanummer (alinea), uit word_lezen.');
        var blokken = html !== null ? htmlNaarBlokken(html) : tekstNaarBlokken(tekst);
        if (!blokken.length) throw melding('Er is niets om in te voegen (lege tekst).');
        return inEditor('word_invoegen', { blokken: blokken, positie: positie, alinea: geheel(a.alinea, 0), zetBijhoudenAan: bijhoudenScope(), auteur: AUTEUR }).then(function (u) { return u.tekst; });
      },
      word_vervangen: function (a) {
        moet('Word');
        var door = typeof a.vervangen_door === 'string' ? a.vervangen_door : null;
        if (door === null) throw melding('vervangen_door ontbreekt.');
        if (!a.alinea && !a.zoeken) throw melding('Geef zoeken (tekst om te vervangen) of alinea (nummer van een hele alinea).');
        if (a.zoeken && String(a.zoeken).length > 255) throw melding('De zoektekst is te lang (hooguit 255 tekens). Vervang dan de hele alinea met alinea.');
        if (!a.alinea && a.hele_woorden === true) {
          throw melding('Alleen hele woorden zoeken kan in AIDG Werkplek niet. Zoek met een langere, unieke tekst (bijvoorbeeld met een woord ervoor of erna), of vervang de hele alinea.');
        }
        return inEditor('word_vervangen', {
          alinea: a.alinea ? geheel(a.alinea, 0) : 0,
          zoeken: a.alinea ? '' : String(a.zoeken),
          door: door,
          alle: a.alle !== false,
          hoofdletters: a.hoofdlettergevoelig === true,
          zetBijhoudenAan: bijhoudenScope(),
          auteur: AUTEUR,
        }).then(function (u) { return u.tekst; });
      },
      word_opmerking: function (a) {
        moet('Word');
        var opm = String(a.opmerking || '').trim();
        if (!opm) throw melding('De opmerking is leeg.');
        return inEditor('word_opmerking', { opmerking: opm, zoeken: a.zoeken ? String(a.zoeken).slice(0, 255) : '', auteur: AUTEUR }).then(function (u) { return u.tekst; });
      },
      word_wijzigingen_bijhouden: function (a) {
        moet('Word');
        if (typeof a.aan === 'boolean') staat.bijhoudenGeregeld = true;
        return inEditor('word_wijzigingen_bijhouden', { aan: typeof a.aan === 'boolean' ? a.aan : null }).then(function (u) { return u.tekst; });
      },

      excel_werkbladen: function () {
        moet('Excel');
        return inEditor('excel_werkbladen', {}).then(function (u) { return u.tekst; });
      },
      excel_lezen: function (a) {
        moet('Excel');
        var b = a.bereik ? bereikArg(a.bereik, a.blad) : { blad: a.blad ? String(a.blad) : '', adres: '' };
        return inEditor('excel_lezen', { blad: b.blad, adres: b.adres, maxCellen: MAX_CELLEN, maxKolommen: MAX_KOLOMMEN }).then(function (u) { return bereikTekst(u.data); });
      },
      excel_selectie_lezen: function () {
        moet('Excel');
        return inEditor('excel_lezen', { selectie: true, maxCellen: MAX_CELLEN, maxKolommen: MAX_KOLOMMEN }).then(function (u) { return 'Selectie van de gebruiker: ' + bereikTekst(u.data); });
      },
      excel_schrijven: function (a) {
        moet('Excel');
        var w = a.waarden;
        if (!Array.isArray(w) || !w.length || !w.every(Array.isArray)) throw melding('waarden moet een tabel zijn: een lijst van rijen, bv. [["Naam","Bedrag"],["Huur",950]].');
        var kolommen = Math.max.apply(null, w.map(function (r) { return r.length; }));
        if (!kolommen) throw melding('waarden is leeg.');
        if (w.length * kolommen > MAX_SCHRIJVEN) throw melding('Te veel cellen in één keer (' + w.length * kolommen + '; hooguit ' + MAX_SCHRIJVEN + '). Schrijf in delen.');
        var b = bereikArg(a.begincel, a.blad);
        return scheidingstekens().then(function (sep) {
          // Zoals de gebruiker het zou typen in de taal van de editor: 9,5 en =ALS(A1>1;2;0,5) in het Nederlands.
          var rijen = w.map(function (r) {
            return r.map(function (c) {
              if (c === null || c === undefined) return null;
              if (typeof c === 'number') return Number.isFinite(c) ? String(c).replace('.', sep.decimaal) : null;
              if (typeof c === 'boolean') return c ? 'TRUE' : 'FALSE';
              var t = typeof c === 'object' ? JSON.stringify(c) : String(c);
              return t.charAt(0) === '=' ? lokaleFormule(t, sep) : t;
            });
          });
          return inEditor('excel_schrijven', { blad: b.blad, rij: b.a.rij, kolom: b.a.kolom, waarden: rijen }).then(function (u) { return { u: u, rijen: rijen }; });
        }).then(function (x) {
          var u = x.u;
          var rijen = x.rijen;
          var d = u.data;
          var doel = adresVan(b.a.rij, b.a.kolom, b.a.rij + rijen.length - 1, b.a.kolom + kolommen - 1);
          return 'Geschreven naar ' + d.blad + '!' + doel + ' (' + rijen.length + ' × ' + kolommen + ' cellen' + (d.formules ? ', met formules' : '') + ').' +
            (d.overschreven ? ' ' + d.overschreven + ' cel(len) hadden al inhoud en zijn overschreven.' : '') +
            ' Iedereen in het bestand ziet het meteen; Ongedaan maken (Ctrl+Z) werkt.';
        });
      },
      excel_opmaak: function (a) {
        moet('Excel');
        var b = bereikArg(a.bereik, a.blad);
        var scope = { blad: b.blad, adres: b.adres };
        var iets = false;
        if (typeof a.vet === 'boolean') { scope.vet = a.vet; iets = true; }
        if (typeof a.cursief === 'boolean') { scope.cursief = a.cursief; iets = true; }
        if (a.tekstkleur) {
          scope.tekstkleur = rgb(a.tekstkleur);
          if (!scope.tekstkleur) throw melding('Onbekende tekstkleur "' + a.tekstkleur + '". Gebruik #RRGGBB.');
          iets = true;
        }
        if (a.opvulkleur) {
          if (/^(geen|none)$/i.test(String(a.opvulkleur))) scope.geenOpvulling = true;
          else {
            scope.opvulkleur = rgb(a.opvulkleur);
            if (!scope.opvulkleur) throw melding('Onbekende opvulkleur "' + a.opvulkleur + '". Gebruik #RRGGBB of "geen".');
          }
          iets = true;
        }
        if (a.getalnotatie) {
          if ((b.a.rij2 - b.a.rij + 1) * (b.a.kolom2 - b.a.kolom + 1) > MAX_SCHRIJVEN) throw melding('Het bereik is te groot voor een getalnotatie in één keer. Kies een kleiner bereik.');
          scope.getalnotatie = String(a.getalnotatie);
          iets = true;
        }
        if (a.kolommen_passend) { scope.kolommenPassend = true; iets = true; }
        if (!iets) throw melding('Geef minstens één opmaak: vet, cursief, tekstkleur, opvulkleur, getalnotatie of kolommen_passend.');
        return inEditor('excel_opmaak', scope).then(function (u) { return u.tekst; });
      },
      excel_tabel_maken: function (a) {
        moet('Excel');
        var b = bereikArg(a.bereik, a.blad);
        if (b.a.rij2 === b.a.rij) throw melding('Een tabel heeft minstens twee rijen nodig (kopjes en gegevens). Kies een groter bereik.');
        return inEditor('excel_tabel_maken', { blad: b.blad, adres: b.adres }).then(function (u) {
          var extra = [];
          if (a.kopteksten === false) extra.push('de eerste rij is in AIDG Werkplek altijd de kop');
          if (a.naam) extra.push('een eigen tabelnaam kan hier niet');
          if (a.stijl) extra.push('de stijl is de standaardstijl (een andere kan de gebruiker kiezen via Tabelinstellingen)');
          return u.tekst + (extra.length ? ' Let op: ' + extra.join('; ') + '.' : '');
        });
      },
      excel_grafiek_maken: function (a) {
        moet('Excel');
        var soort = GRAFIEK[a.soort || 'kolom'];
        if (!soort) throw melding('Onbekende soort grafiek. Kies kolom, staaf, lijn, taart, vlak of spreiding.');
        var b = bereikArg(a.bereik, a.blad);
        var rijen = b.a.rij2 - b.a.rij + 1;
        var kolommen = b.a.kolom2 - b.a.kolom + 1;
        var per = a.reeksen_per || 'automatisch';
        var inRijen = per === 'rijen' ? true : per === 'kolommen' ? false : kolommen > rijen;
        var absoluut = '$' + kolomLetters(b.a.kolom) + '$' + b.a.rij + ':$' + kolomLetters(b.a.kolom2) + '$' + b.a.rij2;
        return inEditor('excel_grafiek_maken', {
          blad: b.blad, adres: b.adres, absoluut: absoluut, soort: soort, inRijen: inRijen,
          titel: a.titel ? String(a.titel) : '', naastKolom: b.a.kolom2 + 1, vanafRij: b.a.rij - 1,
        }).then(function (u) { return u.tekst; });
      },
      excel_werkblad_toevoegen: function (a) {
        moet('Excel');
        var naam = String(a.naam || '').trim();
        if (!naam || /[\\/?*[\]:]/.test(naam) || naam.length > 31) throw melding('Ongeldige bladnaam: hooguit 31 tekens, zonder \\ / ? * [ ] :');
        return inEditor('excel_werkblad_toevoegen', { naam: naam, activeren: a.activeren !== false }).then(function (u) { return u.tekst; });
      },

      powerpoint_lezen: function () {
        moet('PowerPoint');
        return inEditor('powerpoint_lezen', { max: MAX_TEKENS }).then(function (u) { return u.tekst; });
      },
      powerpoint_dia_toevoegen: function (a) {
        moet('PowerPoint');
        var titel = String(a.titel || '').trim();
        if (!titel) throw melding('De titel ontbreekt.');
        var punten = Array.isArray(a.punten) ? a.punten.map(function (p) { return String(p); }) : [];
        return inEditor('powerpoint_dia_toevoegen', { titel: titel, punten: punten }).then(function (u) { return u.tekst; });
      },
      powerpoint_tekst_vervangen: function (a) {
        moet('PowerPoint');
        var zoek = String(a.zoeken || '');
        if (!zoek) throw melding('zoeken is leeg.');
        return inEditor('powerpoint_tekst_vervangen', { zoeken: zoek, door: String(a.vervangen_door == null ? '' : a.vervangen_door), dia: a.dia ? geheel(a.dia, 0) : 0 }).then(function (u) { return u.tekst; });
      },
    };

    return async function voerUit(tool, args) {
      var f = tools[tool];
      if (!f) throw melding('Deze opdracht kent AIDG Werkplek (nog) niet: ' + tool + '. Herlaad het document en probeer het opnieuw.');
      var uit = await f(args && typeof args === 'object' ? args : {});
      return String(uit);
    };
  }

  return {
    maakUitvoerder: maakUitvoerder,
    foutTekst: foutTekst,
    htmlNaarBlokken: htmlNaarBlokken,
    tekstNaarBlokken: tekstNaarBlokken,
    lokaleFormule: lokaleFormule,
    bereikTekst: bereikTekst,
    splitsBereik: splitsBereik,
    leesAdres: leesAdres,
    kolomLetters: kolomLetters,
    rgb: rgb,
    PROGRAMMA: PROGRAMMA,
    IN_EDITOR: IN_EDITOR,
    MAX_TEKENS: MAX_TEKENS,
  };
});
