/*
 * AIDG-koppeling voor de AI van AIDG Documenten (het AI-tabblad in de editors
 * en de AI-agent in het startscherm).
 *
 * De installatie zet dit script vóór de eigen scripts van de AI-pagina's
 * (index.html op de achtergrond, chat.html en settings.html als venster).
 * Het zet "AIDG" als aanbieder en model in de instellingen van de AI, langs
 * één van twee wegen:
 *
 *  1. Met de AIDG-app: de app zet per gebruiker een verbinding neer
 *     (data/sdkjs-plugins/aidg-koppeling/verbinding.json in het eigen profiel).
 *     De AI gaat dan via de app, met het eigen account en het gekozen model.
 *  2. Zonder app: "Inloggen met AIDG". De office maakt een eigen sleutel, de
 *     gebruiker keurt die goed in het portaal (aidg.nl/portal/office-koppelen.php)
 *     en de AI praat daarna rechtstreeks met de AIDG-gateway, op het eigen
 *     tegoed. Ontkoppelen kan in het portaal; dan vraagt de office opnieuw.
 *
 * Heeft de gebruiker zelf een andere aanbieder gekozen (eigen API-sleutel),
 * dan blijft die staan. Gaat er iets mis, dan doet dit script niets en werkt
 * de pagina zoals altijd.
 */
(function () {
  'use strict';
  var GATEWAY = 'https://privacy.aidg.nl/mistral/v1';
  var PORTAAL = 'https://aidg.nl/portal/office-koppelen.php';
  var TEGOED = 'https://aidg.nl/portal/topup.php';
  var ACCOUNT = 'aidg-office-account'; // {sleutel, email}
  var NIET_NU = 'aidg-office-niet-nu';
  var NAAM = 'AIDG';

  var pagina = String(window.location.pathname || '').split('/').pop();
  var venster = pagina === 'chat.html' || pagina === 'settings.html';

  var lees = function (k) {
    try {
      return JSON.parse(window.localStorage.getItem(k) || 'null');
    } catch (e) {
      return null;
    }
  };
  var schrijf = function (k, w) {
    try {
      if (w === null) window.localStorage.removeItem(k);
      else window.localStorage.setItem(k, JSON.stringify(w));
    } catch (e) {
      /* opslag dicht: dan maar zonder */
    }
  };

  /** AIDG in de lijst met aanbieders, en als huidige tenzij de gebruiker bewust iets anders koos. */
  function zetAanbieder(naam, baseUrl, sleutel, modelId, modelNaam) {
    var aanbieder = { name: naam, type: 'openaicompatible', baseUrl: baseUrl, key: sleutel };
    var model = { id: modelId, name: modelNaam || naam, provider: naam };
    var lijst = lees('providers');
    if (!Array.isArray(lijst)) lijst = [];
    var gevonden = false;
    for (var i = 0; i < lijst.length; i++) {
      if (lijst[i] && String(lijst[i].name).toLowerCase() === naam.toLowerCase()) {
        lijst[i] = aanbieder;
        gevonden = true;
      }
    }
    if (!gevonden) lijst.unshift(aanbieder);
    schrijf('providers', lijst);
    var huidig = lees('current-provider');
    if (!huidig || String(huidig.name).toLowerCase() === naam.toLowerCase()) {
      schrijf('current-provider', aanbieder);
      var m = lees('current-model');
      if (!m || m.provider === naam || !huidig) schrijf('current-model', model);
    }
  }

  /** AIDG weer uit de lijst (na ontkoppelen in het portaal). */
  function haalAanbiederWeg(naam) {
    var lijst = lees('providers');
    if (Array.isArray(lijst)) {
      schrijf('providers', lijst.filter(function (p) {
        return !(p && String(p.name).toLowerCase() === naam.toLowerCase());
      }));
    }
    var huidig = lees('current-provider');
    if (huidig && String(huidig.name).toLowerCase() === naam.toLowerCase()) {
      schrijf('current-provider', null);
      schrijf('current-model', null);
    }
  }

  /** De verbinding van de AIDG-app, of null. */
  function leesVerbinding() {
    if (!window.AscDesktopEditor || !window.AscDesktopEditor.GetInstallPlugins) return null;
    var lijsten = JSON.parse(window.AscDesktopEditor.GetInstallPlugins());
    var eigen = lijsten && lijsten[1] && lijsten[1].url;
    if (!eigen) return null;
    // file:///C:/Users/... → onlyoffice://plugin/C:/Users/... (zelfde herkomst als deze pagina)
    var pad = String(eigen).replace(/^file:\/\/\//, '').split(' ').join('%20');
    if (pad.charAt(pad.length - 1) !== '/') pad += '/';
    var xhr = new XMLHttpRequest();
    // Geen ?t=… erachter: dan geeft de office een leeg bestand terug (gemeten, 9.4.0).
    xhr.open('GET', 'onlyoffice://plugin/' + pad + 'aidg-koppeling/verbinding.json', false);
    xhr.send(null);
    if (xhr.status !== 200 && xhr.status !== 0) return null;
    var v = JSON.parse(xhr.responseText || 'null');
    return v && v.versie === 1 && v.baseUrl && v.sleutel ? v : null;
  }

  /* ---------- sha256 (crypto.subtle is er niet op het plugin-schema) ---------- */
  function sha256hex(tekst) {
    var K = [], H = [], p = 0, n = 2;
    while (p < 64) {
      var priem = true;
      for (var d = 2; d * d <= n; d++) if (n % d === 0) { priem = false; break; }
      if (priem) {
        if (p < 8) H[p] = (Math.pow(n, 1 / 2) * 4294967296) | 0;
        K[p++] = (Math.pow(n, 1 / 3) * 4294967296) | 0;
      }
      n++;
    }
    var b = unescape(encodeURIComponent(tekst)), w = [], l = b.length * 8, i, j;
    for (i = 0; i < b.length; i++) w[i >> 2] |= b.charCodeAt(i) << (24 - (i % 4) * 8);
    w[l >> 5] |= 0x80 << (24 - (l % 32));
    w[(((l + 64) >> 9) << 4) + 15] = l;
    var r = function (x, c) { return (x >>> c) | (x << (32 - c)); };
    for (j = 0; j < w.length; j += 16) {
      var a = H.slice(0), W = [];
      for (i = 0; i < 64; i++) {
        if (i < 16) W[i] = w[j + i] | 0;
        else {
          var s0 = r(W[i - 15], 7) ^ r(W[i - 15], 18) ^ (W[i - 15] >>> 3);
          var s1 = r(W[i - 2], 17) ^ r(W[i - 2], 19) ^ (W[i - 2] >>> 10);
          W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
        }
        var t1 = (a[7] + (r(a[4], 6) ^ r(a[4], 11) ^ r(a[4], 25)) + ((a[4] & a[5]) ^ (~a[4] & a[6])) + K[i] + W[i]) | 0;
        var t2 = ((r(a[0], 2) ^ r(a[0], 13) ^ r(a[0], 22)) + ((a[0] & a[1]) ^ (a[0] & a[2]) ^ (a[1] & a[2]))) | 0;
        a = [(t1 + t2) | 0].concat(a.slice(0, 7));
        a[4] = (a[4] + t1) | 0;
      }
      for (i = 0; i < 8; i++) H[i] = (H[i] + a[i]) | 0;
    }
    var uit = '';
    for (i = 0; i < 8; i++) uit += ('00000000' + (H[i] >>> 0).toString(16)).slice(-8);
    return uit;
  }

  function nieuweSleutel() {
    var bytes = new Uint8Array(32);
    window.crypto.getRandomValues(bytes);
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return 'aidgo_' + btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function post(url, data) {
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    }).then(function (r) {
      return r.json();
    });
  }

  /* ---------- het inlogvenster ---------- */
  var STIJL =
    '#aidg-in{position:fixed;inset:0;z-index:2147483000;display:flex;align-items:center;justify-content:center;' +
    'background:rgba(0,0,0,.35);font:13px/1.45 "Segoe UI",Arial,sans-serif}' +
    '#aidg-in .k{background:#fff;color:#1f2328;border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,.25);' +
    'padding:20px 22px;max-width:340px;width:calc(100% - 40px)}' +
    '#aidg-in h2{font-size:16px;margin:0 0 8px}#aidg-in p{margin:0 0 12px}' +
    '#aidg-in .c{font-size:26px;font-weight:700;letter-spacing:3px;text-align:center;margin:6px 0 12px}' +
    '#aidg-in button{font:inherit;border-radius:6px;padding:7px 12px;cursor:pointer;border:1px solid #c9ced6;background:#fff;color:#1f2328}' +
    '#aidg-in button.h{background:#3b5bdb;border-color:#3b5bdb;color:#fff}' +
    '#aidg-in .r{display:flex;gap:8px;flex-wrap:wrap}#aidg-in .f{color:#c92a2a}' +
    '#aidg-in a{color:#3b5bdb}#aidg-in .m{color:#5c6370;font-size:12px}' +
    '#aidg-pil{position:fixed;right:10px;bottom:10px;z-index:2147483000;font:12px "Segoe UI",Arial,sans-serif;' +
    'border-radius:14px;padding:5px 11px;border:1px solid #3b5bdb;background:#3b5bdb;color:#fff;cursor:pointer}';

  function zorgStijl() {
    if (document.getElementById('aidg-stijl')) return;
    var s = document.createElement('style');
    s.id = 'aidg-stijl';
    s.textContent = STIJL;
    document.head.appendChild(s);
  }

  function sluit() {
    var o = document.getElementById('aidg-in');
    if (o) o.parentNode.removeChild(o);
  }

  function kaart(html) {
    zorgStijl();
    sluit();
    var o = document.createElement('div');
    o.id = 'aidg-in';
    o.innerHTML = '<div class="k">' + html + '</div>';
    document.body.appendChild(o);
    return o;
  }

  function opent(o, url) {
    var a = o.querySelector('a');
    if (a) a.onclick = function (e) {
      e.preventDefault();
      window.open(url, '_blank');
    };
  }

  function toonStart(fout) {
    var o = kaart(
      '<h2>AI van AIDG</h2>' +
        '<p>Log in met je AIDG-account om de AI te gebruiken. Het verbruik gaat van je AIDG-tegoed.</p>' +
        (fout ? '<p class="f">' + fout + '</p>' : '') +
        '<div class="r"><button class="h" data-a="in">Inloggen met AIDG</button>' +
        '<button data-a="nee">Niet nu</button></div>' +
        '<p class="m" style="margin-top:12px">Eigen API-sleutel? Kies dan “Niet nu” en stel je eigen aanbieder in bij de AI-instellingen.</p>'
    );
    o.querySelector('[data-a=in]').onclick = koppel;
    o.querySelector('[data-a=nee]').onclick = function () {
      schrijf(NIET_NU, true);
      sluit();
      toonPil();
    };
  }

  function toonPil() {
    if (pagina !== 'settings.html' || document.getElementById('aidg-pil')) return;
    zorgStijl();
    var b = document.createElement('button');
    b.id = 'aidg-pil';
    b.textContent = 'Inloggen met AIDG';
    b.onclick = function () {
      b.parentNode.removeChild(b);
      toonStart();
    };
    document.body.appendChild(b);
  }

  function klaar(email) {
    var k = kaart(
      '<h2>Ingelogd</h2>' +
        '<p>De AI van AIDG staat klaar' + (email ? ' voor ' + email : '') + '.' +
        ' Sluit dit document en open het opnieuw om de AI te gebruiken.</p>' +
        '<p class="m">Tegoed bijkopen of ontkoppelen kan in het <a href="' + TEGOED + '" target="_blank">AIDG-portaal</a>.</p>' +
        '<div class="r"><button class="h" data-a="ok">Oké</button></div>'
    );
    opent(k, TEGOED);
    k.querySelector('[data-a=ok]').onclick = sluit;
  }

  function koppel() {
    var sleutel = nieuweSleutel();
    var hash = sha256hex(sleutel);
    kaart('<h2>AI van AIDG</h2><p>Even geduld…</p>');
    post(PORTAAL, { actie: 'start', sleutel_hash: hash })
      .then(function (d) {
        if (!d || !d.ok || !d.code) throw new Error('start');
        var o = kaart(
          '<h2>Bevestig in je browser</h2>' +
            '<p>Je browser opent het AIDG-portaal. Log daar in en keur deze code goed:</p>' +
            '<div class="c">' + d.code + '</div>' +
            '<p class="m">Opent er niets? Ga naar <a href="' + d.url + '" target="_blank">aidg.nl/portal/office-koppelen.php</a> en vul de code in.</p>' +
            '<div class="r"><button data-a="stop">Annuleren</button></div>'
        );
        var gestopt = false;
        o.querySelector('[data-a=stop]').onclick = function () {
          gestopt = true;
          toonStart();
        };
        opent(o, d.url);
        window.open(d.url, '_blank');
        var tot = Date.now() + (Number(d.verloopt_in) || 600) * 1000;
        var vraag = function () {
          if (gestopt) return;
          if (Date.now() > tot) return toonStart('De code is verlopen. Probeer het opnieuw.');
          post(PORTAAL, { actie: 'poll', code: d.code, sleutel_hash: hash })
            .then(function (p) {
              if (gestopt) return;
              if (p && p.status === 'goedgekeurd') {
                schrijf(ACCOUNT, { sleutel: sleutel, email: p.email || '' });
                schrijf(NIET_NU, null);
                zetAanbieder(NAAM, GATEWAY, sleutel, 'auto', NAAM);
                return klaar(p.email);
              }
              if (p && (p.status === 'verlopen' || p.status === 'onbekend')) {
                return toonStart('De code is verlopen. Probeer het opnieuw.');
              }
              setTimeout(vraag, 3000);
            })
            .catch(function () {
              setTimeout(vraag, 5000);
            });
        };
        setTimeout(vraag, 3000);
      })
      .catch(function () {
        toonStart('Het AIDG-portaal is niet bereikbaar. Controleer je internet en probeer het opnieuw.');
      });
  }

  /** Is de sleutel nog geldig? Bij 401 (ontkoppeld in het portaal) opnieuw laten inloggen. */
  function controleer(acc) {
    fetch(GATEWAY + '/models', { headers: { Authorization: 'Bearer ' + acc.sleutel } })
      .then(function (r) {
        if (r.status !== 401) return;
        schrijf(ACCOUNT, null);
        haalAanbiederWeg(NAAM);
        toonStart('Deze computer is ontkoppeld van je AIDG-account. Log opnieuw in.');
      })
      .catch(function () {
        /* offline: niets doen */
      });
  }

  function opScherm(f) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', f);
    else f();
  }

  // Test-haakje (alleen gebruikt door test/koppeling.test.js, buiten de office).
  if (window.__aidgKoppelingTest) window.__aidgKoppelingTest({ sha256hex: sha256hex, nieuweSleutel: nieuweSleutel });

  try {
    var v = leesVerbinding();
    if (v) {
      zetAanbieder(v.naam || NAAM, v.baseUrl, v.sleutel, v.model || 'aidg', v.modelNaam);
      return;
    }
  } catch (e) {
    /* geen app: verder zonder */
  }
  try {
    var acc = lees(ACCOUNT);
    if (acc && acc.sleutel) {
      zetAanbieder(NAAM, GATEWAY, acc.sleutel, 'auto', NAAM);
      if (venster) opScherm(function () { controleer(acc); });
      return;
    }
    if (!venster) return;
    var huidig = lees('current-provider');
    var eigen = huidig && String(huidig.name).toLowerCase() !== NAAM.toLowerCase();
    opScherm(function () {
      if (eigen || lees(NIET_NU)) toonPil();
      else toonStart();
    });
  } catch (e) {
    /* nooit de AI zelf breken */
  }
})();
