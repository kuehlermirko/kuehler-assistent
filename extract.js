/* Kühler-Assistent – Fristen und Teilenummern aus Mails herauslesen (ohne KI, läuft auch unter Node) */
(function (root) {
  "use strict";

  var MONTHS = {
    januar: 0, jan: 0, january: 0, februar: 1, feb: 1, february: 1, märz: 2, maerz: 2, mrz: 2, march: 2, mar: 2,
    april: 3, apr: 3, mai: 4, may: 4, juni: 5, jun: 5, june: 5, juli: 6, jul: 6, july: 6, august: 7, aug: 7,
    september: 8, sep: 8, sept: 8, oktober: 9, okt: 9, october: 9, oct: 9, november: 10, nov: 10, dezember: 11, dez: 11, december: 11, dec: 11
  };
  var MONTH_RX = "(januar|februar|märz|maerz|april|mai|juni|juli|august|september|oktober|november|dezember|january|february|march|may|june|july|october|december|jan|feb|mrz|mar|apr|jun|jul|aug|sept?|okt|oct|nov|dez|dec)";
  var DEADLINE_WORDS = /(bis|spätestens|spaetestens|frist|termin|deadline|by|until|no later than|due|benötig|benoetig|need|liefer|delivery|antwort|response|reply)/i;

  /* Nur den neuen Teil der Mail auswerten (ohne zitierten Verlauf) */
  function newPart(body) {
    var b = String(body || "");
    var cut = b.search(/^\s*(?:From|Von|Gesendet|Sent)\s*:.*$|^-{2,}\s*(?:Original|Ursprüngliche)/im);
    return cut > 0 ? b.slice(0, cut) : b;
  }

  function mk(y, m, d) {
    var x = new Date(y, m, d, 9, 0, 0);
    return isNaN(x) || x.getMonth() !== m ? null : x;
  }
  function isoWeekMonday(year, week) {
    var jan4 = new Date(year, 0, 4);
    var day = (jan4.getDay() + 6) % 7;
    var mon = new Date(year, 0, 4 - day + (week - 1) * 7, 9, 0, 0);
    return mon;
  }
  function nextYearIfPast(date, ref) {
    // „27.11.“ ohne Jahr: liegt das Datum mehr als 2 Monate zurück, ist das Folgejahr gemeint
    if (date && date < new Date(ref.getFullYear(), ref.getMonth() - 2, ref.getDate())) date.setFullYear(date.getFullYear() + 1);
    return date;
  }

  function snippetAround(text, idx, len) {
    var s = Math.max(0, idx - 60), e = Math.min(text.length, idx + len + 60);
    return (s > 0 ? "…" : "") + text.slice(s, e).replace(/\s+/g, " ").trim() + (e < text.length ? "…" : "");
  }

  function deadlines(mail, refDate) {
    var ref = refDate ? new Date(refDate) : new Date();
    var text = (mail.subject || "") + "\n" + newPart(mail.body);
    var out = [], seen = {};
    function add(date, idx, len, label) {
      if (!date || isNaN(date)) return;
      var key = date.toDateString();
      if (seen[key]) return;
      var ctx = text.slice(Math.max(0, idx - 40), idx);
      seen[key] = 1;
      out.push({ date: date, label: label, snippet: snippetAround(text, idx, len), strong: DEADLINE_WORDS.test(ctx) });
    }
    var m, rx;
    // 27.11.2026 / 27.11.26 / 27.11.
    rx = /\b(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})?(?!\d)/g;
    while ((m = rx.exec(text)) !== null) {
      var d = +m[1], mo = +m[2] - 1, y = m[3] ? (m[3].length === 2 ? 2000 + +m[3] : +m[3]) : ref.getFullYear();
      if (d < 1 || d > 31 || mo < 0 || mo > 11) continue;
      var dt = mk(y, mo, d);
      if (!m[3]) dt = nextYearIfPast(dt, ref);
      add(dt, m.index, m[0].length, m[0]);
    }
    // 2026-11-27
    rx = /\b(20\d{2})-(\d{2})-(\d{2})\b/g;
    while ((m = rx.exec(text)) !== null) add(mk(+m[1], +m[2] - 1, +m[3]), m.index, m[0].length, m[0]);
    // 27. November (2026) / November 27(, 2026)
    rx = new RegExp("\\b(\\d{1,2})\\.?\\s+" + MONTH_RX + "\\.?(?:\\s+(20\\d{2}))?", "gi");
    while ((m = rx.exec(text)) !== null) {
      var y2 = m[3] ? +m[3] : ref.getFullYear(), dt2 = mk(y2, MONTHS[m[2].toLowerCase()], +m[1]);
      if (!m[3]) dt2 = nextYearIfPast(dt2, ref);
      add(dt2, m.index, m[0].length, m[0]);
    }
    rx = new RegExp("\\b" + MONTH_RX + "\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:,?\\s+(20\\d{2}))?\\b", "gi");
    while ((m = rx.exec(text)) !== null) {
      var y3 = m[3] ? +m[3] : ref.getFullYear(), dt3 = mk(y3, MONTHS[m[1].toLowerCase()], +m[2]);
      if (!m[3]) dt3 = nextYearIfPast(dt3, ref);
      add(dt3, m.index, m[0].length, m[0]);
    }
    // Anfang / Mitte / Ende Dezember, beginning/mid/end of December
    rx = new RegExp("\\b(anfang|mitte|ende|beginning of|early|mid|middle of|end of|late)\\s+" + MONTH_RX + "(?:\\s+(20\\d{2}))?", "gi");
    while ((m = rx.exec(text)) !== null) {
      var part = m[1].toLowerCase(), day = /anfang|beginning|early/.test(part) ? 1 : /mitte|mid|middle/.test(part) ? 15 : 25;
      var y4 = m[3] ? +m[3] : ref.getFullYear(), dt4 = mk(y4, MONTHS[m[2].toLowerCase()], day);
      if (!m[3]) dt4 = nextYearIfPast(dt4, ref);
      add(dt4, m.index, m[0].length, m[0]);
    }
    // KW 49 / KW49/2026 / CW 49 / week 49
    rx = /\b(?:KW|CW|week|Woche)\s*(\d{1,2})(?:\s*[\/.]\s*(20\d{2}))?\b/gi;
    while ((m = rx.exec(text)) !== null) {
      var wk = +m[1]; if (wk < 1 || wk > 53) continue;
      var y5 = m[2] ? +m[2] : ref.getFullYear();
      var mon = isoWeekMonday(y5, wk);
      if (!m[2] && mon < new Date(ref.getFullYear(), ref.getMonth() - 2, 1)) mon = isoWeekMonday(y5 + 1, wk);
      add(mon, m.index, m[0].length, m[0] + " (Montag)");
    }
    // heute / morgen / today / tomorrow / ASAP
    rx = /\b(heute|today|morgen|tomorrow|asap|umgehend|sofort|urgent|dringend)\b/gi;
    while ((m = rx.exec(text)) !== null) {
      if (/(guten|good)\s*$/i.test(text.slice(Math.max(0, m.index - 8), m.index))) continue; // „Guten Morgen“
      // Haftungsausschlüsse in Signaturen („… benachrichtigen Sie uns sofort …“) sind keine Frist
      if (/(irrt(ü|ue)mlich|in\s+error|benachrichtigen|notify|vertraulich|confidential)/i.test(text.slice(Math.max(0, m.index - 160), m.index + 80))) continue;
      var w = m[1].toLowerCase(), base = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate(), 15, 0, 0);
      if (w === "morgen" || w === "tomorrow") base.setDate(base.getDate() + 1);
      add(base, m.index, m[0].length, m[0]);
    }
    // Nur sinnvolle Zeitspanne: ab 2 Monaten zurück bis 18 Monate voraus
    var min = new Date(ref.getFullYear(), ref.getMonth() - 2, 1), max = new Date(ref.getFullYear(), ref.getMonth() + 18, 1);
    out = out.filter(function (x) { return x.date >= min && x.date <= max; });
    out.sort(function (a, b) { return (b.strong - a.strong) || (a.date - b.date); });
    return out.slice(0, 6);
  }

  /* ---------- Teile-Index ---------- */
  function splitLine(line, delim) {
    delim = delim || ";";
    if (line.indexOf('"') < 0) return line.split(delim);
    var out = [], cur = "", inq = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (c === '"') { if (inq && line[i + 1] === '"') { cur += '"'; i++; } else inq = !inq; }
      else if (c === delim && !inq) { out.push(cur); cur = ""; }
      else cur += c;
    }
    out.push(cur);
    return out;
  }
  function parse(text) {
    var lines = String(text).replace(/^﻿/, "").split(/\r?\n/);
    var head = splitLine(lines[0] || "").map(function (h) { return h.trim(); });
    function col(name) { for (var i = 0; i < head.length; i++) if (head[i].toLowerCase().indexOf(name) === 0) return i; return -1; }
    var iNr = col("nummer"), iAn = col("an-nummern"), iKid = col("kid"), iBez = col("bezeichnung"), iRma = col("reklamation"), iFirst = col("erstes"), iLast = col("letztes");
    var map = new Map();
    for (var k = 1; k < lines.length; k++) {
      if (!lines[k]) continue;
      var f = splitLine(lines[k]);
      var nr = (f[iNr] || "").trim();
      if (!nr) continue;
      var rec = {
        nummer: nr, kid: (f[iKid] || "").trim(), bezeichnung: (f[iBez] || "").trim(), an: (f[iAn] || "").trim(),
        rma: iRma >= 0 ? (f[iRma] || "").trim() : "", von: (f[iFirst] || "").trim(), bis: (f[iLast] || "").trim()
      };
      map.set(nr.toUpperCase(), rec);
      rec.kid.split(/[,\s]+/).forEach(function (k) { if (k && !map.has("KID:" + k)) map.set("KID:" + k, rec); });
    }
    return map;
  }
  /* Kandidaten: Zeichnungs-/Artikelnummern und KIDs mit mind. 6 Ziffern, auch mit Index-Endung */
  function textTokens(text) {
    var rx = /\b[A-Z]{0,4}\d[A-Z0-9]{5,15}(?:[_\-.][A-Z0-9]{1,5})?\b/gi, m, set = {};
    while ((m = rx.exec(text)) !== null) {
      var t = m[0].toUpperCase();
      if ((t.match(/\d/g) || []).length < 6) continue;
      if (/^(AN|AU|BE|RE|EB|AB|LS)\d{7}$/.test(t) || /^20\d{6}$/.test(t)) continue;
      set[t] = 1;
      var stem = t.split(/[_\-.]/)[0];
      if (stem !== t && stem.length >= 6) set[stem] = 1;
    }
    return Object.keys(set);
  }
  function partTokens(mail) {
    return textTokens((mail.subject || "") + "\n" + (mail.body || "") + "\n" + (mail.attachments || []).join("\n"));
  }
  /* Kompletter SelectLine-Artikelstamm (export.csv): Schlüssel KID, Kundenartikelnummer, Zeichnungsnummer */
  function parseStamm(text) {
    var lines = String(text).replace(/^\uFEFF/, "").split(/\r?\n/);
    var delim = (lines[0] || "").split(";").length > (lines[0] || "").split(",").length ? ";" : ",";
    var head = splitLine(lines[0] || "", delim).map(function (h) { return h.trim().toLowerCase(); });
    function col(names) { for (var i = 0; i < head.length; i++) for (var j = 0; j < names.length; j++) if (head[i].indexOf(names[j]) === 0) return i; return -1; }
    var iKid = col(["kühler id", "kuehler id", "kid"]), iBez = col(["bezeichnung"]), iZus = col(["zusatz"]), iKd = col(["kundenartikel"]),
      iZng = col(["zeichungsnummer", "zeichnungsnummer"]), iIdx = col(["index"]), iSl = col(["stueckliste", "stückliste"]);
    var map = new Map();
    function key(k, rec) { k = String(k || "").trim().toUpperCase(); if (k.length >= 5 && !map.has(k)) map.set(k, rec); var c = k.replace(/[\s\-.]/g, ""); if (c !== k && c.length >= 6 && !map.has(c)) map.set(c, rec); }
    for (var n = 1; n < lines.length; n++) {
      if (!lines[n]) continue;
      var f = splitLine(lines[n], delim);
      var kid = (f[iKid] || "").trim();
      if (!kid) continue;
      var rec = { kid: kid, bezeichnung: (f[iBez] || "").trim(), zusatz: iZus >= 0 ? (f[iZus] || "").trim() : "", kdnr: iKd >= 0 ? (f[iKd] || "").trim() : "",
        zng: iZng >= 0 ? (f[iZng] || "").trim() : "", index: iIdx >= 0 ? (f[iIdx] || "").trim() : "", sl: iSl >= 0 ? (f[iSl] || "").trim() : "" };
      map.set("KID:" + kid, rec);
      if (/^\d{6,}$/.test(kid) && !map.has(kid)) map.set(kid, rec);
      key(rec.kdnr, rec); key(rec.zng, rec);
    }
    return map;
  }
  /* Treffer aus beiden Quellen zusammenführen: Stamm (KID, Bezeichnung) und Teile-Index (frühere AN, Reklamationen) */
  function matchAll(sources, stamm, index) {
    var hits = [], seen = {}, checked = 0;
    sources.forEach(function (src) {
      src.tokens.forEach(function (t) {
        checked++;
        var rec = stamm ? stamm.get(t) : null;
        var hist = index ? (index.get(t) || (rec && (index.get(String(rec.kdnr).toUpperCase()) || index.get(String(rec.zng).toUpperCase()) || index.get("KID:" + rec.kid)))) : null;
        if (!rec && hist && hist.kid && stamm) rec = stamm.get("KID:" + hist.kid.split(/[,\s]+/)[0]) || null;
        if (!rec && !hist) return;
        var id = rec ? "K" + rec.kid : "N" + hist.nummer;
        if (seen[id]) { if (seen[id].quellen.indexOf(src.name) < 0) seen[id].quellen.push(src.name); return; }
        var h = { token: t, rec: rec, hist: hist, quellen: [src.name] };
        seen[id] = h; hits.push(h);
      });
    });
    return { hits: hits, checked: checked };
  }
  function matchParts(mail, map) {
    var toks = partTokens(mail), hits = [], seen = {};
    toks.forEach(function (t) {
      var r = map.get(t);
      if (r && !seen[r.nummer]) { seen[r.nummer] = 1; hits.push(r); }
    });
    return { hits: hits, checked: toks.length };
  }

  var api = { deadlines: deadlines, newPart: newPart, parse: parse, parseStamm: parseStamm, textTokens: textTokens, partTokens: partTokens, matchParts: matchParts, matchAll: matchAll };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else { root.KuehlerFristen = api; root.KuehlerTeile = api; }
})(this);
