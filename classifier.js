/* Kühler-Assistent – regelbasierte Erkennung von Mailtyp, Kunde und Vorgangsnummern
   Läuft im Add-in und zum Testen auch unter Node. */
(function (root) {
  "use strict";

  var REF_PATTERNS = {
    an: /\bAN\s?(\d{7})\b/gi,
    au: /\bAU\s?(\d{7})\b/gi,
    be: /\bBE\s?(\d{7})\b/gi,
    re: /\bRE\s?(\d{7})\b/gi,
    po: /\bPO\s?#?\s?(\d{8,10})\b/gi,
    eb: /\bEB\s?(\d{5,10})\b/gi,
    sap: /\b(4500\d{6})\b/g,
    gdels: /\b(6000\d{6})\b/g,
    rfq: /\bRFQ\s*(?:No\.?|Nr\.?|#)?\s*(\d{6,10})\b/gi,
    projekt: /\bL-(\d{4})\b/g
  };
  var REF_PREFIX = { an: "AN", au: "AU", be: "BE", re: "RE", po: "PO", eb: "EB", sap: "", gdels: "", rfq: "RFQ ", projekt: "L-" };

  var RX = {
    order: /(purchase\s+order|\bbestellung\b|\border\s+no\b|\bPO\s?#?\s?\d{6,}|\bEB\s?\d{5,}|\b4500\d{6}\b|blanket\s+order|abruf)/i,
    rfq: /(\bRFQ\b|request\s+for\s+(a\s+)?quot|quotation\s+request|\banfrage\b|angebotsanfrage|preisanfrage|\binquiry\b|\benquiry\b|request\s+for\s+offer)/i,
    ab: /(auftragsbest(ä|ae)tigung|order\s+confirmation|best(ä|ae)tigung\s+ihrer\s+bestellung|confirmation\s+of\s+(your\s+)?order|\bAB[\s-]?Nr|\bAB\s?\d{3,})/i,
    angebot: /(\bangebot\b|\bquotation\b|\bquote\b|\boffer\b|\bofferte\b|\bpreisangebot\b)/i,
    rechnung: /(\brechnung\b|\binvoice\b|\bgutschrift\b|\bcredit\s+note\b)/i
  };

  function uniq(arr) {
    var seen = {}, out = [];
    for (var i = 0; i < arr.length; i++) if (!seen[arr[i]]) { seen[arr[i]] = 1; out.push(arr[i]); }
    return out;
  }

  function findRefs(text) {
    var refs = {}, all = [];
    Object.keys(REF_PATTERNS).forEach(function (k) {
      var rx = new RegExp(REF_PATTERNS[k].source, REF_PATTERNS[k].flags), m, list = [];
      while ((m = rx.exec(text)) !== null) list.push(REF_PREFIX[k] + m[1]);
      refs[k] = uniq(list);
      all = all.concat(refs[k]);
    });
    refs.all = uniq(all);
    return refs;
  }

  function domainOf(address) {
    var at = (address || "").lastIndexOf("@");
    return at >= 0 ? address.slice(at + 1).toLowerCase() : "";
  }

  function domainMatches(domain, list) {
    for (var i = 0; i < list.length; i++) {
      var d = list[i].toLowerCase();
      if (domain === d || domain.slice(-(d.length + 1)) === "." + d) return true;
    }
    return false;
  }

  function findCustomer(mail, cfg) {
    var domain = domainOf(mail.fromAddress);
    var byDomain = null;
    for (var i = 0; i < cfg.customers.length; i++) {
      var c = cfg.customers[i];
      if (c.domains.length && domainMatches(domain, c.domains)) { byDomain = c; break; }
    }
    // Ohne Domain-Treffer zählt nur der Absendername – sonst würden Lieferanten,
    // die im Text einen Kunden erwähnen, fälschlich als Kunde erkannt.
    // Mit Domain-Treffer verfeinern Betreff/Text (z. B. Stadler US innerhalb von stadlerrail.com).
    var hay = byDomain
      ? (mail.fromName || "") + " " + (mail.subject || "") + " " + (mail.body || "").slice(0, 1500)
      : (mail.fromName || "");
    hay = hay.toLowerCase();
    for (var j = 0; j < cfg.customers.length; j++) {
      var n = cfg.customers[j];
      if (byDomain && n.name.indexOf(byDomain.name.split(" ")[0]) !== 0) continue;
      for (var k = 0; k < n.names.length; k++) {
        if (hay.indexOf(n.names[k].toLowerCase()) >= 0) return n.name;
      }
    }
    return byDomain ? byDomain.name : null;
  }

  function classify(mail, cfg) {
    var subject = mail.subject || "";
    var body = (mail.body || "").slice(0, 8000);
    var text = subject + "\n" + body;
    var refs = findRefs(text);
    var domain = domainOf(mail.fromAddress);
    var intern = domainMatches(domain, cfg.ownDomains);
    var customer = intern ? null : findCustomer(mail, cfg);
    var type;

    if (intern) {
      type = "intern";
    } else if (customer) {
      if (RX.order.test(subject) && !RX.rfq.test(subject)) type = "kundenbestellung";
      else if (RX.rfq.test(subject)) type = "rfq";
      else if (RX.order.test(body)) type = "kundenbestellung";
      else if (RX.rfq.test(body)) type = "rfq";
      else type = "sonstige";
    } else {
      if (RX.rechnung.test(subject)) type = "rechnung";
      else if (RX.ab.test(subject)) type = "lieferanten_ab";
      else if (RX.angebot.test(subject)) type = "lieferantenangebot";
      else if (RX.ab.test(body)) type = "lieferanten_ab";
      else if (RX.angebot.test(body) && mail.attachments && mail.attachments.length) type = "lieferantenangebot";
      else type = "sonstige";
    }

    var t = cfg.types[type];
    var vorgang = [];
    if (customer) vorgang.push(customer);
    vorgang = vorgang.concat(refs.all.slice(0, 4));

    return {
      type: type,
      label: t.label,
      tone: t.tone,
      task: t.task,
      skill: t.skill,
      skills: t.skills,
      customer: customer,
      partner: customer || (intern ? "Intern" : (mail.fromName || domain)),
      refs: refs,
      vorgang: vorgang.length ? vorgang.join(" · ") : "Kein Vorgang erkannt",
      language: /[äöüß]|\b(sehr geehrte|hallo|anbei|ihrer|bitte|vielen dank|mit freundlichen)/i.test(text) ? "Deutsch" : "Englisch"
    };
  }

  function fmtDate(d) {
    if (!d) return "";
    var x = new Date(d);
    if (isNaN(x)) return String(d);
    function p(n) { return (n < 10 ? "0" : "") + n; }
    return p(x.getDate()) + "." + p(x.getMonth() + 1) + "." + x.getFullYear() + ", " + p(x.getHours()) + ":" + p(x.getMinutes());
  }

  function mailHeader(mail, result) {
    var lines = [
      "Kühler-Assistent – Übergabe aus Outlook",
      "Mail: „" + (mail.subject || "(ohne Betreff)") + "“",
      "Von: " + (mail.fromName ? mail.fromName + " <" + mail.fromAddress + ">" : mail.fromAddress),
      "Empfangen: " + fmtDate(mail.received)
    ];
    if (mail.messageId) lines.push("Message-ID: " + mail.messageId);
    lines.push("Erkannt: " + result.label + " · " + result.vorgang);
    if (mail.attachments && mail.attachments.length) lines.push("Anhänge: " + mail.attachments.join(", "));
    return lines;
  }

  function buildHandoffPrompt(mail, result) {
    var lines = mailHeader(mail, result);
    lines.push("");
    lines.push("Auftrag: " + result.task);
    lines.push("Bitte die Mail samt Anhängen selbst über den Microsoft-365-Konnektor holen (Skill mail-holen) und vorher das Kühler-Gehirn prüfen.");
    if (result.skill) lines.push("Skill: /" + result.skill + (result.skills.length > 1 ? " (dazu " + result.skills.filter(function (s) { return s !== result.skill; }).join(", ") + ")" : ""));
    lines.push("Mails nur als Entwurf anlegen, nichts senden.");
    return lines.join("\n");
  }

  function buildReplyPrompt(mail, result, opts) {
    var lines = mailHeader(mail, result);
    lines.push("");
    lines.push("Auftrag: Diese Mail beantworten.");
    lines.push("Ton: " + opts.tone + " · Sprache: " + opts.language);
    if (opts.instruction) lines.push("Inhalt der Antwort: " + opts.instruction);
    if (opts.attachments) lines.push("Anhänge für die Antwort: " + opts.attachments);
    lines.push("Antwort auf die Originalmail in Outlook Web anlegen: Signatur „Kühler“, Arial 10, gegliedert mit Fett und Kursiv, im Ordner „Entwürfe“ speichern, nicht senden.");
    lines.push("Vorher die Mail über den Microsoft-365-Konnektor holen (Skill mail-holen) und das Kühler-Gehirn zu Absender und Vorgang prüfen.");
    return lines.join("\n");
  }

  var api = { classify: classify, findRefs: findRefs, buildHandoffPrompt: buildHandoffPrompt, buildReplyPrompt: buildReplyPrompt, fmtDate: fmtDate, domainOf: domainOf };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.KuehlerClassifier = api;
})(this);
