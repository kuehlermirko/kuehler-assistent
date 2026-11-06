/* Kühler-Assistent – Hausvorlagen, die direkt in Outlook geöffnet werden (ohne KI)
   Gelb markierte Stellen sind Platzhalter, die vor dem Senden ergänzt werden. */
(function (root) {
  "use strict";

  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function ph(text) {
    return '<span style="background:#FFF2A8">' + esc(text) + "</span>";
  }
  function wrap(inner) {
    return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#000000">' + inner + "</div>";
  }
  function p(html) { return '<p style="margin:0 0 10pt 0">' + html + "</p>"; }

  function names(fromName) {
    var n = String(fromName || "").replace(/\(.*?\)/g, "").trim();
    if (!n) return { first: "", last: "" };
    if (n.indexOf(",") > 0) {
      var parts = n.split(",");
      return { last: parts[0].trim(), first: (parts[1] || "").trim().split(/\s+/)[0] || "" };
    }
    var t = n.split(/\s+/);
    return { first: t[0], last: t[t.length - 1] };
  }

  function dateOnly(d) {
    var x = new Date(d);
    if (isNaN(x)) return "";
    function z(v) { return (v < 10 ? "0" : "") + v; }
    return z(x.getDate()) + "." + z(x.getMonth() + 1) + "." + x.getFullYear();
  }

  function salutation(mail, lang, tone) {
    var n = names(mail.fromName);
    if (tone === "Freundschaftlich" && n.first) return (lang === "Englisch" ? "Hi " : "Hallo ") + esc(n.first) + ",";
    if (lang === "Englisch") return "Dear " + ph("Mr/Ms") + " " + esc(n.last || "Sir or Madam") + ",";
    return n.last ? "Sehr geehrte" + ph("r Herr / Frau") + " " + esc(n.last) + "," : "Sehr geehrte Damen und Herren,";
  }

  function firstOr(list, fallback) { return list && list.length ? esc(list[0]) : ph(fallback); }

  /* Jede Vorlage: label, mode ("reply" | "replyAll" | "new"), build(mail, result, ctx) -> {html, subject?, to?} */
  var TEMPLATES = {
    ab_kunde: {
      label: "AB-Mail an Kunden (Hausvorlage)",
      mode: "replyAll",
      build: function (mail, r) {
        var po = r.refs.po.concat(r.refs.po2 || [], r.refs.sap, r.refs.eb, r.refs.gdels || []);
        return {
          html: wrap(
            p(salutation(mail, "Englisch", "Formell")) +
            p("many thanks for your purchase order. Enclosed you will find the order confirmation (" + firstOr(r.refs.au, "AU…") +
              ") for your order (" + firstOr(po, "PO…") + ") from " + esc(dateOnly(mail.received)) + ".") +
            p("If you have any questions or suggestions, please do not hesitate to contact me!")
          ),
          hint: "Betreff vorne um „Kühler GmbH & Co. KG - order confirmation …“ ergänzen, CC C. Müller, F. Jäger, A. vom Berg, V. Barbuto setzen, AB-PDF und Bestell-PDF anhängen."
        };
      }
    },
    rfq_eingang: {
      label: "Eingangsbestätigung der Anfrage",
      mode: "reply",
      build: function (mail, r, ctx) {
        var ref = r.refs.rfq.concat(r.refs.gdels || [], r.refs.an)[0];
        var en = ctx.language === "Englisch";
        return {
          html: wrap(en
            ? p(salutation(mail, "Englisch", ctx.tone)) +
              p("thank you for your request for quotation" + (ref ? " " + esc(ref) : "") + ". We have received the documents and are currently reviewing them.") +
              p("You will receive our offer by " + ph("TT.MM.JJJJ") + ". Should we have any technical questions in the meantime, we will get back to you.")
            : p(salutation(mail, "Deutsch", ctx.tone)) +
              p("vielen Dank für Ihre Anfrage" + (ref ? " " + esc(ref) : "") + ". Die Unterlagen sind bei uns eingegangen und werden derzeit geprüft.") +
              p("Unser Angebot erhalten Sie bis " + ph("TT.MM.JJJJ") + ". Sollten sich vorab technische Rückfragen ergeben, melden wir uns bei Ihnen.")
          )
        };
      }
    },
    ab_info: {
      label: "Info-Mail „AB eingegangen“ an mich",
      mode: "new",
      build: function (mail, r, ctx) {
        var be = r.refs.be[0] || "BE…";
        var th = 'style="background:#1F3A5F;color:#FFFFFF;border:1px solid #1F3A5F;padding:4pt 6pt;text-align:left"';
        var td = 'style="border:1px solid #C9D6E3;padding:4pt 6pt"';
        return {
          to: [ctx.me],
          subject: "AB eingegangen – bitte in SelectLine einpflegen – " + be,
          html: wrap(
            p("Sehr geehrter Herr Spieler,") +
            p("zu folgender Bestellung ist die Auftragsbestätigung eingegangen:") +
            '<table style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;font-size:10pt;margin:0 0 10pt 0">' +
            "<tr><th " + th + ">Bestellung</th><th " + th + ">Lieferant</th><th " + th + ">AB-Nr.</th><th " + th + ">Termin</th><th " + th + ">Prüfung</th></tr>" +
            "<tr><td " + td + ">" + esc(be) + "</td><td " + td + ">" + esc(mail.fromName || mail.fromAddress) + "</td><td " + td + ">" + ph("AB-Nr.") +
            "</td><td " + td + ">" + ph("Termin") + "</td><td " + td + ">" + ph("Menge/Preis/Termin geprüft") + "</td></tr></table>" +
            "<ul style=\"margin:0 0 10pt 0\"><li>AB in SelectLine einpflegen</li><li>AB im Projektordner (07_Auftragsbearbeitung) ablegen</li></ul>"
          ),
          hint: "Ohne Signatur „Kühler“ – falls Outlook sie einsetzt, bitte entfernen."
        };
      }
    },
    angebot_dank: {
      label: "Dank für das Angebot",
      mode: "reply",
      build: function (mail, r, ctx) {
        var en = ctx.language === "Englisch";
        return {
          html: wrap(en
            ? p(salutation(mail, "Englisch", ctx.tone)) +
              p("many thanks for your quotation" + (r.refs.an[0] ? " regarding our request " + esc(r.refs.an[0]) : "") + ". We are reviewing it and will get back to you shortly.")
            : p(salutation(mail, "Deutsch", ctx.tone)) +
              p((ctx.tone === "Freundschaftlich" ? "vielen Dank für dein Angebot" : "vielen Dank für Ihr Angebot") +
                (r.refs.an[0] ? " zu unserer Anfrage " + esc(r.refs.an[0]) : "") +
                (ctx.tone === "Freundschaftlich" ? ". Wir prüfen es und melden uns kurzfristig bei dir." : ". Wir prüfen es und melden uns kurzfristig bei Ihnen."))
          )
        };
      }
    },
    eigener_text: {
      label: "Antwort mit meinem Text",
      mode: "reply",
      build: function (mail, r, ctx) {
        var body = (ctx.instruction || "").split(/\n{2,}/).map(function (para) { return p(esc(para).replace(/\n/g, "<br>")); }).join("");
        return {
          html: wrap(p(salutation(mail, ctx.language, ctx.tone)) + (body || p(ph("Text"))))
        };
      }
    }
  };

  var BY_TYPE = {
    kundenbestellung: ["ab_kunde", "eigener_text"],
    rfq: ["rfq_eingang", "eigener_text"],
    lieferanten_ab: ["ab_info", "eigener_text"],
    lieferantenangebot: ["angebot_dank", "eigener_text"],
    rechnung: ["eigener_text"],
    intern: ["eigener_text"],
    sonstige: ["eigener_text"]
  };

  function forType(type) {
    return (BY_TYPE[type] || ["eigener_text"]).map(function (id) { return { id: id, label: TEMPLATES[id].label, mode: TEMPLATES[id].mode }; });
  }
  function build(id, mail, result, ctx) {
    var t = TEMPLATES[id];
    var out = t.build(mail, result, ctx);
    out.mode = t.mode;
    return out;
  }

  var api = { forType: forType, build: build, names: names };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.KuehlerTemplates = api;
})(this);
