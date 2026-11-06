/* Kühler-Assistent – Microsoft Graph (nur lesend): Anmeldung, Mailverlauf, Kühler-Gehirn */
(function (root) {
  "use strict";
  var cfg = root.KUEHLER_CONFIG;
  var pca = null;
  var GRAPH = "https://graph.microsoft.com/v1.0";

  function configured() {
    return cfg.clientId && cfg.clientId.indexOf("__") !== 0;
  }

  async function client() {
    if (pca) return pca;
    // Nested App Authentication: Anmeldung läuft über das bereits in Outlook angemeldete Konto
    pca = await root.msal.createNestablePublicClientApplication({
      auth: { clientId: cfg.clientId, authority: "https://login.microsoftonline.com/organizations" },
      cache: { cacheLocation: "localStorage" }
    });
    return pca;
  }

  async function token() {
    var app = await client();
    var req = { scopes: cfg.graphScopes };
    try {
      var r = await app.acquireTokenSilent(req);
      return r.accessToken;
    } catch (e) {
      var r2 = await app.acquireTokenPopup(req);
      return r2.accessToken;
    }
  }

  async function get(path, extraHeaders) {
    var t = await token();
    var headers = { Authorization: "Bearer " + t };
    if (extraHeaders) Object.keys(extraHeaders).forEach(function (k) { headers[k] = extraHeaders[k]; });
    var res = await fetch(path.indexOf("http") === 0 ? path : GRAPH + path, { headers: headers });
    if (!res.ok) {
      var msg = res.status + " " + res.statusText;
      try { var j = await res.json(); if (j.error && j.error.message) msg = j.error.message; } catch (e) { /* ignorieren */ }
      throw new Error(msg);
    }
    var ct = res.headers.get("content-type") || "";
    return ct.indexOf("json") >= 0 ? res.json() : res.text();
  }

  function clean(term) {
    var parts = String(term).replace(/["\\]/g, " ").trim().split(/\s+/);
    return parts[parts.length - 1];
  }

  /* Frühere Mails: Suche über alle Ordner nach Vorgangsnummern, sonst nach Absender */
  async function history(mail, result) {
    var terms = result.refs.all.slice(0, 3);
    var query = terms.length ? terms.map(clean).join(" OR ") : (mail.fromAddress ? "from:" + clean(mail.fromAddress) : "");
    if (!query) return [];
    var url = "/me/messages?$search=" + encodeURIComponent('"' + query + '"') +
      "&$top=25&$select=subject,from,receivedDateTime,webLink,internetMessageId";
    var data = await get(url);
    var items = (data.value || []).filter(function (m) { return m.internetMessageId !== mail.messageId; });
    items.sort(function (a, b) { return new Date(b.receivedDateTime) - new Date(a.receivedDateTime); });
    return items.slice(0, 8).map(function (m) {
      return {
        date: m.receivedDateTime,
        subject: m.subject || "(ohne Betreff)",
        from: m.from && m.from.emailAddress ? (m.from.emailAddress.name || m.from.emailAddress.address) : "",
        link: m.webLink
      };
    });
  }

  function stripFrontmatter(md) {
    return md.replace(/^---[\s\S]*?\n---\s*\n/, "");
  }

  function snippet(md, words) {
    var text = stripFrontmatter(md);
    var lines = text.split(/\r?\n/).filter(function (l) { return l.trim() && !/^#{1,6}\s/.test(l); });
    var lower = words.map(function (w) { return w.toLowerCase(); });
    var hits = lines.filter(function (l) {
      var ll = l.toLowerCase();
      return lower.some(function (w) { return w && ll.indexOf(w) >= 0; });
    });
    var pick = (hits.length ? hits : lines).slice(0, 3).join(" ");
    pick = pick.replace(/\[\[([^\]|]+)\|?([^\]]*)\]\]/g, function (_, a, b) { return b || a; }).replace(/[*_`>]/g, "").replace(/\s+/g, " ").trim();
    return pick.length > 260 ? pick.slice(0, 257) + "…" : pick;
  }

  /* Kühler-Gehirn: Suche im eigenen OneDrive unterhalb des hinterlegten Vault-Ordners */
  async function brain(mail, result, settings) {
    if (!settings.vaultPath) return null;
    var base = settings.vaultPath.replace(/^\/+|\/+$/g, "");
    var terms = [];
    if (result.customer) terms.push(result.customer);
    if (mail.fromName && !result.customer) terms.push(mail.fromName);
    terms = terms.concat(result.refs.an, result.refs.au, result.refs.be, result.refs.rfq).slice(0, 4);
    if (!terms.length && mail.fromName) terms.push(mail.fromName);
    var found = {}, order = [];
    for (var i = 0; i < terms.length; i++) {
      var url = "/me/drive/root:/" + encodeURI(base) + ":/search(q='" + encodeURIComponent(terms[i].replace(/'/g, "''")) + "')" +
        "?$top=10&$select=id,name,webUrl,parentReference,file";
      var data;
      try { data = await get(url); } catch (e) { if (i === 0) throw e; continue; }
      (data.value || []).forEach(function (it) {
        if (!it.file || !/\.md$/i.test(it.name) || found[it.id]) return;
        found[it.id] = it; order.push(it);
      });
      if (order.length >= 6) break;
    }
    var top = order.slice(0, 4);
    var out = [];
    for (var j = 0; j < top.length; j++) {
      var it = top[j], text = "";
      try { text = await get("/me/drive/items/" + it.id + "/content"); } catch (e) { text = ""; }
      var parent = (it.parentReference && it.parentReference.path) || "";
      var rel = parent.replace(/^\/drive\/root:\/?/, "");
      rel = decodeURIComponent(rel);
      var inVault = rel.toLowerCase().indexOf(base.toLowerCase()) === 0 ? rel.slice(base.length).replace(/^\/+/, "") : "";
      var file = (inVault ? inVault + "/" : "") + it.name.replace(/\.md$/i, "");
      out.push({
        title: it.name.replace(/\.md$/i, ""),
        text: snippet(String(text), terms),
        obsidian: "obsidian://open?vault=" + encodeURIComponent(settings.vaultName || "Kühler") + "&file=" + encodeURIComponent(file),
        web: it.webUrl
      });
    }
    return out;
  }

  root.KuehlerGraph = { configured: configured, history: history, brain: brain };
})(this);
