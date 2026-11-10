/* Kühler-Assistent – Microsoft Graph: Anmeldung, Kühler-Gehirn, Projektordner, Ablage, Planner, Rücklauf, Teile-Index */
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

  /* ---------- Hilfen ---------- */
  async function send(method, path, body, headers) {
    var t = await token();
    var h = { Authorization: "Bearer " + t };
    if (headers) Object.keys(headers).forEach(function (k) { h[k] = headers[k]; });
    var res = await fetch(path.indexOf("http") === 0 ? path : GRAPH + path, { method: method, headers: h, body: body });
    if (!res.ok) {
      var msg = res.status + " " + res.statusText;
      try { var j = await res.json(); if (j.error && j.error.message) msg = j.error.message; } catch (e) { /* ignorieren */ }
      var err = new Error(msg); err.status = res.status; throw err;
    }
    var ct = res.headers.get("content-type") || "";
    return ct.indexOf("json") >= 0 ? res.json() : null;
  }
  function encPath(p) { return p.split("/").map(encodeURIComponent).join("/"); }
  function q(s) { return encodeURIComponent(String(s).replace(/'/g, "''")); }

  /* ---------- Projektordner finden (SharePoint „Projekte Thüringen“) ---------- */
  async function projectFolder(result) {
    var d = cfg.drives.projekte;
    var refs = [].concat(result.refs.an, result.refs.po, result.refs.po2 || [], result.refs.eb, result.refs.sap, result.refs.gdels || [])
      .map(function (r) { return r.replace(/^PO/, "").trim(); })
      .filter(function (v, i, a) { return v && a.indexOf(v) === i; }).slice(0, 4);
    for (var i = 0; i < refs.length; i++) {
      var ref = refs[i];
      var data = await get("/drives/" + d + "/root/search(q='" + q(ref) + "')?$top=50&$select=id,name,webUrl,folder,parentReference");
      var cands = (data.value || []).filter(function (it) {
        return it.folder && /^AN\d{7}/.test(it.name) && it.name.indexOf(ref) >= 0;
      });
      // Projekt-Hauptordner erkennen: enthält Unterordner wie 01_…, 03_Kalkulation
      for (var j = 0; j < cands.length && j < 6; j++) {
        var kids = await get("/drives/" + d + "/items/" + cands[j].id + "/children?$top=60&$select=id,name,folder,webUrl");
        var subs = (kids.value || []).filter(function (k) { return k.folder; });
        if (subs.some(function (k) { return /^0[1-9]_/.test(k.name); })) {
          return { id: cands[j].id, name: cands[j].name, webUrl: cands[j].webUrl, driveId: d,
                   subfolders: subs.map(function (k) { return { id: k.id, name: k.name, webUrl: k.webUrl }; }).sort(function (a, b) { return a.name.localeCompare(b.name); }) };
        }
      }
    }
    return null;
  }

  async function ensureSubfolder(folder, name) {
    var hit = folder.subfolders.filter(function (s) { return s.name === name; })[0];
    if (hit) return hit;
    try {
      var created = await send("POST", "/drives/" + folder.driveId + "/items/" + folder.id + "/children",
        JSON.stringify({ name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }), { "Content-Type": "application/json" });
      var s = { id: created.id, name: created.name, webUrl: created.webUrl };
      folder.subfolders.push(s);
      return s;
    } catch (e) {
      if (e.status !== 409) throw e;
      var it = await get("/drives/" + folder.driveId + "/items/" + folder.id + ":/" + encodeURIComponent(name) + "?$select=id,name,webUrl");
      return { id: it.id, name: it.name, webUrl: it.webUrl };
    }
  }

  /* Datei hochladen: bis 4 MB direkt, darüber in Teilen (Upload-Sitzung) */
  async function upload(driveId, parentId, fileName, bytes) {
    var name = fileName.replace(/[\\/:*?"<>|#%]/g, "_");
    var base = "/drives/" + driveId + "/items/" + parentId + ":/" + encodeURIComponent(name) + ":";
    if (bytes.length <= 4 * 1024 * 1024) {
      return send("PUT", base + "/content?@microsoft.graph.conflictBehavior=rename", bytes, { "Content-Type": "application/octet-stream" });
    }
    var session = await send("POST", base + "/createUploadSession",
      JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }), { "Content-Type": "application/json" });
    var chunk = 320 * 1024 * 12, last = null;
    for (var start = 0; start < bytes.length; start += chunk) {
      var end = Math.min(start + chunk, bytes.length);
      var res = await fetch(session.uploadUrl, { method: "PUT", body: bytes.slice(start, end),
        headers: { "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + bytes.length } });
      if (!res.ok && res.status !== 202) throw new Error("Upload abgebrochen (" + res.status + ")");
      if (res.status === 200 || res.status === 201) last = await res.json();
    }
    return last;
  }

  /* Mail selbst als .eml (MIME) laden */
  async function mailMime(restId) {
    var t = await token();
    var res = await fetch(GRAPH + "/me/messages/" + encodeURIComponent(restId) + "/$value", { headers: { Authorization: "Bearer " + t } });
    if (!res.ok) throw new Error("Mail konnte nicht geladen werden (" + res.status + ")");
    return new Uint8Array(await res.arrayBuffer());
  }

  /* ---------- Planner: offene Aufgaben im Bucket der AN ---------- */
  async function plannerTasks(an) {
    var b = await get("/planner/plans/" + cfg.planner.planId + "/buckets?$select=id,name");
    var bucket = (b.value || []).filter(function (x) { return x.name.indexOf(an) === 0; })[0];
    if (!bucket) return { bucket: null, tasks: [] };
    var t = await get("/planner/buckets/" + bucket.id + "/tasks");
    var open = (t.value || []).filter(function (x) { return x.percentComplete < 100; });
    open.sort(function (a, c) { return (a.dueDateTime || "9999") < (c.dueDateTime || "9999") ? -1 : 1; });
    return { bucket: bucket.name, tasks: open.map(function (x) { return { title: x.title, due: x.dueDateTime, started: x.percentComplete > 0 }; }) };
  }

  /* ---------- Rücklaufliste Lieferantenanfragen ---------- */
  async function sheet(path, name) {
    var r = await get("/drives/" + cfg.drives.projekte + "/root:/" + encPath(path) + ":/workbook/worksheets('" + q(name) + "')/usedRange(valuesOnly=true)?$select=values");
    var rows = r.values || [];
    var head = (rows[0] || []).map(function (h) { return String(h).trim(); });
    return rows.slice(1).filter(function (row) { return row.some(function (v) { return v !== "" && v != null; }); }).map(function (row) {
      var o = {}; head.forEach(function (h, i) { o[h] = row[i]; }); return o;
    });
  }
  async function ruecklauf(domain, address, ans) {
    var list = await sheet(cfg.ruecklauf.path, "Lieferanten");
    var dom = (domain || "").toLowerCase(), adr = (address || "").toLowerCase();
    var mine = list.filter(function (r) {
      var key = String(r["Schlüssel (Domain)"] || "").toLowerCase();
      var mails = String(r["E-Mail"] || "").toLowerCase();
      return (key && (key === dom || key === adr)) || (adr && mails.indexOf(adr) >= 0);
    });
    var pakete = [];
    if (ans && ans.length) {
      var p = await sheet(cfg.ruecklauf.path, "Pakete");
      pakete = p.filter(function (r) { return ans.indexOf(String(r["AN-Nr."]).trim()) >= 0; });
    }
    return { lieferant: mine.slice(-8).reverse(), pakete: pakete };
  }

  /* ---------- Teile-Index aus dem SelectLine-Export (Zwischenspeicher im Browser, 12 h) ---------- */
  function idb() {
    return new Promise(function (resolve) {
      try {
        var r = indexedDB.open("kuehler-assistent", 1);
        r.onupgradeneeded = function () { r.result.createObjectStore("cache"); };
        r.onsuccess = function () { resolve(r.result); };
        r.onerror = function () { resolve(null); };
      } catch (e) { resolve(null); }
    });
  }
  async function cacheGet(key) {
    var db = await idb(); if (!db) return null;
    return new Promise(function (resolve) {
      try { var g = db.transaction("cache").objectStore("cache").get(key); g.onsuccess = function () { resolve(g.result || null); }; g.onerror = function () { resolve(null); }; }
      catch (e) { resolve(null); }
    });
  }
  async function cacheSet(key, val) {
    var db = await idb(); if (!db) return;
    try { db.transaction("cache", "readwrite").objectStore("cache").put(val, key); } catch (e) { /* ohne Zwischenspeicher weiter */ }
  }
  var teileMap = null;
  async function teileIndex(force) {
    if (teileMap && !force) return teileMap;
    var c = force ? null : await cacheGet("teile");
    var text;
    if (c && Date.now() - c.at < 12 * 3600 * 1000) text = c.text;
    else {
      var t = await token();
      var res = await fetch(GRAPH + "/drives/" + cfg.drives.dataExchange + "/root:/" + encPath(cfg.teileIndexPath) + ":/content", { headers: { Authorization: "Bearer " + t } });
      if (!res.ok) throw new Error("Teile-Index nicht abrufbar (" + res.status + ")");
      text = new TextDecoder("utf-8").decode(await res.arrayBuffer());
      await cacheSet("teile", { at: Date.now(), text: text });
    }
    teileMap = root.KuehlerTeile.parse(text);
    return teileMap;
  }

  /* ---------- Warteschlange für die Bearbeitung im Hintergrund (Weg 2) ---------- */
  async function queueJob(job) {
    var stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
    var name = stamp + "_" + (job.skill || "mail") + "_" + (job.ref || "ohne-ref").replace(/[^A-Za-z0-9-]/g, "") + ".json";
    var bytes = new TextEncoder().encode(JSON.stringify(job, null, 2));
    return send("PUT", "/drives/" + cfg.drives.dataExchange + "/root:/" + encPath(cfg.queuePath + "/" + name) + ":/content?@microsoft.graph.conflictBehavior=rename",
      bytes, { "Content-Type": "application/json" });
  }

  root.KuehlerGraph = {
    configured: configured, history: history, brain: brain,
    projectFolder: projectFolder, ensureSubfolder: ensureSubfolder, upload: upload, mailMime: mailMime,
    plannerTasks: plannerTasks, ruecklauf: ruecklauf, teileIndex: teileIndex, queueJob: queueJob
  };
})(this);
