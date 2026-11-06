/* Kühler-Assistent – Seitenleiste */
(function () {
  "use strict";
  var cfg = window.KUEHLER_CONFIG;
  var C = window.KuehlerClassifier;
  var G = window.KuehlerGraph;
  var state = { mail: null, result: null, loadId: 0 };

  function $(id) { return document.getElementById(id); }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }
  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

  /* ---------- Einstellungen (wandern mit dem Postfach, je Nutzer) ---------- */
  function getSettings() {
    var rs = Office.context.roamingSettings;
    return { vaultPath: rs.get("vaultPath") || "", vaultName: rs.get("vaultName") || "Kühler" };
  }
  function saveSettings() {
    var rs = Office.context.roamingSettings;
    rs.set("vaultPath", $("setVaultPath").value.trim());
    rs.set("vaultName", $("setVaultName").value.trim() || "Kühler");
    rs.saveAsync(function (r) {
      $("settingsMsg").textContent = r.status === Office.AsyncResultStatus.Succeeded ? "Gespeichert" : "Speichern fehlgeschlagen";
      if (state.mail) loadBrain();
    });
  }
  function fillSettings() {
    var s = getSettings();
    $("setVaultPath").value = s.vaultPath;
    $("setVaultName").value = s.vaultName;
  }

  /* ---------- Mail lesen ---------- */
  function readItem() {
    var item = Office.context.mailbox.item;
    if (!item) return Promise.resolve(null);
    return new Promise(function (resolve) {
      item.body.getAsync(Office.CoercionType.Text, function (r) {
        var body = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "";
        var from = item.from || item.sender || {};
        resolve({
          subject: item.subject || "",
          fromName: from.displayName || "",
          fromAddress: from.emailAddress || "",
          received: item.dateTimeCreated,
          messageId: item.internetMessageId || "",
          body: body,
          attachments: (item.attachments || []).filter(function (a) { return !a.isInline; }).map(function (a) { return a.name; })
        });
      });
    });
  }

  /* ---------- Darstellung ---------- */
  function renderResult() {
    var r = state.result;
    var badge = $("typeBadge");
    badge.textContent = r.label;
    badge.className = "badge t-" + r.tone;
    $("vorgang").textContent = r.vorgang;
    $("partner").textContent = (r.via ? r.partner : (state.mail.fromName || state.mail.fromAddress)) + (state.mail.attachments.length ? " · " + state.mail.attachments.length + (state.mail.attachments.length === 1 ? " Anhang" : " Anhänge") : "");
    $("task").textContent = r.task;
    var sk = $("skills"); clear(sk);
    r.skills.forEach(function (s) { sk.appendChild(el("span", "skill", s)); });
    $("btnHandoff").textContent = r.skill ? "An Claude übergeben (/" + r.skill + ")" : "An Claude übergeben";
    $("handoffMsg").hidden = true;
  }

  function showError(container, prefix, err) {
    clear(container);
    container.appendChild(el("div", "err", prefix + (err && err.message ? err.message : String(err))));
  }

  function openUrl(url) {
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* weiter mit window.open */ }
    window.open(url, "_blank", "noopener");
  }

  /* ---------- Kühler-Gehirn ---------- */
  function loadBrain() {
    var id = state.loadId;
    var box = $("brain");
    var settings = getSettings();
    if (!settings.vaultPath) {
      $("brainCard").hidden = true;
      return;
    }
    $("brainCard").hidden = false;
    if (!G.configured()) { showError(box, "", "Anmeldung noch nicht eingerichtet (Anwendungs-ID fehlt in config.js)."); return; }
    clear(box); box.appendChild(el("span", "muted small", "Wird geladen …"));
    G.brain(state.mail, state.result, settings).then(function (notes) {
      if (id !== state.loadId) return;
      clear(box);
      if (!notes || !notes.length) { box.appendChild(el("span", "muted small", "Keine passende Notiz gefunden.")); return; }
      notes.forEach(function (n) {
        var d = el("div", "note");
        var a = el("a", "", n.title);
        a.href = n.obsidian;
        a.title = "In Obsidian öffnen";
        a.addEventListener("click", function (ev) { ev.preventDefault(); openUrl(n.obsidian); });
        d.appendChild(a);
        if (n.text) d.appendChild(el("p", "", n.text));
        box.appendChild(d);
      });
    }).catch(function (e) { if (id === state.loadId) showError(box, "Kühler-Gehirn nicht erreichbar: ", e); });
  }

  /* ---------- Übergabe an Claude ---------- */
  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text).then(function () { return true; }, function () { return legacyCopy(text); });
    }
    return Promise.resolve(legacyCopy(text));
  }
  function legacyCopy(text) {
    var t = $("clip");
    t.value = text;
    t.select();
    try { return document.execCommand("copy"); } catch (e) { return false; }
  }
  function openDesktop(url) {
    // Eigenes URL-Schema der Claude-Desktop-App; Office.context.ui.openBrowserWindow akzeptiert nur https
    var w = null;
    try { w = window.open(url, "_blank"); } catch (e) { w = null; }
    if (!w) {
      try { var a = document.createElement("a"); a.href = url; a.rel = "noopener"; document.body.appendChild(a); a.click(); a.remove(); } catch (e2) { /* Link unten bleibt */ }
    }
  }
  function handoff(prompt, msgBox) {
    var q = cfg.claudePrefill ? "?q=" + encodeURIComponent(prompt.slice(0, 13500)) : "";
    var desktopUrl = cfg.claudeDesktopUrl + q;
    var webUrl = cfg.claudeWebUrl + q;
    copyText(prompt).then(function (copied) {
      openDesktop(desktopUrl);
      msgBox.hidden = false; msgBox.className = "ok";
      clear(msgBox);
      msgBox.appendChild(el("div", "", "Claude-App wird geöffnet, der Auftrag steht im Eingabefeld." +
        (copied ? " Er liegt zusätzlich in der Zwischenablage (Cmd+V bzw. Strg+V)." : "")));
      var row = el("div", "row wrap");
      row.style.marginTop = "6px";
      var a1 = el("a", "", "Claude-App öffnen");
      a1.href = desktopUrl;
      a1.addEventListener("click", function (ev) { ev.preventDefault(); openDesktop(desktopUrl); });
      var a2 = el("a", "", "Stattdessen im Browser öffnen");
      a2.href = webUrl;
      a2.addEventListener("click", function (ev) { ev.preventDefault(); openUrl(webUrl); });
      row.appendChild(a1); row.appendChild(el("span", "muted", "·")); row.appendChild(a2);
      msgBox.appendChild(row);
      if (!copied) {
        var pre = el("textarea", "");
        pre.rows = 6; pre.readOnly = true; pre.value = prompt;
        msgBox.appendChild(pre);
      }
    });
  }

  /* ---------- Ablauf ---------- */
  function load() {
    state.loadId++;
    $("loading").hidden = false;
    $("content").hidden = true;
    readItem().then(function (mail) {
      $("loading").hidden = true;
      if (!mail) { $("loading").hidden = false; $("loading").textContent = "Bitte eine Mail auswählen."; return; }
      state.mail = mail;
      state.result = C.classify(mail, cfg);
      $("content").hidden = false;
      renderResult();
      loadBrain();
    });
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      $("loading").textContent = "Dieses Add-in läuft nur in Outlook.";
      return;
    }
    fillSettings();
    $("btnSettings").addEventListener("click", function () {
      var s = $("settings"); s.hidden = !s.hidden; $("settingsMsg").textContent = "";
    });
    $("btnSaveSettings").addEventListener("click", saveSettings);
    $("btnHandoff").addEventListener("click", function () {
      handoff(C.buildHandoffPrompt(state.mail, state.result), $("handoffMsg"));
    });
    // Angeheftete Seitenleiste: bei Wechsel der Mail neu laden
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, load); } catch (e) { /* ältere Clients */ }
    load();
  });
})();
