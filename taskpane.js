/* Kühler-Assistent – Seitenleiste */
(function () {
  "use strict";
  var cfg = window.KUEHLER_CONFIG;
  var C = window.KuehlerClassifier;
  var G = window.KuehlerGraph;
  var state = { mail: null, result: null, tone: "Formell", language: "Deutsch", loadId: 0 };

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
    $("partner").textContent = (state.mail.fromName || state.mail.fromAddress) + (state.mail.attachments.length ? " · " + state.mail.attachments.length + (state.mail.attachments.length === 1 ? " Anhang" : " Anhänge") : "");
    $("task").textContent = r.task;
    var sk = $("skills"); clear(sk);
    r.skills.forEach(function (s) { sk.appendChild(el("span", "skill", s)); });
    $("btnHandoff").textContent = r.skill ? "An Claude übergeben (/" + r.skill + ")" : "An Claude übergeben";
    $("handoffMsg").hidden = true;
    $("replyMsg").hidden = true;
    $("replyText").value = "";
    $("replyAttach").value = "";
    state.tone = (r.type === "intern") ? "Kurz" : "Formell";
    state.language = r.language;
    renderChips();
  }

  function chipGroup(container, values, current, onPick) {
    clear(container);
    values.forEach(function (v) {
      var b = el("button", "chip", v);
      b.type = "button";
      b.setAttribute("aria-pressed", v === current ? "true" : "false");
      b.addEventListener("click", function () { onPick(v); });
      container.appendChild(b);
    });
  }
  function renderChips() {
    chipGroup($("toneChips"), ["Formell", "Kurz", "Freundschaftlich"], state.tone, function (v) { state.tone = v; renderChips(); });
    chipGroup($("langChips"), ["Deutsch", "Englisch"], state.language, function (v) { state.language = v; renderChips(); });
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

  /* ---------- Frühere Mails ---------- */
  function loadHistory() {
    var id = state.loadId;
    var box = $("history");
    if (!G.configured()) { showError(box, "", "Anmeldung noch nicht eingerichtet (Anwendungs-ID fehlt in config.js)."); return; }
    clear(box); box.appendChild(el("span", "muted small", "Wird geladen …"));
    G.history(state.mail, state.result).then(function (items) {
      if (id !== state.loadId) return;
      clear(box);
      if (!items.length) { box.appendChild(el("span", "muted small", "Keine früheren Mails gefunden.")); return; }
      items.forEach(function (m) {
        var row = el("div", "hist");
        row.appendChild(el("span", "d", C.fmtDate(m.date).split(",")[0]));
        var a = el("a", "", m.subject + (m.from ? " – " + m.from : ""));
        a.href = m.link || "#";
        a.addEventListener("click", function (ev) { ev.preventDefault(); if (m.link) openUrl(m.link); });
        row.appendChild(a);
        box.appendChild(row);
      });
    }).catch(function (e) { if (id === state.loadId) showError(box, "Mailverlauf nicht abrufbar: ", e); });
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
  function handoff(prompt, msgBox) {
    copyText(prompt).then(function (copied) {
      var url = cfg.claudeUrl + (cfg.claudePrefill ? "?q=" + encodeURIComponent(prompt) : "");
      openUrl(url);
      msgBox.hidden = false;
      msgBox.textContent = copied
        ? "Claude wird geöffnet. Der Auftrag liegt zusätzlich in der Zwischenablage – falls das Eingabefeld leer ist, mit Cmd+V bzw. Strg+V einfügen."
        : "Claude wird geöffnet. Bitte den Auftrag unten markieren und kopieren:";
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
      loadHistory();
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
    $("btnReply").addEventListener("click", function () {
      handoff(C.buildReplyPrompt(state.mail, state.result, {
        tone: state.tone,
        language: state.language,
        instruction: $("replyText").value.trim(),
        attachments: $("replyAttach").value.trim()
      }), $("replyMsg"));
    });
    // Angeheftete Seitenleiste: bei Wechsel der Mail neu laden
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, load); } catch (e) { /* ältere Clients */ }
    load();
  });
})();
