/* Kühler-Assistent – Seitenleiste */
(function () {
  "use strict";
  var cfg = window.KUEHLER_CONFIG;
  var C = window.KuehlerClassifier;
  var G = window.KuehlerGraph;
  var X = window.KuehlerFristen;
  var state = { mail: null, result: null, handoff: null, folder: null, loadId: 0 };

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
          restId: (function () { try { return Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { return ""; } })(),
          attachmentItems: (item.attachments || []).filter(function (a) { return !a.isInline; }).map(function (a) { return { id: a.id, name: a.name, size: a.size, type: a.attachmentType }; }),
          attachments: (item.attachments || []).filter(function (a) { return !a.isInline; }).map(function (a) { return a.name; })
        });
      });
    });
  }

  /* ---------- Darstellung ---------- */
  /* ---------- Lernen: Skill-Korrekturen je Absender merken ---------- */
  function learned() {
    try { return JSON.parse(Office.context.roamingSettings.get("learned") || "{}"); } catch (e) { return {}; }
  }
  function senderKeys() {
    var adr = (state.mail.fromAddress || "").toLowerCase();
    var orig = state.result.via ? (state.result.partner || "") : "";
    return { adr: adr, dom: C.domainOf(adr), via: orig };
  }
  function learnedSkill() {
    var L = learned(), k = senderKeys();
    var hit = L["a:" + k.adr] || L["d:" + k.dom];
    if (hit && hit.type && hit.type !== state.result.type) return null; // nur für denselben erkannten Mailtyp
    return hit ? hit.skill : null;
  }
  function saveLearned(skill) {
    var L = learned(), k = senderKeys();
    // Kunden je Person (dort gibt es viele Ansprechpartner), Lieferanten je Firma (Domain), Freemail je Adresse
    var perDomain = !state.result.customer && k.dom && !/(gmail|t-online|web|gmx|outlook|hotmail|yahoo|icloud)\./.test(k.dom);
    var key = perDomain ? "d:" + k.dom : "a:" + k.adr;
    if (!skill || skill === state.result.skill) delete L[key];
    else L[key] = { skill: skill, type: state.result.type, at: new Date().toISOString().slice(0, 10) };
    var keys = Object.keys(L);
    if (keys.length > 200) keys.sort(function (a, b) { return (L[a].at || "") < (L[b].at || "") ? -1 : 1; }).slice(0, keys.length - 200).forEach(function (x) { delete L[x]; });
    Office.context.roamingSettings.set("learned", JSON.stringify(L));
    Office.context.roamingSettings.saveAsync(function () { fillLearnInfo(); });
  }
  function fillLearnInfo() {
    var n = Object.keys(learned()).length;
    $("learnInfo").textContent = n ? n + " Absender mit eigener Skill-Zuordnung." : "Noch nichts gelernt.";
  }

  /* Aktiver Übergabe-Auftrag: erkannter Typ oder manuell gewählter Skill */
  function renderHandoff() {
    var h = state.handoff;
    $("task").textContent = h.task;
    var sk = $("skills"); clear(sk);
    h.skills.forEach(function (s) { sk.appendChild(el("span", "skill", s)); });
    $("btnHandoff").textContent = h.skill ? "An Claude übergeben (/" + h.skill + ")" : "An Claude übergeben";
    $("handoffMsg").hidden = true;
  }
  function fillSkillSelect() {
    var sel = $("skillSelect"); clear(sel);
    var first = el("option", "", "Wie erkannt" + (state.result.skill ? " (/" + state.result.skill + ")" : ""));
    first.value = "";
    sel.appendChild(first);
    (cfg.skillMenu || []).forEach(function (m) {
      var o = el("option", "", m[1] + "  (/" + m[0] + ")");
      o.value = m[0];
      sel.appendChild(o);
    });
    sel.value = "";
  }
  function pickSkill(id) {
    var r = state.result;
    if (!id) { state.handoff = r; }
    else {
      var label = id;
      (cfg.skillMenu || []).forEach(function (m) { if (m[0] === id) label = m[1]; });
      var copy = {};
      Object.keys(r).forEach(function (k) { copy[k] = r[k]; });
      copy.skill = id;
      copy.skills = [id, "zweites-gehirn-kuehler"].filter(function (v, i, a) { return a.indexOf(v) === i; });
      copy.task = label + " – aus dieser Mail heraus nach Hausablauf bearbeiten.";
      state.handoff = copy;
    }
    renderHandoff();
  }

  function renderResult() {
    var r = state.result;
    var badge = $("typeBadge");
    badge.textContent = r.label;
    badge.className = "badge t-" + r.tone;
    $("vorgang").textContent = r.vorgang;
    $("partner").textContent = (r.via ? r.partner : (state.mail.fromName || state.mail.fromAddress)) + (state.mail.attachments.length ? " · " + state.mail.attachments.length + (state.mail.attachments.length === 1 ? " Anhang" : " Anhänge") : "");
    fillSkillSelect();
    state.handoff = r;
    var ls = learnedSkill();
    $("learnNote").hidden = true;
    if (ls && ls !== r.skill) {
      $("skillSelect").value = ls;
      pickSkill(ls);
      $("learnNote").hidden = false;
      $("learnNote").textContent = "Gelernt: Bei diesem Absender hast du zuletzt /" + ls + " gewählt.";
    } else renderHandoff();
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

  /* ---------- Projektordner und Planner ---------- */
  function fmtDay(d) {
    var x = d instanceof Date ? d : new Date(d);
    if (isNaN(x)) return "";
    function z(v) { return (v < 10 ? "0" : "") + v; }
    return z(x.getDate()) + "." + z(x.getMonth() + 1) + "." + x.getFullYear();
  }
  function defaultTarget(type, folder) {
    var want = { rfq: "01_Kundenanfrage", rfq_gdels: "01_Kundenanfrage", richtpreis: "01_Kundenanfrage", spezifikation: "01_Kundenanfrage",
      kundenbestellung: "07_Auftragsbearbeitung", kunden_ab: "07_Auftragsbearbeitung", lieferanten_ab: "07_Auftragsbearbeitung", fai: "07_Auftragsbearbeitung",
      lieferantenangebot: "04_Anfragen", lieferant_rueckmeldung: "04_Anfragen", lieferanten_korrespondenz: "04_Anfragen", lieferanten_bestellung: "07_Auftragsbearbeitung" }[type] || "_Mail-Eingang";
    return want;
  }
  function loadProjekt() {
    var id = state.loadId, box = $("projekt");
    clear(box);
    $("ablageCard").hidden = true;
    var r = state.result;
    var hasRef = r.refs.an.length || r.refs.po.length || (r.refs.po2 || []).length || r.refs.eb.length || r.refs.sap.length || (r.refs.gdels || []).length;
    if (!G.configured() || !hasRef) { renderAblage(null); return; }
    box.appendChild(el("span", "muted small", "Projektordner wird gesucht …"));
    G.projectFolder(r).then(function (f) {
      if (id !== state.loadId) return;
      state.folder = f;
      clear(box);
      if (!f) { box.appendChild(el("span", "muted small", "Kein Projektordner zu " + r.refs.all.slice(0, 2).join(", ") + " gefunden.")); renderAblage(null); return; }
      var d = el("div", "proj");
      var a = el("a", "", "Projektordner: " + f.name);
      a.href = f.webUrl; a.addEventListener("click", function (ev) { ev.preventDefault(); openUrl(f.webUrl); });
      d.appendChild(a);
      box.appendChild(d);
      renderAblage(f);
      var an = (f.name.match(/^AN\d{7}/) || [])[0];
      if (an && cfg.planner.planId) loadPlanner(an, d, id);
    }).catch(function (e) { if (id === state.loadId) { showError(box, "Projektordner nicht abrufbar: ", e); renderAblage(null); } });
  }
  function loadPlanner(an, holder, id) {
    G.plannerTasks(an).then(function (p) {
      if (id !== state.loadId) return;
      if (!p.bucket) { holder.appendChild(el("span", "muted small", "Kein Planner-Bucket zu " + an + ".")); return; }
      var link = el("a", "", "Planner: " + (p.tasks.length ? p.tasks.length + " offene Aufgabe" + (p.tasks.length === 1 ? "" : "n") : "alles erledigt"));
      var url = "https://planner.cloud.microsoft/ui/web/plan/" + cfg.planner.planId + "/view/board";
      link.href = url; link.addEventListener("click", function (ev) { ev.preventDefault(); openUrl(url); });
      holder.appendChild(link);
      var today = new Date(); today.setHours(0, 0, 0, 0);
      p.tasks.slice(0, 5).forEach(function (t) {
        var row = el("div", "task-row");
        var due = t.due ? new Date(t.due) : null;
        var dd = el("span", "d" + (due && due < today ? " late" : ""), due ? fmtDay(due) : "ohne Termin");
        row.appendChild(dd); row.appendChild(el("span", "", t.title));
        holder.appendChild(row);
      });
    }).catch(function (e) { if (id === state.loadId) holder.appendChild(el("span", "muted small", "Planner nicht abrufbar: " + (e.message || e))); });
  }

  /* ---------- Anhänge ablegen ---------- */
  function renderAblage(folder) {
    var list = $("ablageList"); clear(list);
    var items = state.mail.attachmentItems || [];
    $("ablageMsg").hidden = true;
    if (!folder) { $("ablageCard").hidden = true; return; }
    $("ablageCard").hidden = false;
    if (!items.length) list.appendChild(el("span", "muted small", "Keine Anhänge – nur die Mail selbst kann abgelegt werden."));
    items.forEach(function (a, i) {
      var lab = el("label", "check");
      var cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = true; cb.dataset.idx = i;
      lab.appendChild(cb);
      lab.appendChild(document.createTextNode(" " + a.name + (a.size ? " (" + Math.max(1, Math.round(a.size / 1024)) + " KB)" : "")));
      list.appendChild(lab);
    });
    var sel = $("ablageZiel"); clear(sel);
    var want = defaultTarget(state.result.type, folder);
    var names = folder.subfolders.map(function (s) { return s.name; });
    if (names.indexOf(want) < 0) names.push(want);
    names.sort().forEach(function (n) {
      var o = el("option", "", n + (folder.subfolders.some(function (s) { return s.name === n; }) ? "" : " (wird angelegt)"));
      o.value = n; sel.appendChild(o);
    });
    sel.value = want;
  }
  function attachmentBytes(id) {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.getAttachmentContentAsync(id, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error(r.error && r.error.message || "Anhang nicht lesbar")); return; }
        var v = r.value;
        if (v.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          var bin = atob(v.content), out = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
          resolve(out);
        } else if (v.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
          resolve(new TextEncoder().encode(v.content));
        } else reject(new Error("Cloud-Anhang (Link) – bitte über OneDrive ablegen"));
      });
    });
  }
  async function doAblage() {
    var msg = $("ablageMsg"), btn = $("btnAblage");
    var f = state.folder; if (!f) return;
    var picks = Array.prototype.slice.call(document.querySelectorAll("#ablageList input[type=checkbox]")).filter(function (c) { return c.checked; })
      .map(function (c) { return state.mail.attachmentItems[+c.dataset.idx]; });
    var withEml = $("ablageEml").checked;
    if (!picks.length && !withEml) { msg.hidden = false; msg.className = "err"; msg.textContent = "Bitte mindestens eine Datei auswählen."; return; }
    btn.disabled = true; msg.hidden = false; msg.className = "ok"; msg.textContent = "Wird abgelegt …";
    var done = [], failed = [];
    try {
      var target = await G.ensureSubfolder(f, $("ablageZiel").value);
      for (var i = 0; i < picks.length; i++) {
        try {
          var a = picks[i], bytes = await attachmentBytes(a.id);
          var name = a.type === "item" && !/\.eml$/i.test(a.name) ? a.name + ".eml" : a.name;
          await G.upload(f.driveId, target.id, name, bytes); done.push(name);
        } catch (e) { failed.push(picks[i].name + " (" + (e.message || e) + ")"); }
      }
      if (withEml) {
        try {
          if (!state.mail.restId) throw new Error("Mail-ID fehlt");
          var d = new Date(state.mail.received || Date.now());
          var nm = d.toISOString().slice(0, 10) + "_" + (state.mail.subject || "Mail").replace(/^(AW|WG|RE|FW|Fwd)\s*:\s*/gi, "").slice(0, 80) + ".eml";
          await G.upload(f.driveId, target.id, nm, await G.mailMime(state.mail.restId)); done.push(nm);
        } catch (e) { failed.push("Mail als .eml (" + (e.message || e) + ")"); }
      }
      clear(msg);
      msg.className = failed.length ? "err" : "ok";
      msg.appendChild(el("div", "", done.length ? done.length + " Datei" + (done.length === 1 ? "" : "en") + " abgelegt in " + target.name + "." : "Nichts abgelegt."));
      if (failed.length) msg.appendChild(el("div", "", "Nicht abgelegt: " + failed.join("; ")));
      if (target.webUrl) {
        var a2 = el("a", "", "Ordner öffnen"); a2.href = target.webUrl;
        a2.addEventListener("click", function (ev) { ev.preventDefault(); openUrl(target.webUrl); });
        msg.appendChild(a2);
      }
    } catch (e) {
      msg.className = "err"; msg.textContent = "Ablage fehlgeschlagen: " + (e.message || e);
    }
    btn.disabled = false;
  }

  /* ---------- Fristen ---------- */
  function renderFristen() {
    var box = $("fristen"); clear(box);
    $("fristMsg").hidden = true;
    var list = X.deadlines(state.mail);
    $("fristCard").hidden = !list.length;
    list.forEach(function (f) {
      var row = el("div", "frist");
      row.appendChild(el("span", "when" + (f.strong ? "" : " weak"), fmtDay(f.date)));
      row.appendChild(el("span", "ctx", f.snippet));
      var b = el("button", "secondary", "Erinnerung");
      b.type = "button";
      b.addEventListener("click", function () { remind(f); });
      row.appendChild(b);
      box.appendChild(row);
    });
  }
  function remind(f) {
    var start = new Date(f.date), end = new Date(start.getTime() + 30 * 60000);
    var r = state.result, subj = "Frist: " + (state.mail.subject || "Mail") ;
    try {
      Office.context.mailbox.displayNewAppointmentForm({
        start: start, end: end,
        subject: subj.slice(0, 250),
        body: "Frist aus der Mail von " + (state.mail.fromName || state.mail.fromAddress) + " (" + fmtDay(state.mail.received) + "):\n" + f.snippet +
          "\n\nVorgang: " + r.vorgang
      });
      var m = $("fristMsg"); m.hidden = false; m.className = "ok"; m.textContent = "Termin-Fenster geöffnet – bitte prüfen und speichern.";
    } catch (e) {
      var m2 = $("fristMsg"); m2.hidden = false; m2.className = "err"; m2.textContent = "Kalender konnte nicht geöffnet werden: " + (e.message || e);
    }
  }

  /* ---------- Lieferanten-Rücklauf ---------- */
  function statusClass(s) {
    s = String(s || "").toLowerCase();
    if (s.indexOf("angebot") === 0) return "st-ok";
    if (s.indexOf("absage") === 0 || s.indexOf("unzustell") === 0) return "st-no";
    return "st-open";
  }
  function excelDate(v) {
    if (typeof v === "number" && v > 30000) { var d = new Date(Math.round((v - 25569) * 86400000)); return fmtDay(d); }
    var s = String(v || ""); var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + "." + m[2] + "." + m[1] : s;
  }
  function loadRuecklauf() {
    var id = state.loadId, box = $("rueck"), r = state.result;
    clear(box);
    var isSupplier = !r.customer && r.type !== "intern";
    if (!G.configured() || !isSupplier) { $("rueckCard").hidden = true; return; }
    $("rueckCard").hidden = false;
    box.appendChild(el("span", "muted small", "Rücklaufliste wird gelesen …"));
    G.ruecklauf(C.domainOf(state.mail.fromAddress), state.mail.fromAddress, r.refs.an).then(function (res) {
      if (id !== state.loadId) return;
      clear(box);
      if (!res.lieferant.length && !res.pakete.length) { $("rueckCard").hidden = true; return; }
      if (res.lieferant.length) {
        var t = el("table", "mini");
        var h = el("tr"); ["AN / Gewerk", "Frist", "Status", "Eingang"].forEach(function (x) { h.appendChild(el("th", "", x)); }); t.appendChild(h);
        res.lieferant.forEach(function (row) {
          var tr = el("tr");
          tr.appendChild(el("td", "", row["AN-Nr."] + " · " + row["Gewerk"]));
          tr.appendChild(el("td", "", excelDate(row["Angebotsfrist"])));
          tr.appendChild(el("td", statusClass(row["Status"]), row["Status"] || "offen"));
          tr.appendChild(el("td", "", excelDate(row["Eingang am"])));
          t.appendChild(tr);
        });
        box.appendChild(t);
      }
      res.pakete.forEach(function (p) {
        box.appendChild(el("div", "small", "Paket " + p["AN-Nr."] + "_" + p["Gewerk"] + ": " + p["angefragt"] + " angefragt · " + p["Angebot"] + " Angebote · " +
          p["Absage"] + " Absagen · " + p["offen"] + " offen · Frist " + excelDate(p["Angebotsfrist"]) + " · " + (p["Status Paket"] || "")));
      });
    }).catch(function (e) { if (id === state.loadId) showError(box, "Rücklaufliste nicht lesbar: ", e); });
  }

  /* ---------- Teile im Artikelstamm ---------- */
  function checkTeile(force) {
    var box = $("teile"), id = state.loadId;
    clear(box); box.appendChild(el("span", "muted small", "Artikelstamm wird abgeglichen …"));
    G.teileIndex(force).then(function (map) {
      if (id !== state.loadId) return;
      var res = X.matchParts(state.mail, map);
      clear(box);
      if (!res.hits.length) {
        box.appendChild(el("span", "muted small", res.checked ? res.checked + " Nummern geprüft – keine im Artikelstamm bekannt (neue Teile?)." : "Keine Teilenummern in der Mail gefunden."));
        return;
      }
      box.appendChild(el("span", "muted small", res.hits.length + " von " + res.checked + " Nummern bekannt:"));
      res.hits.slice(0, 15).forEach(function (h) {
        var d = el("div", "part");
        var top = el("div"); top.appendChild(el("b", "", h.nummer)); top.appendChild(document.createTextNode(h.kid ? "  →  KID " + h.kid : "  →  keine KID"));
        d.appendChild(top);
        if (h.bezeichnung) d.appendChild(el("div", "", h.bezeichnung));
        if (h.an) d.appendChild(el("div", "muted", "Bisher: " + h.an + (h.von ? " (" + h.von + (h.bis && h.bis !== h.von ? "–" + h.bis : "") + ")" : "")));
        if (h.rma) d.appendChild(el("div", "rma", "Reklamation: " + h.rma));
        box.appendChild(d);
      });
    }).catch(function (e) {
      if (id !== state.loadId) return;
      showError(box, "Artikelstamm nicht abrufbar: ", e);
    });
  }
  function resetTeile() {
    var box = $("teile"); clear(box);
    var b = el("button", "secondary wide", "Nummern aus der Mail prüfen"); b.id = "btnTeile";
    b.type = "button"; b.addEventListener("click", function () { checkTeile(false); });
    box.appendChild(b);
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

  /* ---------- Weg 2: Auftrag in die Warteschlange ---------- */
  function queue() {
    var msg = $("handoffMsg"), h = state.handoff || state.result;
    msg.hidden = false; msg.className = "ok"; msg.textContent = "Auftrag wird in die Warteschlange gestellt …";
    var job = {
      version: 1,
      erstellt: new Date().toISOString(),
      von: Office.context.mailbox.userProfile ? Office.context.mailbox.userProfile.emailAddress : "",
      mail: { betreff: state.mail.subject, absender: state.mail.fromName, adresse: state.mail.fromAddress,
              empfangen: state.mail.received, messageId: state.mail.messageId, restId: state.mail.restId, anhaenge: state.mail.attachments },
      erkannt: { typ: h.type, label: h.label, vorgang: h.vorgang, kunde: h.customer, nummern: h.refs.all },
      projektordner: state.folder ? state.folder.webUrl : null,
      skill: h.skill, auftrag: h.task,
      prompt: C.buildHandoffPrompt(state.mail, h),
      ref: (h.refs.all[0] || "")
    };
    G.queueJob(job).then(function () {
      msg.textContent = "Auftrag liegt in der Warteschlange und wird beim nächsten Hintergrund-Lauf bearbeitet. Ergebnis und Rückfragen kommen per Mail an dich.";
      if (job.skill) saveLearned(($("skillSelect").value || ""));
    }).catch(function (e) { msg.className = "err"; msg.textContent = "Warteschlange nicht erreichbar: " + (e.message || e); });
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
      state.folder = null;
      renderResult();
      renderFristen();
      resetTeile();
      loadProjekt();
      loadRuecklauf();
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
    $("skillSelect").addEventListener("change", function () { pickSkill(this.value); });
    $("btnHandoff").addEventListener("click", function () {
      saveLearned($("skillSelect").value || "");
      handoff(C.buildHandoffPrompt(state.mail, state.handoff || state.result), $("handoffMsg"));
    });
    $("btnQueue").addEventListener("click", queue);
    $("btnAblage").addEventListener("click", doAblage);
    $("btnForget").addEventListener("click", function () {
      Office.context.roamingSettings.remove("learned");
      Office.context.roamingSettings.saveAsync(function () { fillLearnInfo(); });
    });
    fillLearnInfo();
    // Angeheftete Seitenleiste: bei Wechsel der Mail neu laden
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, load); } catch (e) { /* ältere Clients */ }
    load();
  });
})();
