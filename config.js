/* Kühler-Assistent – zentrale Einstellungen
   Nach jeder Änderung die Dateien neu hochladen; Outlook lädt sie beim nächsten Öffnen. */
(function (root) {
  var config = {
    // Anwendungs-ID aus der Entra-App-Registrierung (setup.mjs trägt sie ein)
    clientId: "f344e357-496c-4240-a81f-65018e75b8e5",

    // Microsoft-Graph-Berechtigungen (nur lesend)
    graphScopes: ["User.Read", "Mail.Read", "Files.Read"],

    // Übergabe an Claude
    claudeUrl: "https://claude.ai/new",
    // true: Auftrag wird zusätzlich als ?q=… mitgegeben (vorausgefüllt, falls Claude das unterstützt)
    claudePrefill: true,

    // Eigene Domains (Mails von hier gelten als intern)
    ownDomains: ["kuehler-gmbh.de", "kuehlerds.de"],

    // Kunden: Erkennung über Absender-Domain oder Namen im Absender/Betreff.
    // Domains bitte bei Bedarf ergänzen bzw. prüfen.
    customers: [
      { name: "Stadler US", domains: [], names: ["Stadler US", "Stadler Rail US", "Salt Lake"] },
      { name: "Stadler Rail", domains: ["stadlerrail.com"], names: ["Stadler"] },
      { name: "GDELS-Mowag", domains: ["gdels.com"], names: ["GDELS", "Mowag"] },
      { name: "Siemens Mobility", domains: ["siemens.com"], names: ["Siemens Mobility"] },
      { name: "VC999", domains: [], names: ["VC999"] }
    ],

    // Mailtypen: Bezeichnung, Übergabe-Skill und begleitende Skills
    types: {
      rfq: {
        label: "Kundenanfrage (RFQ)",
        tone: "rfq",
        task: "Neue Kundenanfrage nach Hausablauf starten: AN-Nummer, Projektstruktur, Sichtung ALLER Zeichnungen, danach Rückfragen zu Umfang und Paketen.",
        skill: "anfragetool",
        skills: ["anfragetool", "kuehler-anfrage-deckblatt", "kuehler-planner-aufgaben", "zweites-gehirn-kuehler"]
      },
      kundenbestellung: {
        label: "Kundenbestellung",
        tone: "kunde",
        task: "Kundenbestellung verarbeiten: Kalkulationsübersicht mit Artikelstamm-Abgleich, Projektordner auf Auftrag umbenennen, Deckblatt Auftrag, AB-Mail vorbereiten.",
        skill: "stadler-bestellungen",
        skills: ["stadler-bestellungen", "kuehler-anfrage-deckblatt", "ab-mail-kunde", "zweites-gehirn-kuehler"]
      },
      lieferanten_ab: {
        label: "Lieferanten-AB",
        tone: "ab",
        task: "Auftragsbestätigung des Lieferanten mit unserer Bestellung abgleichen (Menge, Preis, Termin), Info-Mail „AB eingegangen“ vorbereiten, AB ablegen.",
        skill: "ab-nachfassen-mirko",
        skills: ["ab-nachfassen-mirko", "lieferanten-rueckstaende", "zweites-gehirn-kuehler"]
      },
      lieferantenangebot: {
        label: "Lieferantenangebot",
        tone: "lief",
        task: "Lieferantenangebot einlesen: Positionen in die Kalkulation, Preisbasis-Kalibrierung fortschreiben, Rücklaufliste aktualisieren.",
        skill: "lieferantenangebote-einlesen",
        skills: ["lieferantenangebote-einlesen", "preisbasis-kalibrierung", "lieferanten-nachfassen-mirko"]
      },
      rechnung: {
        label: "Rechnung",
        tone: "sonst",
        task: "Rechnung prüfen und dem Vorgang zuordnen.",
        skill: "",
        skills: ["zweites-gehirn-kuehler"]
      },
      intern: {
        label: "Interne Mail",
        tone: "sonst",
        task: "Interne Mail bearbeiten.",
        skill: "",
        skills: ["zweites-gehirn-kuehler"]
      },
      sonstige: {
        label: "Sonstige Mail",
        tone: "sonst",
        task: "Mail bearbeiten.",
        skill: "",
        skills: ["zweites-gehirn-kuehler"]
      }
    }
  };
  if (typeof module !== "undefined" && module.exports) module.exports = config;
  else root.KUEHLER_CONFIG = config;
})(this);
