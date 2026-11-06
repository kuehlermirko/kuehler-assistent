/* Kühler-Assistent – zentrale Einstellungen
   Nach jeder Änderung die Dateien neu hochladen; Outlook lädt sie beim nächsten Öffnen. */
(function (root) {
  var config = {
    // Anwendungs-ID aus der Entra-App-Registrierung (setup.mjs trägt sie ein)
    clientId: "f344e357-496c-4240-a81f-65018e75b8e5",

    // Microsoft-Graph-Berechtigungen (nur lesend)
    graphScopes: ["User.Read", "Mail.Read", "Files.Read"],

    // Übergabe an Claude: zuerst die Desktop-App (claude://), Browser nur als Ausweichweg
    claudeDesktopUrl: "claude://claude.ai/new",
    claudeWebUrl: "https://claude.ai/new",
    // true: Auftrag wird als ?q=… vorausgefüllt (max. ca. 14.000 Zeichen)
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
      rfq_gdels: {
        label: "Kundenanfrage GDELS-Mowag",
        tone: "rfq",
        task: "Neue GDELS-Mowag-Anfrage nach Hausablauf starten: AN-Nummer, Projektstruktur, Sichtung ALLER Zeichnungen, danach Rückfragen zu Umfang und Paketen.",
        skill: "anfragetool",
        skills: ["anfragetool", "anfragetool-gdels", "kuehler-anfrage-deckblatt", "kuehler-planner-aufgaben"]
      },
      richtpreis: {
        label: "Richtpreisanfrage",
        tone: "rfq",
        task: "Richtpreis-Kalkulation aus Zeichnung und STEP erstellen; vorher AN-Nr., Stückzahl und Ansätze klären und Faktor aus der Preisbasis-Kalibrierung nennen.",
        skill: "richtpreis-kalkulation",
        skills: ["richtpreis-kalkulation", "preisbasis-kalibrierung", "zweites-gehirn-kuehler"]
      },
      kunden_ab: {
        label: "Auftrag / AU zum Kunden",
        tone: "kunde",
        task: "Stadler-Kalkulationstabelle zur Auftragsbestätigung (AU) erstellen bzw. aktualisieren und AB-Mail an den Kunden vorbereiten.",
        skill: "stadler-kalkulationsexcel",
        skills: ["stadler-kalkulationsexcel", "ab-mail-kunde", "zweites-gehirn-kuehler"]
      },
      spezifikation: {
        label: "Stadler-Spezifikation",
        tone: "kunde",
        task: "Stadler-Spezifikation 1:1 als Kühler-Dokument (DEU/EN, ohne Stadler-Spur) umschreiben.",
        skill: "stadler-spez-umschreiben",
        skills: ["stadler-spez-umschreiben"]
      },
      rueckstaende: {
        label: "Rückstandsliste",
        tone: "kunde",
        task: "Rückstandsliste auswerten; vorher Stichtage für den sicheren Umsatz vorschlagen und Freigabe abwarten.",
        skill: "stadler-rueckstaende",
        skills: ["stadler-rueckstaende"]
      },
      fai: {
        label: "FAI / Erstmuster",
        tone: "kunde",
        task: "FAI nach EN 15085 CL1 bzw. Kundenvorgabe vorbereiten.",
        skill: "kuehler-fai-cl1",
        skills: ["kuehler-fai-cl1", "zweites-gehirn-kuehler"]
      },
      lieferant_rueckmeldung: {
        label: "Lieferanten-Rückmeldung",
        tone: "lief",
        task: "Rückmeldung bzw. Absage des Lieferanten in der Rücklaufliste und Lieferantenbewertung erfassen.",
        skill: "lieferanten-nachfassen-mirko",
        skills: ["lieferanten-nachfassen-mirko", "lieferantenpool-mirko"]
      },
      werkstatt: {
        label: "Werkstatt / Montage",
        tone: "sonst",
        task: "Werkstattauftrag und Montage zum Kundenauftrag bearbeiten (Ordner 08_Montage_Kühler, Werkstattbestellung, Werkauftrag, Montageplan).",
        skill: "kuehler-werkstattauftrag-montage",
        skills: ["kuehler-werkstattauftrag-montage", "selectline-bestellung"]
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
    },

    // Auswahl „Anderer Skill“ in der Seitenleiste (Reihenfolge = Anzeige)
    skillMenu: [
      ["anfragetool", "Kundenanfrage bearbeiten (Anfragetool)"],
      ["stadler-bestellungen", "Kundenbestellung Stadler (EB/PO/4500)"],
      ["stadler-kalkulationsexcel", "Auftrag AU – Kalkulationstabelle"],
      ["ab-mail-kunde", "AB-Mail an Kunden"],
      ["richtpreis-kalkulation", "Richtpreis-Kalkulation"],
      ["lieferantenpool-mirko", "Lieferantenanfrage versenden"],
      ["lieferantenangebote-einlesen", "Lieferantenangebot einlesen"],
      ["lieferanten-nachfassen-mirko", "Lieferanten nachfassen / Rücklauf"],
      ["angebotsueberwachung-mirko", "Angebot beim Kunden nachfassen"],
      ["selectline-bestellung", "Bestellung an Lieferant (SelectLine)"],
      ["ab-nachfassen-mirko", "Lieferanten-AB prüfen / nachfassen"],
      ["lieferanten-rueckstaende", "Lieferanten-Rückstände"],
      ["stadler-rueckstaende", "Rückstandsliste Stadler"],
      ["kuehler-werkstattauftrag-montage", "Werkstattauftrag / Montage"],
      ["kuehler-fai-cl1", "FAI / Erstmuster"],
      ["stadler-spez-umschreiben", "Stadler-Spezifikation umschreiben"],
      ["projektmanagement", "Projektakte / Projektmanagement"],
      ["kuehler-briefpapier", "Dokument auf Kühler-Briefpapier"],
      ["zweites-gehirn-kuehler", "Nur Kühler-Gehirn fortschreiben"]
    ]
  };
  if (typeof module !== "undefined" && module.exports) module.exports = config;
  else root.KUEHLER_CONFIG = config;
})(this);
