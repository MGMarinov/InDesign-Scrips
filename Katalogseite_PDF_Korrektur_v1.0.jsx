/*******************************************************************************
 *
 * Katalogseite PDF Korrektur
 *
 * Beschreibung:
 * Dieses Skript automatisiert den Export von Katalogseiten-PDFs aus InDesign.
 * Es analysiert die Ebenenstruktur des aktiven Dokuments, um festzustellen,
 * welche Versionen (AD, CD, CHFR) exportiert werden können. Basierend auf den
 * gefundenen Ebenen schaltet es die Sichtbarkeit um und exportiert die PDFs
 * mit vordefinierten Exporteinstellungen. Das Skript findet oder erstellt
 * automatisch den richtigen Zielordner basierend auf dem Dateinamen und
 * erstellt zusätzlich eine Intranet-Version für die Freigabe.
 *
 * Verwendung:
 * 1. Öffne die zu exportierende Katalogseite in InDesign.
 * 2. Stelle sicher, dass die im Skript definierten Ebenennamen, PDF-Vorgaben
 * und der Basispfad für den Export korrekt konfiguriert sind.
 * 3. Führe das Skript über das InDesign-Skripten-Bedienfeld aus.
 *
 * Version: 1.0
 *
 ******************************************************************************/

#targetengine "session"

/**
 * @description Hauptobjekt, das den gesamten PDF-Exportprozess kapselt.
 * @type {Object}
 */
var PDFExportierer = {
    // Konfiguration für den Export
    konfiguration: {
        basisPfad: "t:/SERVUS-MULTI-DEPARTMENT/MKTG_Korrekturen/05_Katalog Seiten/",
        pdfVorgabe: "Intranets - 155 dpi",
        pdfVorgabe_Intranet: "Intranets - 155 dpi",
        ebenenSets: {
            AD: [
                "Preiskreis -30%",
                "CHFR Preiskreis -30%",
                "Text",
                "CHFR Text",
                "Bilder",
                "Standard"
            ],
            AD_Minimal: [
                "Preiskreis -30%",
                "Text",
                "Bilder",
                "Standard"
            ],
            CD: [
                "CH Preiskreis -30%",
                "CHFR Preiskreis -30%",
                "CH Text",
                "CHFR Text",
                "Bilder",
                "Standard"
            ],
            CHFR: [
                "CH Preiskreis -30%",
                "CHFR Preiskreis -30%",
                "CH Text",
                "CHFR Text",
                "Bilder",
                "Standard"
            ]
        }
    },

    // Speichert den Zustand des Skripts während der Ausführung
    zustand: {
        dokument: null,
        basisName: "",
        zielordner: null,
        exportierteDateienAnzahl: 0,
        freigabeVersionErstellt: false
    },

    /**
     * Hauptmethode, die den gesamten Prozess steuert: Dokumentenprüfung,
     * Ebenenanalyse, Benutzerbestätigung und Export der verschiedenen Versionen.
     */
    ausfuehren: function() {
        if (!this.pruefeDokument()) return;
        this.zustand.dokument = app.activeDocument;

        var vollstaendigeADEbenen = this.pruefeEbenenExistenz(this.konfiguration.ebenenSets.AD);
        var cdEbenen = this.pruefeEbenenExistenz(this.konfiguration.ebenenSets.CD);
        var minimaleADEbenen = this.pruefeEbenenExistenz(this.konfiguration.ebenenSets.AD_Minimal);

        var zuExportierendeVersionen = [];
        var warnmeldung = "";
        var erstelleFreigabe = false;

        if (vollstaendigeADEbenen && cdEbenen) {
            zuExportierendeVersionen = ["AD", "CD", "CHFR"];
            erstelleFreigabe = true;
        } else if (vollstaendigeADEbenen && !cdEbenen) {
            zuExportierendeVersionen = ["AD"];
            warnmeldung = "Warnung\nDie Ebenen für den CD/CHFR-Export wurden nicht gefunden. Es wird nur die AD-Version der PDF-Datei erstellt.";
        } else if (minimaleADEbenen) {
            zuExportierendeVersionen = ["AD_Minimal"];
            warnmeldung = "Warnung\nNur eine minimale Anzahl von Ebenen für die AD-Version wurde gefunden. Die AD-PDF wird mit diesen Ebenen erstellt und die CD/CHFR-Versionen werden übersprungen.";
        } else {
            alert("Fehler\nKeine der erforderlichen Ebenengruppen (weder vollständig noch minimal) für den AD-Export wurde in deinem Dokument gefunden. Das Skript wird beendet.");
            return;
        }

        if (!this.startBestaetigung()) return;
        
        if (warnmeldung) {
            alert(warnmeldung);
        }
        
        this.zustand.basisName = this.basisNamenExtrahieren(this.zustand.dokument.name);
        if (!this.zustand.basisName) return;

        this.zustand.zielordner = this.zielordnerFinden(this.zustand.basisName);
        if (!this.zustand.zielordner) return;

        if (!this.ueberschreibBestaetigung(zuExportierendeVersionen)) return;
        
        var erfolg = true;
        for (var i = 0; i < zuExportierendeVersionen.length; i++) {
            erfolg = erfolg && this.exportiereVersion(zuExportierendeVersionen[i]);
        }
        
        if (erstelleFreigabe && this.zustand.exportierteDateienAnzahl === 3) {
            this.erstelleFreigabeVersion();
        }
        
        this.stelleEndgueltigenEbenenZustandHer();

        if (this.zustand.exportierteDateienAnzahl > 0) {
            this.erfolgsMeldung();
        }
    },

    /**
     * Erstellt eine spezielle PDF-Version für das Intranet (nur Seite 1, minimale Ebenen)
     * im "Freigabe"-Unterordner.
     */
    erstelleFreigabeVersion: function() {
        try {
            var freigabeOrdner = new Folder(this.zustand.zielordner.fsName + "/Freigabe");
            if (!freigabeOrdner.exists) { freigabeOrdner.create(); }
            if (!freigabeOrdner.exists) {
                alert("Warnung\nDer 'Freigabe'-Unterordner konnte nicht erstellt werden. Die Intranet-Version wird übersprungen.");
                return;
            }

            var ebenenSet = this.konfiguration.ebenenSets.AD_Minimal;
            this.setzeEbenenSichtbarkeit(ebenenSet);

            var dateiNameOhneEndung = this.zustand.dokument.name.replace(/\.indd$/, '');
            var zielDatei = new File(freigabeOrdner.fsName + "/" + dateiNameOhneEndung + "_AD.pdf");

            var pdfVorgabe = app.pdfExportPresets.itemByName(this.konfiguration.pdfVorgabe_Intranet);
            if (!pdfVorgabe.isValid) {
                alert("Warnung\nDie PDF-Vorgabe '" + this.konfiguration.pdfVorgabe_Intranet + "' wurde nicht gefunden. Die Intranet-Version wird übersprungen.");
                return;
            }
            
            app.pdfExportPreferences.pageRange = "1";
            this.zustand.dokument.exportFile(ExportFormat.PDF_TYPE, zielDatei, false, pdfVorgabe);
            this.zustand.freigabeVersionErstellt = true;
            app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;

        } catch (e) {
            alert("Ein Fehler ist bei der Erstellung der Freigabe-Version aufgetreten:\n\n" + e.message);
            app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
        }
    },
    
    /**
     * Stellt am Ende des Skripts einen logischen, vordefinierten Ebenenzustand wieder her.
     */
    stelleEndgueltigenEbenenZustandHer: function() {
        var adEbenenVollstaendigExistieren = this.pruefeEbenenExistenz(this.konfiguration.ebenenSets.AD);
        if (adEbenenVollstaendigExistieren) {
            this.setzeEbenenSichtbarkeit(this.konfiguration.ebenenSets.AD);
        } else {
            this.setzeEbenenSichtbarkeit(this.konfiguration.ebenenSets.AD_Minimal);
        }
    },

    /**
     * Überprüft, ob alle Ebenen aus einer gegebenen Liste im Dokument existieren.
     * @param {string[]} ebenenNamen Ein Array von Ebenennamen, die überprüft werden sollen.
     * @returns {boolean} `true`, wenn alle Ebenen existieren, sonst `false`.
     */
    pruefeEbenenExistenz: function(ebenenNamen) {
        for (var i = 0; i < ebenenNamen.length; i++) {
            var ebene = this.zustand.dokument.layers.itemByName(ebenenNamen[i]);
            if (!ebene.isValid) { return false; }
        }
        return true;
    },

    /**
     * Stellt sicher, dass ein InDesign-Dokument geöffnet ist.
     * @returns {boolean} `true`, wenn ein Dokument geöffnet ist, sonst `false`.
     */
    pruefeDokument: function() {
        if (app.documents.length === 0) {
            alert("Fehler\nKein InDesign-Dokument geöffnet. Bitte öffne ein Dokument und starte das Skript erneut.");
            return false;
        }
        return true;
    },

    /**
     * Zeigt einen initialen Bestätigungsdialog für den Skriptstart an.
     * @returns {boolean} `true`, wenn der Benutzer bestätigt, sonst `false`.
     */
    startBestaetigung: function() {
        return confirm("Start des PDF-Exports\n\nDieses Skript erstellt PDFs mit spezifischen Ebeneneinstellungen.\n\nMöchtest du fortfahren?");
    },
    
    /**
     * Extrahiert den Basisnamen aus dem Dateinamen des Dokuments, der für den Zielordner verwendet wird.
     * @param {string} dateiName Der Name der InDesign-Datei.
     * @returns {string|null} Der extrahierte Basisname oder null bei einem Fehler.
     */
    basisNamenExtrahieren: function(dateiName) {
        var match = dateiName.match(/^[^_]+(?:_0,5|_0,25)?/);
        if (match && match[0]) {
            return match[0];
        }
        alert("Fehler\nAus dem Dateinamen '" + dateiName + "' konnte kein gültiger Ordnername extrahiert werden. Überprüfe das Namensformat.");
        return null;
    },

    /**
     * Findet den Zielordner basierend auf dem Basisnamen. Wenn der Ordner nicht existiert,
     * wird angeboten, ihn zu erstellen.
     * @param {string} basisName Der Basisname für die Ordnersuche.
     * @returns {Folder|null} Das Folder-Objekt des Zielordners oder null bei Abbruch/Fehler.
     */
    zielordnerFinden: function(basisName) {
        var hauptordner = new Folder(this.konfiguration.basisPfad);
        if (!hauptordner.exists) {
            alert("Fehler\nDer Basispfad '" + this.konfiguration.basisPfad + "' wurde nicht gefunden.");
            return null;
        }
        var unterordner = hauptordner.getFiles();
        for (var i = 0; i < unterordner.length; i++) {
            var aktuellerOrdner = unterordner[i];
            if (aktuellerOrdner instanceof Folder && aktuellerOrdner.name === basisName) {
                aktuellerOrdner.execute();
                return aktuellerOrdner;
            }
        }
        var neuerOrdner = new Folder(hauptordner.fsName + "/" + basisName);
        if (confirm("Information\n\n" + "Der Zielordner für die ID '" + basisName + "' wurde nicht gefunden.\n\n" + "Soll der folgende Ordner erstellt werden?\n" + neuerOrdner.fsName)) {
            if (neuerOrdner.create()) {
                neuerOrdner.execute();
                return neuerOrdner;
            }
            alert("Fehler\nDer Ordner '" + neuerOrdner.fsName + "' konnte nicht erstellt werden. Bitte überprüfe deine Berechtigungen.");
            return null;
        }
        alert("Vorgang abgebrochen.\nDer Ordner wurde nicht erstellt und der Export wurde beendet.");
        return null;
    },

    /**
     * Prüft, ob bereits PDF-Dateien im Zielordner existieren und fragt den Benutzer,
     * ob diese überschrieben werden sollen.
     * @param {string[]} versionen Die zu exportierenden Versionen (z.B. ["AD", "CD"]).
     * @returns {boolean} `true`, wenn der Export fortgesetzt werden soll, sonst `false`.
     */
    ueberschreibBestaetigung: function(versionen) {
        var existierendeDateien = [];
        var dateiNameOhneEndung = this.zustand.dokument.name.replace(/\.indd$/, '');
        for (var i = 0; i < versionen.length; i++) {
            var versionKey = versionen[i];
            var dateiEndung = versionKey.replace("_Minimal", "");
            var pdfDatei = new File(this.zustand.zielordner.fsName + "/" + dateiNameOhneEndung + "_" + dateiEndung + ".pdf");
            if (pdfDatei.exists) { existierendeDateien.push(pdfDatei.name); }
        }
        if (existierendeDateien.length > 0) {
            return confirm("Warnung: Dateien existieren bereits\n\n" + "Die folgenden Dateien sind im Zielordner bereits vorhanden:\n" + existierendeDateien.join("\n") + "\n\n" + "Möchtest du sie überschreiben?");
        }
        return true;
    },

    /**
     * Schaltet alle Ebenen unsichtbar und macht dann nur die in der Liste angegebenen Ebenen sichtbar.
     * @param {string[]} ebenenNamen Ein Array der sichtbar zu schaltenden Ebenennamen.
     */
    setzeEbenenSichtbarkeit: function(ebenenNamen) {
        var alleEbenen = this.zustand.dokument.layers;
        for (var i = 0; i < alleEbenen.length; i++) {
            alleEbenen[i].visible = false;
        }
        for (var j = 0; j < ebenenNamen.length; j++) {
            var ebene = alleEbenen.itemByName(ebenenNamen[j]);
            if (ebene.isValid) {
                ebene.visible = true;
            }
        }
    },

    /**
     * Exportiert eine einzelne PDF-Version basierend auf einem Versionsschlüssel.
     * @param {string} versionKey Der Schlüssel für die Version (z.B. "AD", "AD_Minimal").
     * @returns {boolean} `true` bei Erfolg, `false` bei einem Fehler.
     */
    exportiereVersion: function(versionKey) {
        try {
            app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
            var ebenenSet = this.konfiguration.ebenenSets[versionKey];
            this.setzeEbenenSichtbarkeit(ebenenSet);
            var dateiEndung = versionKey.replace("_Minimal", "");
            var dateiNameOhneEndung = this.zustand.dokument.name.replace(/\.indd$/, '');
            var zielDatei = new File(this.zustand.zielordner.fsName + "/" + dateiNameOhneEndung + "_" + dateiEndung + ".pdf");
            var pdfVorgabe = app.pdfExportPresets.itemByName(this.konfiguration.pdfVorgabe);
            if (!pdfVorgabe.isValid) {
                alert("Fehler\nDie PDF-Vorgabe '" + this.konfiguration.pdfVorgabe + "' wurde nicht gefunden.");
                return false;
            }
            this.zustand.dokument.exportFile(ExportFormat.PDF_TYPE, zielDatei, false, pdfVorgabe);
            this.zustand.exportierteDateienAnzahl++;
            return true;
        } catch (e) {
            alert("Ein Fehler ist beim Export der Version '" + versionKey + "' aufgetreten:\n\n" + e.message);
            return false;
        }
    },

    /**
     * Zeigt am Ende eine Erfolgsmeldung mit einer Zusammenfassung der exportierten Dateien an.
     */
    erfolgsMeldung: function() {
        var anzahl = this.zustand.exportierteDateienAnzahl;
        var dateiText = (anzahl === 1) ? "PDF-Datei wurde" : "PDF-Dateien wurden";
        var basisNachricht = "Insgesamt " + anzahl + " " + dateiText + " im folgenden Ordner gespeichert:\n" + this.zustand.zielordner.fsName;

        if (this.zustand.freigabeVersionErstellt) {
            basisNachricht += "\n\nZusätzlich wurde eine Intranet-Version (nur Seite 1) im 'Freigabe'-Unterordner erstellt.";
        }
        
        alert("Export erfolgreich abgeschlossen!\n\n" + basisNachricht);
    }
};

// Startet den gesamten Prozess
PDFExportierer.ausfuehren();