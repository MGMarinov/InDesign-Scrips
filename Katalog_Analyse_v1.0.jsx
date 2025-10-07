// Katalog_Analyse.jsx
// Version: 1.3.0
// Beschreibung: Analysiert alle verknüpften PDF- und INDD-Dateien. Verwendet eine robuste Methode,
// um die Seitenzahl der Quelle zu lesen und zeigt klar an, wenn PDF-interne Daten nicht analysierbar sind.

(function() {

    /**
     * Hauptklasse zur Analyse der Verknüpfungen eines InDesign-Katalogs.
     * Kapselt die gesamte Logik für die Verarbeitung und Berichterstellung.
     */
    function KatalogAnalysator() {
        // --- Konstanten und Eigenschaften ---
        this.SKRIPT_NAME = "Katalog Analysator";
        this.SKRIPT_VERSION = "6.1.0";
        this.aktivesDokument = null;
        this.csvDatei = null;
        this.reportDaten = [];

        /**
         * Führt das Skript aus. Dies ist die primäre Steuerungsmethode.
         */
        this.ausfuehren = function() {
            if (!this._vorabpruefungen()) {
                return;
            }
            this._setzeSpeicherort();

            var zuVerarbeitendeLinks = this._filtereRelevanteLinks();
            if (zuVerarbeitendeLinks.length === 0) {
                alert("Im Dokument wurden keine verknüpften PDF- oder INDD-Dateien gefunden.");
                return;
            }

            this._linksAnalysieren(zuVerarbeitendeLinks);
            this._schreibeDatei();

            alert("Die Analyse wurde abgeschlossen.\nDer Report wurde hier gespeichert:\n" + this.csvDatei.fsName);
        };

        /**
         * Prüft, ob ein Dokument geöffnet und gespeichert ist.
         * @returns {boolean} True, wenn die Prüfungen erfolgreich sind, sonst false.
         * @private
         */
        this._vorabpruefungen = function() {
            if (app.documents.length === 0) {
                alert("Bitte öffne ein Dokument, bevor Du dieses Skript ausführst.");
                return false;
            }
            this.aktivesDokument = app.activeDocument;
            if (!this.aktivesDokument.saved || !this.aktivesDokument.filePath) {
                alert("Dein Dokument muss gespeichert sein, damit der Report automatisch am richtigen Ort abgelegt werden kann.");
                return false;
            }
            return true;
        };
        
        /**
         * Definiert den automatischen Speicherort und Dateinamen für den Report.
         * @private
         */
        this._setzeSpeicherort = function() {
            var basisName = this.aktivesDokument.name.replace(/\.indd$/i, '');
            var zeitstempel = this._erstelleZeitstempel();
            var pfad = this.aktivesDokument.filePath;
            this.csvDatei = new File(pfad + "/" + basisName + "_Analyse-Report_" + zeitstempel + ".csv");
        };

        /**
         * Filtert die Links des Dokuments und gibt nur PDF- und INDD-Dateien zurück.
         * @returns {Array<Link>} Ein Array mit den zu verarbeitenden Links.
         * @private
         */
        this._filtereRelevanteLinks = function() {
            var alleLinks = this.aktivesDokument.links;
            var relevanteLinks = [];
            for (var i = 0; i < alleLinks.length; i++) {
                var linkObj = alleLinks[i];
                var dateiTyp = this._getDateiTyp(linkObj.name).toLowerCase();
                if (dateiTyp === 'indd' || dateiTyp === 'pdf') {
                    relevanteLinks.push(linkObj);
                }
            }
            return relevanteLinks;
        };

        /**
         * Iteriert durch die gefilterten Verknüpfungen und sammelt die Daten.
         * @param {Array<Link>} links - Die zu analysierenden Verknüpfungen.
         * @private
         */
        this._linksAnalysieren = function(links) {
            this._erstelleCsvKopfzeile();
            
            var fortschritt = new Window("palette", "Analyse läuft... (" + this.SKRIPT_VERSION + ")", undefined, { closeButton: false });
            fortschritt.add("statictext", undefined, "PDF- & INDD-Dateien werden analysiert...");
            var balken = fortschritt.add("progressbar", undefined, 0, links.length);
            balken.preferredSize.width = 400;
            fortschritt.show();

            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

            for (var i = 0; i < links.length; i++) {
                balken.value = i + 1;
                var verknuepfung = links[i];
                var dateiTyp = this._getDateiTyp(verknuepfung.name);
                var linkDatei = new File(verknuepfung.filePath);
                
                var inddDetails = {};
                if (dateiTyp.toLowerCase() === 'indd' && linkDatei.exists) {
                    inddDetails = this._inddDetailsAnalysieren(linkDatei);
                }

                var zeilendaten = this._erstelleZeilendaten(verknuepfung, dateiTyp, linkDatei, inddDetails);
                this.reportDaten.push(zeilendaten.join(";"));
            }
            
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
            fortschritt.close();
        };

        /**
         * Öffnet eine verknüpfte InDesign-Datei, um deren Eigenschaften zu analysieren.
         * @param {File} linkDatei - Die zu analysierende InDesign-Datei.
         * @returns {object} Ein Objekt mit den Details der InDesign-Datei.
         * @private
         */
        this._inddDetailsAnalysieren = function(linkDatei) {
            var details = {
                seitenanzahl: "N/A", breite: "N/A", hoehe: "N/A",
                anschnittOben: "N/A", anschnittUnten: "N/A", anschnittInnen: "N/A", anschnittAussen: "N/A",
                version: "N/A"
            };
            
            var quelldokument = null;
            try {
                quelldokument = app.open(linkDatei, false);
                
                var prefs = quelldokument.documentPreferences;
                details.seitenanzahl = quelldokument.pages.length;
                details.breite = prefs.pageWidth;
                details.hoehe = prefs.pageHeight;
                details.anschnittOben = prefs.documentBleedTopOffset;
                details.anschnittUnten = prefs.documentBleedBottomOffset;
                details.anschnittInnen = prefs.documentBleedInsideOrLeftOffset;
                details.anschnittAussen = prefs.documentBleedOutsideOrRightOffset;
                details.version = quelldokument.metadataPreferences.getProperty("http://ns.adobe.com/xap/1.0/", "CreatorTool") || "N/A";

            } catch (e) {
                // Bei einem Fehler bleiben die Werte "N/A"
            } finally {
                if (quelldokument) {
                    quelldokument.close(SaveOptions.NO);
                }
            }
            return details;
        };

        /**
         * Erstellt die Kopfzeile für die CSV-Datei.
         * @private
         */
        this._erstelleCsvKopfzeile = function() {
            var kopfzeile = [
                "Katalogseite", "Quelldokumentseite", "Dateiname", "Typ", "Status",
                "Erstellungsdatum", "InDesign-Version",
                "Grafik-Breite (mm)", "Grafik-Höhe (mm)", "Import-Begrenzungsrahmen",
                "Quell-Seitenanzahl", "Quell-Breite (mm)", "Quell-Höhe (mm)",
                "Quell-Anschnitt Oben (mm)", "Quell-Anschnitt Unten (mm)", "Quell-Anschnitt Innen (mm)", "Quell-Anschnitt Aussen (mm)",
                "Dateipfad"
            ];
            this.reportDaten.push(kopfzeile.join(";"));
        };
        
        /**
         * Stellt die Daten für eine einzelne Zeile der CSV-Datei zusammen.
         * @returns {Array} Ein Array von Zeichenketten für die CSV-Zeile.
         * @private
         */
        this._erstelleZeilendaten = function(verknuepfung, dateiTyp, linkDatei, inddDetails) {
            var grafik = verknuepfung.parent;
            var grafikBounds = grafik.geometricBounds;
            var grafikBreite = grafikBounds[3] - grafikBounds[1];
            var grafikHoehe = grafikBounds[2] - grafikBounds[0];

            var begrenzungsrahmen = "N/A";
            try {
                if(grafik.hasOwnProperty("importedPageAttributes")) {
                    begrenzungsrahmen = this._getBoundingBoxName(grafik.importedPageAttributes.pageBoundingBox);
                }
            } catch(e) {}
            
            var katalogSeite = grafik.parentPage ? grafik.parentPage.name : "Montagefläche";
            var quellSeite = this._getQuellseitenInfo(verknuepfung);
            var erstellungsdatum = linkDatei.exists ? linkDatei.created.toLocaleString() : "N/A";

            // Platzhalter für nicht analysierbare PDF-Eigenschaften definieren
            var pdfPlatzhalter = "PDF - nicht analysierbar";
            var istPDF = (dateiTyp.toLowerCase() === 'pdf');

            var quellVersion     = istPDF ? pdfPlatzhalter : (inddDetails.version || "N/A");
            var quellSeitenanzahl= istPDF ? pdfPlatzhalter : (inddDetails.seitenanzahl || "N/A");
            var quellBreite      = istPDF ? pdfPlatzhalter : this._formatiereEinheit(inddDetails.breite);
            var quellHoehe       = istPDF ? pdfPlatzhalter : this._formatiereEinheit(inddDetails.hoehe);
            var quellAnschnittO  = istPDF ? pdfPlatzhalter : this._formatiereEinheit(inddDetails.anschnittOben);
            var quellAnschnittU  = istPDF ? pdfPlatzhalter : this._formatiereEinheit(inddDetails.anschnittUnten);
            var quellAnschnittI  = istPDF ? pdfPlatzhalter : this._formatiereEinheit(inddDetails.anschnittInnen);
            var quellAnschnittA  = istPDF ? pdfPlatzhalter : this._formatiereEinheit(inddDetails.anschnittAussen);
            
            return [
                '"' + katalogSeite + '"',
                '"' + quellSeite + '"',
                '"' + verknuepfung.name + '"',
                '"' + dateiTyp + '"',
                '"' + this._getLinkStatus(verknuepfung.status) + '"',
                '"' + erstellungsdatum + '"',
                '"' + quellVersion + '"',
                '"' + this._formatiereEinheit(grafikBreite) + '"',
                '"' + this._formatiereEinheit(grafikHoehe) + '"',
                '"' + begrenzungsrahmen + '"',
                '"' + quellSeitenanzahl + '"',
                '"' + quellBreite + '"',
                '"' + quellHoehe + '"',
                '"' + quellAnschnittO + '"',
                '"' + quellAnschnittU + '"',
                '"' + quellAnschnittI + '"',
                '"' + quellAnschnittA + '"',
                '"' + verknuepfung.filePath + '"'
            ];
        };

        /**
         * Schreibt den gesammelten Inhalt in die CSV-Datei.
         * @private
         */
        this._schreibeDatei = function() {
            this.csvDatei.encoding = "UTF-8";
            this.csvDatei.open("w");
            this.csvDatei.write("\uFEFF"); // BOM
            this.csvDatei.write(this.reportDaten.join("\n"));
            this.csvDatei.close();
        };
        
        // --- Hilfsmethoden ---

        this._getQuellseitenInfo = function(verknuepfung) {
            try {
                var uebergeordnetesObjekt = verknuepfung.parent;

                if (uebergeordnetesObjekt.constructor.name === "PDF") {
                    if (uebergeordnetesObjekt.pdfAttributes.hasOwnProperty("pageNumber")) {
                        return uebergeordnetesObjekt.pdfAttributes.pageNumber;
                    }
                }
                if (uebergeordnetesObjekt.hasOwnProperty("importedPageAttributes")) {
                     if (uebergeordnetesObjekt.importedPageAttributes.hasOwnProperty("pageNumber")) {
                        return uebergeordnetesObjekt.importedPageAttributes.pageNumber;
                    }
                }
                if (verknuepfung.hasOwnProperty("linkXRef") && verknuepfung.linkXRef) {
                    return verknuepfung.linkXRef.sourcePageNumber + " (XRef)";
                }
                return "Unbekannt";
            } catch (e) {
                return "Fehler";
            }
        };

        this._erstelleZeitstempel = function() {
            var jetzt = new Date();
            var jahr = jetzt.getFullYear();
            var monat = ("0" + (jetzt.getMonth() + 1)).slice(-2);
            var tag = ("0" + jetzt.getDate()).slice(-2);
            var stunde = ("0" + jetzt.getHours()).slice(-2);
            var minute = ("0" + jetzt.getMinutes()).slice(-2);
            var sekunde = ("0" + jetzt.getSeconds()).slice(-2);
            return jahr + monat + tag + "-" + stunde + minute + sekunde;
        };

        this._getLinkStatus = function(status) {
            switch (status) {
                case LinkStatus.NORMAL: return "OK";
                case LinkStatus.LINK_MISSING: return "Fehlend";
                case LinkStatus.LINK_OUT_OF_DATE: return "Modifiziert";
                default: return "Unbekannt";
            }
        };
        
        this._getBoundingBoxName = function(boxEnum) {
            switch(boxEnum) {
                case BoundingBoxOptions.ART_BOX: return "ART_BOX";
                case BoundingBoxOptions.BLEED_BOX: return "BLEED_BOX";
                case BoundingBoxOptions.CROP_BOX: return "CROP_BOX";
                case BoundingBoxOptions.MEDIA_BOX: return "MEDIA_BOX";
                case BoundingBoxOptions.PAGE_BOUNDING_BOX: return "PAGE_BOUNDING_BOX";
                case BoundingBoxOptions.TRIM_BOX: return "TRIM_BOX";
                default: return "Unbekannt";
            }
        };

        this._getDateiTyp = function(dateiname) {
            var erweiterung = dateiname.split('.').pop();
            return erweiterung ? erweiterung.toUpperCase() : 'UNBEKANNT';
        };
        
        this._formatiereEinheit = function(wert) {
            if (typeof wert === 'number') {
                return wert.toString().replace('.', ',');
            }
            return wert;
        };
    }

    // --- Skriptausführung ---
    var analysator = new KatalogAnalysator();
    analysator.ausfuehren();

})();