// INDD_PDF_Austauscher_und_Validierer.jsx

(function() {
    /**
     * Hauptklasse für den Austausch und die Validierung von Verknüpfungen.
     * Kapselt die gesamte Logik für die Verarbeitung der Verknüpfungen in einem InDesign-Dokument.
     */
    function LinkManager() {
        // --- GRUNDEINSTELLUNGEN ---
        this.SKRIPT_NAME = "INDD <-> PDF Austauscher & Validierer";
        this.SKRIPT_VERSION = "v1.5.1"; // Versionsnummer wegen Optimierung erhöht
        this.DATUM_TOLERANZ_MINUTEN = 20;
        this.LOGGING_AKTIVIERT = false;
        this.FALLBACK_SUCHPFAD = "x:/_Grafik 2024/10_Katalogseiten/"; // Alternativer Suchpfad für fehlende Dateien
        this.FALLBACK_SUCHTIEFE = 3; // Rekursive Suchtiefe für den alternativen Pfad

        // --- Interne Zustandsvariablen ---
        this.aktivesDokument = null;
        this.basisDokPfad = null;
        this.basisDokNameOhneErw = null;
        this.logDatei = null;
        this.operationCancelled = false; // Flag für den Abbruch durch den Benutzer

        /**
         * Führt Vorabprüfungen durch, um sicherzustellen, dass das Skript laufen kann.
         * @returns {boolean} True, wenn alle Prüfungen erfolgreich sind, sonst false.
         */
        this.fuehreVorabpruefungenDurch = function() {
            if (app.documents.length === 0) {
                alert("Bitte öffne ein Dokument, bevor Du dieses Skript ausführst.", this.SKRIPT_NAME + " - Kein Dokument geöffnet");
                return false;
            }
            this.aktivesDokument = app.activeDocument;

            if (!this.aktivesDokument.saved || !this.aktivesDokument.filePath) {
                alert("Dein aktives Dokument muss zuerst gespeichert werden, damit der Standardspeicherort für Exporte und Logs bestimmt werden kann.\nBitte speichere Dein Dokument und führe das Skript erneut aus.", this.SKRIPT_NAME + " - Dokument nicht gespeichert");
                return false;
            }
            this.basisDokPfad = this.aktivesDokument.filePath;
            this.basisDokNameOhneErw = this.aktivesDokument.name.replace(/\.indd$/i, "");
            return true;
        };

        /**
         * Initialisiert den Logging-Mechanismus für das Skript, falls aktiviert.
         */
        this.initialisiereLogger = function() {
            if (!this.LOGGING_AKTIVIERT) return;
            var logOrdner = Folder(this.basisDokPfad);
            this.logDatei = new File(logOrdner.fsName + "/" + this.basisDokNameOhneErw + "_verarbeitungs_log.txt");
            this.logDatei.encoding = "UTF-8";
        };

        /**
         * Schreibt eine Nachricht in die Log-Datei, falls aktiviert.
         * @param {string} nachricht Die zu protokollierende Nachricht.
         */
        this.log = function(nachricht) {
            if (!this.LOGGING_AKTIVIERT || !this.logDatei) return;
            try {
                this.logDatei.open("a");
                var zeitstempel = new Date().toTimeString().substr(0, 8);
                this.logDatei.writeln(zeitstempel + " - " + nachricht);
                this.logDatei.close();
            } catch (e) {
                // Fehler beim Schreiben des Logs ignorieren
            }
        };

        /**
         * Formatiert ein Fehler-Objekt in eine lesbare Zeichenkette.
         * @param {Error} errorObj Das Fehler-Objekt.
         * @param {string} dateiInfo Kontextinformation zur Datei, bei der der Fehler auftrat.
         * @returns {string} Eine formatierte Fehlermeldung.
         */
        this.formatiereFehlermeldung = function(errorObj, dateiInfo) {
            var basisNachricht = "Unbekannter Fehler.";
            if (errorObj.message && errorObj.message.length > 0) {
                basisNachricht = errorObj.message;
            } else {
                basisNachricht = "Ein interner Anwendungsfehler ist aufgetreten (keine Detailbeschreibung). " +
                    "Dies kann passieren, wenn eine Datei (z.B. '" + (dateiInfo || 'unbekannt') + "') " +
                    "beim Öffnen/Platzieren eine Dialogbox (z.B. 'Fehlende Schriftarten') auslösen würde.";
            }
            return basisNachricht + " (Zeile: " + errorObj.line + ", Fehlercode: " + errorObj.number + ")";
        };

        /**
         * Sucht rekursiv nach einer Datei in einem Ordner.
         * @param {Folder} ordner Der Startordner für die Suche.
         * @param {string} dateiName Der Name der zu suchenden Datei.
         * @param {Array} alleTreffer Ein Array, in dem die gefundenen Dateien gesammelt werden.
         * @param {number} maxTiefe Die maximale Rekursionstiefe.
         * @param {number} aktuelleTiefe Die aktuelle Tiefe der Rekursion.
         */
        this.sucheRekursiv = function(ordner, dateiName, alleTreffer, maxTiefe, aktuelleTiefe) {
            if (aktuelleTiefe >= maxTiefe) {
                return;
            }
            var dateien = ordner.getFiles();
            for (var i = 0; i < dateien.length; i++) {
                var datei = dateien[i];
                if (datei instanceof File && datei.name.toLowerCase() === dateiName.toLowerCase()) {
                    alleTreffer.push(datei);
                } else if (datei instanceof Folder) {
                    this.sucheRekursiv(datei, dateiName, alleTreffer, maxTiefe, aktuelleTiefe + 1);
                }
            }
        };

        /**
         * Findet eine Datei, auch an einem alternativen Speicherort (Fallback).
         * @param {string} originalPfad Der ursprüngliche, erwartete Pfad der Datei.
         * @returns {File|null} Das gefundene File-Objekt oder null.
         */
        this.findeDateiMitFallback = function(originalPfad) {
            var datei = File(originalPfad);
            if (datei.exists) {
                return datei;
            }

            var fallbackRoot = Folder(this.FALLBACK_SUCHPFAD);
            if (!fallbackRoot.exists) {
                return null;
            }

            var dateiName = datei.name;
            var alleTreffer = [];
            this.sucheRekursiv(fallbackRoot, dateiName, alleTreffer, this.FALLBACK_SUCHTIEFE, 0);

            if (alleTreffer.length === 0) return null;
            if (alleTreffer.length === 1) return alleTreffer[0];

            alleTreffer.sort(function(a, b) {
                return b.modified.getTime() - a.modified.getTime();
            });
            return alleTreffer[0];
        };

        /**
         * Stellt sicher, dass die TrimBox für importierte Objekte verwendet wird.
         * @param {Link} link Der zu konfigurierende Link.
         */
        this.stelleTrimBoxEin = function(link) {
            if (link.parent && (link.parent.constructor.name === "Image" ||
                               link.parent.constructor.name === "PDF" ||
                               link.parent.constructor.name === "ImportedPage")) {
                try {
                    link.parent.importedPageAttributes.pageBoundingBox = BoundingBoxOptions.TRIM_BOX;
                } catch (e) {
                    // Fehler ignorieren, wenn die Eigenschaft nicht existiert
                }
            }
        };

        /**
         * Zeigt den finalen Ergebnisdialog für Austauschaktionen an.
         * @param {string} aktionTitel Titel der durchgeführten Aktion (z.B. "INDD -> PDF").
         * @param {Array} fehlerListe Eine Liste von Objekten mit Fehlerdetails.
         * @param {number} anzahlErsetzungen Die Anzahl der erfolgreichen Ersetzungen.
         */
        this.zeigeErgebnisDialog = function(aktionTitel, fehlerListe, anzahlErsetzungen) {
            if (fehlerListe.length > 0) {
                var fehlerText = "Bei den folgenden Dateien ist ein Problem aufgetreten:\n\n";
                for (var k = 0; k < fehlerListe.length; k++) {
                    fehlerText += "- " + fehlerListe[k].betroffeneDatei + "\n  Grund: " + fehlerListe[k].grund + "\n";
                }
                alert("Die Aktion '" + aktionTitel + "' wurde mit " + anzahlErsetzungen + " erfolgreichen Ersetzungen und " + fehlerListe.length + " Fehlern abgeschlossen.\n\n" + fehlerText, this.SKRIPT_NAME + " - Ergebnis");
            } else if (anzahlErsetzungen > 0) {
                alert(anzahlErsetzungen + " Verknüpfungen wurden erfolgreich ausgetauscht (" + aktionTitel + ").", this.SKRIPT_NAME + " - Erfolg");
            } else {
                alert("Es wurden keine passenden Verknüpfungen für die Aktion '" + aktionTitel + "' gefunden oder es wurden keine Änderungen vorgenommen.", this.SKRIPT_NAME + " - Information");
            }
        };

        /**
         * Erstellt und verwaltet eine Fortschrittsanzeige.
         * @param {string} message Die initiale Nachricht, die im Fenster angezeigt wird.
         */
        this.erstelleFortschrittsanzeige = function(message) {
            var self = this;
            var w = new Window("palette", "Fortschritt", undefined, { closeButton: false });
            w.alignChildren = "fill";

            var t = w.add("statictext", undefined, message);
            t.preferredSize = [450, -1];

            var b = w.add("progressbar", undefined, 0, 100);

            var timeGroup = w.add('group');
            timeGroup.orientation = 'row';
            timeGroup.alignment = 'fill';

            var elapsedLabel = timeGroup.add("statictext", undefined, "Verstrichene Zeit: 00:00:00");
            elapsedLabel.alignment = "left";

            var estimateLabel = timeGroup.add("statictext", undefined, "Verbleibende Zeit: Berechnung...");
            estimateLabel.alignment = "right";

            var cancelButton = w.add("button", undefined, "Abbrechen");
            cancelButton.alignment = "center";

            var startTime = new Date();

            cancelButton.onClick = function() {
                self.operationCancelled = true;
                w.close();
            };

            function formatTime(seconds) {
                var h = Math.floor(seconds / 3600);
                var m = Math.floor((seconds % 3600) / 60);
                var s = Math.floor(seconds % 60);
                h = (h < 10 ? "0" : "") + h;
                m = (m < 10 ? "0" : "") + m;
                s = (s < 10 ? "0" : "") + s;
                return h + ":" + m + ":" + s;
            }

            function updateTiming() {
                var currentTime = new Date();
                var elapsedTime = (currentTime - startTime) / 1000;
                elapsedLabel.text = "Verstrichene Zeit: " + formatTime(elapsedTime);

                var progress = b.value - b.minvalue;
                if (progress > 0) {
                    var totalSteps = b.maxvalue - b.minvalue;
                    var estimatedTotalTime = elapsedTime / progress * totalSteps;
                    var remainingTime = estimatedTotalTime - elapsedTime;
                    estimateLabel.text = "Verbleibende Zeit: " + formatTime(remainingTime);
                } else {
                    estimateLabel.text = "Verbleibende Zeit: Berechnung...";
                }
                w.layout.layout(true);
            }

            this.erstelleFortschrittsanzeige.close = function() { w.close(); };
            this.erstelleFortschrittsanzeige.increment = function(value) { b.value += value; updateTiming(); };
            this.erstelleFortschrittsanzeige.message = function(msg) { t.text = msg; w.update(); };
            this.erstelleFortschrittsanzeige.set = function(steps) { b.maxvalue = steps; updateTiming(); };

            w.show();
        };


        /**
         * Führt die Hauptlogik aus: Ersetzt Verknüpfungen basierend auf den angegebenen Dateierweiterungen.
         * @param {string} quellErweiterung Die Erweiterung der zu suchenden Links (z.B. ".pdf").
         * @param {string} zielErweiterung Die Erweiterung der Zieldateien (z.B. ".indd").
         * @param {string} aktionTitel Ein kurzer Titel für die Aktion (z.B. "PDF -> INDD").
         */
        this.fuehreRelinkDurch = function(quellErweiterung, zielErweiterung, aktionTitel) {
            this.initialisiereLogger();
            this.log("Skript gestartet: " + this.SKRIPT_NAME + " " + this.SKRIPT_VERSION);
            this.log("Aktion gestartet: " + aktionTitel);

            var alleLinks = this.aktivesDokument.links;
            var gefundeneLinks = [];
            for (var i = 0; i < alleLinks.length; i++) {
                if (alleLinks[i].name.toLowerCase().indexOf(quellErweiterung) > -1) {
                    gefundeneLinks.push(alleLinks[i]);
                }
            }

            if (gefundeneLinks.length === 0) {
                alert("Keine " + quellErweiterung.toUpperCase() + "-Verknüpfungen im Dokument gefunden.", this.SKRIPT_NAME);
                this.log("Keine " + quellErweiterung + "-Verknüpfungen gefunden.");
                return;
            }

            this.log(gefundeneLinks.length + " " + quellErweiterung + "-Verknüpfungen zur Verarbeitung gefunden.");
            
            this.operationCancelled = false;
            this.erstelleFortschrittsanzeige("Verknüpfungen werden verarbeitet...");
            this.erstelleFortschrittsanzeige.set(gefundeneLinks.length);

            var fehlerListe = [];
            var ersetzungenVorgenommen = 0;
            var quellRegex = new RegExp(quellErweiterung.replace('.', '\\.') + '$', 'i');
            
            var alteUILevel = app.scriptPreferences.userInteractionLevel;
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

            for (var j = 0; j < gefundeneLinks.length; j++) {
                if(this.operationCancelled) {
                    break;
                }
                var aktuellerLink = gefundeneLinks[j];
                this.erstelleFortschrittsanzeige.message("Verarbeite " + (j + 1) + "/" + gefundeneLinks.length + ": " + aktuellerLink.name);
                
                try {
                    var basisName = aktuellerLink.name.replace(quellRegex, "");
                    var zielName = basisName + zielErweiterung;
                    var verknuepfteDateiOrdner = File(aktuellerLink.filePath).path;
                    var zielDateiPfadOriginal = verknuepfteDateiOrdner + "/" + zielName;
                    
                    var zielDatei = this.findeDateiMitFallback(zielDateiPfadOriginal);

                    if (zielDatei && zielDatei.exists) {
                        this.log("Versuche Ersetzung: '" + aktuellerLink.name + "' -> '" + zielDatei.name + "'");
                        aktuellerLink.relink(zielDatei);
                        
                        this.stelleTrimBoxEin(aktuellerLink);
                        
                        ersetzungenVorgenommen++;
                        this.log("ERFOLG: Ersetzung durchgeführt.");
                    } else {
                        throw new Error("Die passende " + zielErweiterung + "-Datei ('" + zielName + "') wurde nicht gefunden.");
                    }

                } catch (err) {
                    var fehlerMsg = this.formatiereFehlermeldung(err, aktuellerLink.name);
                    this.log("FEHLER bei der Verarbeitung von '" + aktuellerLink.name + "': " + fehlerMsg);
                    fehlerListe.push({
                        betroffeneDatei: aktuellerLink.name,
                        grund: fehlerMsg
                    });
                }
                this.erstelleFortschrittsanzeige.increment(1);
            }
            
            this.erstelleFortschrittsanzeige.close();
            app.scriptPreferences.userInteractionLevel = alteUILevel;

            if (this.operationCancelled) {
                alert("Der Vorgang wurde vom Benutzer abgebrochen.", this.SKRIPT_NAME + " - Abgebrochen");
                this.log("Vorgang wurde vom Benutzer nach " + ersetzungenVorgenommen + " Ersetzungen abgebrochen.");
            } else {
                this.log("Der Relink-Vorgang (" + aktionTitel + ") wurde abgeschlossen. " + ersetzungenVorgenommen + " Ersetzungen vorgenommen.");
                this.zeigeErgebnisDialog(aktionTitel, fehlerListe, ersetzungenVorgenommen);
            }
        };
        
        /**
         * Führt die Validierung der Verknüpfungen durch und erstellt einen CSV-Report.
         */
        this.fuehreValidierungDurch = function() {
            this.initialisiereLogger();
            this.log("Skript gestartet: " + this.SKRIPT_NAME + " " + this.SKRIPT_VERSION);
            this.log("Aktion gestartet: Katalog validieren");
            alert("Die Validierung wird gestartet. Dieser Vorgang kann einige Zeit dauern...", this.SKRIPT_NAME);

            var alleLinks = this.aktivesDokument.links;
            var reportDaten = [];
            var toleranzInMs = this.DATUM_TOLERANZ_MINUTEN * 60 * 1000;

            for (var i = 0; i < alleLinks.length; i++) {
                var link = alleLinks[i];
                var linkDatei = File(link.filePath);
                var bemerkung = "";
                var seitenInfo = "Montagefläche";

                if (link.parent.parentPage) {
                    if (isNaN(parseInt(link.parent.parentPage.name, 10))) {
                        continue;
                    }
                    seitenInfo = link.parent.parentPage.name;
                }
                
                var statusInfo = "OK";
                switch (link.status) {
                    case LinkStatus.LINK_MISSING:
                        statusInfo = "Fehlend";
                        break;
                    case LinkStatus.LINK_OUT_OF_DATE:
                        statusInfo = "Modifiziert";
                        break;
                }
                
                var dateiTyp = "unbekannt";
                var nameTeile = link.name.split('.');
                if (nameTeile.length > 1) {
                    dateiTyp = nameTeile.pop().toLowerCase();
                }

                var linkNameLower = link.name.toLowerCase();
                if (linkDatei.exists && (linkNameLower.indexOf(".indd") > -1 || linkNameLower.indexOf(".pdf") > -1)) {
                    var istIndd = linkNameLower.indexOf(".indd") > -1;
                    var basisName = link.name.replace(istIndd ? /\.indd$/i : /\.pdf$/i, "");
                    var paarName = basisName + (istIndd ? ".pdf" : ".indd");
                    var paarDatei = File(linkDatei.path + "/" + paarName);

                    if (paarDatei.exists) {
                        var linkDatum = linkDatei.modified;
                        var paarDatum = paarDatei.modified;

                        if (istIndd) {
                            if (paarDatum.getTime() < linkDatum.getTime() - toleranzInMs) {
                                bemerkung = "Das zugehörige PDF ist möglicherweise veraltet.";
                            }
                        } else {
                            if (linkDatum.getTime() < paarDatum.getTime() - toleranzInMs) {
                                bemerkung = "PDF ist möglicherweise veraltet (INDD ist neuer).";
                            }
                        }
                    } else {
                        bemerkung = "Keine " + (istIndd ? "PDF" : "INDD") + "-Paardatei gefunden.";
                    }
                }
                
                var statusDetails = statusInfo;
                if (bemerkung !== "") {
                    statusDetails = statusInfo + " (" + bemerkung + ")";
                }

                reportDaten.push({
                    seite: seitenInfo,
                    name: link.name,
                    typ: dateiTyp,
                    status: statusDetails,
                    pfad: linkDatei.fsName
                });
            }

            if (reportDaten.length === 0) {
                alert("Keine auswertbaren Verknüpfungen auf den Dokumentseiten gefunden. Der Report wird nicht erstellt.", this.SKRIPT_NAME + " - Information");
                this.log("Keine relevanten Verknüpfungen für den Report gefunden.");
                return;
            }

            var csvInhalt = "Seite im Katalog;Link-Name;Link-Typ;Status/Details;Dateipfad\n";
            for (var j = 0; j < reportDaten.length; j++) {
                var zeile = reportDaten[j];
                csvInhalt += '"' + zeile.seite + '";' +
                             '"' + zeile.name + '";' +
                             '"' + zeile.typ + '";' +
                             '"' + zeile.status + '";' +
                             '"' + zeile.pfad + '"\n';
            }
            
            var zeitstempel = new Date();
            var datumsString = zeitstempel.getFullYear() +
                ("0" + (zeitstempel.getMonth() + 1)).slice(-2) +
                ("0" + zeitstempel.getDate()).slice(-2) + "-" +
                ("0" + zeitstempel.getHours()).slice(-2) +
                ("0" + zeitstempel.getMinutes()).slice(-2) +
                ("0" + zeitstempel.getSeconds()).slice(-2);

            var reportDateiName = this.basisDokNameOhneErw + "_validierungs_report_" + datumsString + ".csv";
            var reportDatei = new File(this.basisDokPfad + "/" + reportDateiName);
            reportDatei.encoding = "UTF-8";
            reportDatei.open("w");
            reportDatei.write(csvInhalt);
            reportDatei.close();
            
            this.log("Validierungsreport erfolgreich erstellt: " + reportDatei.fsName);
            alert("Validierung abgeschlossen!\nDer Report wurde erfolgreich unter folgendem Pfad gespeichert:\n\n" + reportDatei.fsName, this.SKRIPT_NAME + " - Erfolg");
        };

        /**
         * Zeigt den Hauptdialog zur Aktionsauswahl an.
         */
        this.zeigeHauptdialog = function() {
            var dialog = new Window("dialog", this.SKRIPT_NAME + " " + this.SKRIPT_VERSION);
            dialog.orientation = "column";
            dialog.alignChildren = ["fill", "top"];
            
            var infoText = "Dieses Skript hilft Dir bei folgenden Aktionen:\n" +
                           "- INDD-Verknüpfungen durch PDFs ersetzen.\n" +
                           "- PDF-Verknüpfungen durch INDDs ersetzen.\n" +
                           "- Verknüpfungen im Katalog validieren und einen Report erstellen.\n\n" +
                           "Bitte wähle eine Aktion:";
            dialog.add("statictext", undefined, infoText, { multiline: true });

            var buttonGroup = dialog.add("group");
            buttonGroup.orientation = "row";
            buttonGroup.alignChildren = ["fill", "center"];

            var inddZuPdfBtn = buttonGroup.add("button", undefined, "INDD -> PDF Katalog erstellen");
            var pdfZuInddBtn = buttonGroup.add("button", undefined, "PDF -> INDD Katalog erstellen");
            var validierenBtn = buttonGroup.add("button", undefined, "Katalog validieren");
            
            var abbrechenGroup = dialog.add("group");
            abbrechenGroup.orientation = "row";
            abbrechenGroup.alignChildren = ["center", "center"];
            var abbrechenBtn = abbrechenGroup.add("button", undefined, "Abbrechen", { name: "cancel" });
            abbrechenBtn.alignment = ["center", "bottom"];

            inddZuPdfBtn.onClick = function() { dialog.close(1); };
            pdfZuInddBtn.onClick = function() { dialog.close(2); };
            validierenBtn.onClick = function() { dialog.close(3); };
            abbrechenBtn.onClick = function() { dialog.close(0); };

            var ergebnis = dialog.show();

            if (ergebnis === 1) {
                this.fuehreRelinkDurch(".indd", ".pdf", "INDD -> PDF");
            } else if (ergebnis === 2) {
                this.fuehreRelinkDurch(".pdf", ".indd", "PDF -> INDD");
            } else if (ergebnis === 3) {
                this.fuehreValidierungDurch();
            }
        };

        /**
         * Startpunkt der gesamten Skriptausführung.
         */
        this.ausfuehren = function() {
            if (this.fuehreVorabpruefungenDurch()) {
                this.zeigeHauptdialog();
            }
        };
    }

    // --- Skriptausführung ---
    var manager = new LinkManager();
    manager.ausfuehren();

})();