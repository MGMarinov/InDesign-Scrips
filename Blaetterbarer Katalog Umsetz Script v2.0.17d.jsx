/**
 * Blaetterbarer Katalog Umsetz Script
 * Version: v2.0.17b
 * Datum: 2025-10-13
 * 
 * Zweck: Exportiert Seiten- und Objektdaten aus InDesign für den blätterbaren Katalog.
 * Hinweise:
 *  - Schreibe CSV-Header mit Versionswert "5" (nicht "5.0"), damit der Output bytegenau kompatibel ist.
 *  - Kein Deduplizieren vor dem Schreiben der CSV.
 *  - Für Text-Zeilen verwende die Legacy-Koordinatenberechnung (erster/letzter sichtbarer Buchstabe).
 *  - Verwende überall dieselbe Seitenreferenz (targetPage) für Verarbeitung und Koordinaten.
 */
var SKRIPT_NAME = "Blaetterbarer Katalog Umsetz Script";
var SKRIPT_VERSION = "v2.0.17b";
var SKRIPT_DATUM = "2025-10-13";
var CUSTOMER = "personalshop";
var KEEP_CSV_NAMES = false; 

var FALLBACK_SUCHPFAD = "X:/_Grafik 2024/10_Katalogseiten/";
var FALLBACK_SUCHTIEFE = 3;

var LOGGING_AKTIVIERT = true;
var CSV_VERSION = "5";

var EXPORT_MODE_DIRECT = "DIRECT";
var EXPORT_MODE_CONVERT = "CONVERT";

// Zuordnung für 4-seitige Quelldokumente (Katalog-Rechts/Links → Quelle 2/3)
var FOUR_PAGE_SIDE_MAP = { RIGHT: 2, LEFT: 3 };

var CSV_ZEILE_GRAPHIC = "G";
var CSV_ZEILE_TEXT = "T";
var CSV_ZEILE_PAGEITEM = "W";

var MAX_REKURSIONS_TIEFE = 3;

/**
 * Standard-Einstellungen für Export-Optionen
 */
var STANDARD_EXPORT_GRAPHICS = true;
var STANDARD_EXPORT_TEXTFRAMES = true;
var STANDARD_EXPORT_TEXT_IN_STORIES = true;
var STANDARD_EXPORT_TABLES = false;
var STANDARD_EXPORT_PAGEITEMS = false;

var REGEX_TABLE_IN_STORY = /\d{6}/;
var REGEX_TEXTFRAME_IN_STORY = /(\d{2}[.])?\d{3}[.]\d{3}/;


/**
 * HILFSFUNKTIONEN AUS DEM ORIGINALSKRIPT (1:1 ÜBERNOMMEN)
 */
function concatTempWithShapes(myParentObject, temp){
    var parentType = "" + myParentObject;
    while(parentType.indexOf("Polygon") > 0 || parentType.indexOf("Rectangle") > 0 || parentType.indexOf("Oval") > 0){
        temp = temp + "," + myParentObject.visibleBounds;
        myParentObject = myParentObject.parent;
        parentType = "" + myParentObject;
    }
    return temp;
}

function concatTempWithPageItems(myParentObject, temp){
    var parentType = "" + myParentObject;
    while(parentType.indexOf("Polygon") > 0 || parentType.indexOf("Rectangle") > 0 || parentType.indexOf("Oval") > 0 || parentType.indexOf("group") > 0){
        temp = temp + "," + myParentObject.visibleBounds;
        myParentObject = myParentObject.parent;
        parentType = "" + myParentObject;
    }
    return temp;
}


/**
 * Erstellt ein Export-Item
 *
 * @param {string} typ - "DIRECT"|"CONVERT"
 * @param {File} inddDatei - Die zu verarbeitende INDD-Datei
 * @param {string|null} pdfName - Ursprünglicher PDF-Name (optional)
 * @param {Page} page - Page-Objekt aus dem Hauptdokument
 * @param {Link} link - Original Link-Objekt
 * @returns {Object} ExportItem
 */
function erstelleExportItem(typ, inddDatei, pdfName, page, link) {
    return {
        typ: typ,
        inddDatei: inddDatei,
        pdfName: pdfName || null,
        page: page,
        link: link
    };
}

/**
 * Erstellt ein CSV-Daten-Model
 *
 * @param {Document} inddDoc - Das verarbeitete InDesign-Dokument
 * @param {Object} sourceItem - Das Quell-Item (ExportItem)
 * @param {string} zeilenDaten - Die gesammelten CSV-Zeilen
 * @param {string} verwendeterLayer - Der verwendete Layer-Name
 * @returns {Object} CsvDataModel
 */
function erstelleCsvDataModel(inddDoc, sourceItem, zeilenDaten, verwendeterLayer) {
    /**
     * Erstellt den CSV-Header aus einem Dokument
     * @param {Document} doc - Das InDesign-Dokument
     * @returns {string} CSV-Header-Zeile
     */
    function erstelleCsvHeader(doc) {
        var zeroPoint = doc.zeroPoint;
        var pageWidth = doc.documentPreferences.pageWidth;
        var pageHeight = doc.documentPreferences.pageHeight;
        return zeroPoint + "," + pageWidth + "," + pageHeight + "," + CSV_VERSION + "\n";
    }

    var header = erstelleCsvHeader(inddDoc);
    var zeilenAnzahl = 0;
    if (zeilenDaten && zeilenDaten.length > 0) {
        zeilenAnzahl = zeilenDaten.split('\n').length;
    }

    return {
        quellIndd: sourceItem.inddDatei.name,
        quellPdf: sourceItem.pdfName || null,
        elternSeite: sourceItem.page ? parseInt(sourceItem.page.name, 10) : -1,
        header: header,
        zeilen: zeilenDaten,
        metadaten: {
            verarbeitetAm: new Date(),
            zeilenAnzahl: zeilenAnzahl,
            dokumentName: inddDoc.name,
            verwendeterLayer: verwendeterLayer || "unbekannt"
        }
    };
}


/**
 * StringUtils - Hilfsfunktionen für String-Manipulation.
 */
function StringUtils() {
    this.entferneExtension = function(dateiName) {
        if (!dateiName) return "";
        var letzterpunkt = dateiName.lastIndexOf(".");
        if (letzterpunkt === -1) return dateiName;
        return dateiName.substring(0, letzterpunkt);
    };
}

/**
 * DateUtils - Hilfsfunktionen für Datum und Zeit.
 */
function DateUtils() {
    this.formatiereZeit = function(sekunden) {
        var h = Math.floor(sekunden / 3600);
        var m = Math.floor((sekunden % 3600) / 60);
        var s = Math.floor(sekunden % 60);
        var pad = function(num) { return (num < 10 ? '0' : '') + num; };
        return pad(h) + ":" + pad(m) + ":" + pad(s);
    };
}

/**
 * ErrorHandler - Fehlerbehandlung.
 */
function ErrorHandler() {
    this.formatiereFehler = function(fehlerObj, dateiInfo) {
        var basisNachricht = fehlerObj.message || "Unbekannter Fehler.";
        var zusatz = "";
        if (fehlerObj.line) {
            zusatz += " (Zeile: " + fehlerObj.line + ")";
        }
        return basisNachricht + zusatz;
    };

    this.sammleFehler = function(fehler, kontext, fehlerListe) {
        var fehlerMsg = this.formatiereFehler(fehler, kontext);
        fehlerListe.push({ betroffeneDatei: kontext, grund: fehlerMsg });
    };

    this.erstelleFehlerReport = function(fehlerListe) {
        if (!fehlerListe || fehlerListe.length === 0) return "";
        var report = "Bei den folgenden Dateien ist ein Problem aufgetreten:\n\n";
        for (var i = 0; i < fehlerListe.length; i++) {
            report += "- " + fehlerListe[i].betroffeneDatei + "\n";
            report += "  Grund: " + fehlerListe[i].grund + "\n\n";
        }
        return report;
    };
}

/**
 * Logger - Logging-System.
 */
function Logger() {
    this.logDatei = null;
    this.aktiviert = false;

    this.initialisiere = function(logDatei, aktiviert) {
        this.logDatei = logDatei;
        this.aktiviert = aktiviert;
        if (this.aktiviert && this.logDatei) {
            this.logDatei.encoding = "UTF-8";
        }
    };

    this.log = function(nachricht) {
        this._schreibeEintrag("[INFO] " + nachricht);
    };

    this.logFehler = function(fehler, kontext) {
        this._schreibeEintrag("[ERROR] " + (kontext ? kontext + ": " : "") + fehler);
    };
    
    this.logWarnung = function(nachricht) {
        this._schreibeEintrag("[WARN] " + nachricht);
    };
    
    this._schreibeEintrag = function(eintrag) {
        if (!this.aktiviert || !this.logDatei) return;
        try {
            this.logDatei.open("a");
            var zeitstempel = new Date().toTimeString().substr(0, 8);
            this.logDatei.writeln(zeitstempel + " - " + eintrag);
            this.logDatei.close();
        } catch (e) {}
    };
}

/**
 * FileService - Dateioperationen und Fallback-Suche.
 */
function FileService(dependencies) {
    this.logger = dependencies.logger;

    this.findeDateiMitFallback = function(originalPfad) {
        var datei = File(originalPfad);
        if (datei.exists) return datei;

        this.logger.log("Datei nicht am Originalpfad gefunden, starte Fallback-Suche: " + datei.name);
        var fallbackRoot = Folder(FALLBACK_SUCHPFAD);
        if (!fallbackRoot.exists) {
            this.logger.logWarnung("Fallback-Ordner existiert nicht: " + FALLBACK_SUCHPFAD);
            return null;
        }

        var alleTreffer = [];
        this._sucheRekursiv(fallbackRoot, datei.name, alleTreffer, FALLBACK_SUCHTIEFE, 0);

        if (alleTreffer.length === 0) {
            this.logger.logWarnung("Datei nicht gefunden: " + datei.name);
            return null;
        }

        if (alleTreffer.length > 1) {
            alleTreffer.sort(function(a, b) { return b.modified.getTime() - a.modified.getTime(); });
            this.logger.log("Mehrere Treffer gefunden, verwende neueste: " + alleTreffer[0].fsName);
        } else {
            this.logger.log("Datei via Fallback gefunden: " + alleTreffer[0].fsName);
        }
        return alleTreffer[0];
    };

    this._sucheRekursiv = function(ordner, dateiName, alleTreffer, maxTiefe, aktuelleTiefe) {
        if (aktuelleTiefe >= maxTiefe) return;
        var dateien = ordner.getFiles();
        for (var i = 0; i < dateien.length; i++) {
            var datei = dateien[i];
            if (datei instanceof File && datei.name.toLowerCase() === dateiName.toLowerCase()) {
                alleTreffer.push(datei);
            } else if (datei instanceof Folder) {
                this._sucheRekursiv(datei, dateiName, alleTreffer, maxTiefe, aktuelleTiefe + 1);
            }
        }
    };
    
    this.holeBasisname = function(dateiName) {
        var letzterpunkt = dateiName.lastIndexOf(".");
        return (letzterpunkt === -1) ? dateiName : dateiName.substring(0, letzterpunkt);
    };
}

/**
 * DocumentService - InDesign-Dokumentoperationen.
 */
function DocumentService(dependencies) {
    this.logger = dependencies.logger;
    this.errorHandler = dependencies.errorHandler;

    this.oeffneDokument = function(datei, userInteraction) {
        if (!datei.exists) throw new Error("Datei existiert nicht: " + datei.fsName);
        try {
            this.logger.log("Öffne Dokument: " + datei.name);
            return app.open(datei, !userInteraction, OpenOptions.OPEN_ORIGINAL);
        } catch (fehler) {
            var fehlermeldung = this.errorHandler.formatiereFehler(fehler, datei.name);
            this.logger.logFehler("Fehler beim Öffnen: " + fehlermeldung, datei.name);
            throw new Error("Fehler beim Öffnen: " + fehlermeldung);
        }
    };

    this.schliesseDokument = function(doc, saveOption) {
        try {
            if (doc && doc.isValid) {
                this.logger.log("Schließe Dokument: " + doc.name);
                doc.close(saveOption || SaveOptions.NO);
            }
        } catch (fehler) {
            this.logger.logWarnung("Fehler beim Schließen des Dokuments: " + fehler.message);
        }
    };

    this.holeAlleLayer = function(doc) {
        var layerNamen = [];
        for (var i = 0; i < doc.layers.length; i++) {
            layerNamen.push(doc.layers[i].name);
        }
        return layerNamen;
    };
}

/**
 * CsvService - Verantwortlich für das Schreiben von CSV-Dateien.
 */
function CsvService(dependencies) {
    this.logger = dependencies.logger;

    this.schreibeEinzelneCsv = function(datenArray, csvDatei) {
        this.logger.log("Erstelle CSV-Datei: " + csvDatei.name);
        var inhalt = (datenArray.length > 0) ? datenArray[0].header : "";
        for (var i = 0; i < datenArray.length; i++) {
            if (datenArray[i] && datenArray[i].zeilen) {
                inhalt += datenArray[i].zeilen;
            }
        }
        
        
        inhalt = this._entferneDuplikate(inhalt);
csvDatei.encoding = "UTF-8";
        try {
            csvDatei.open("w");
            if (!csvDatei.write(inhalt)) throw new Error("Fehler beim Schreiben der CSV-Datei");
            csvDatei.writeln("");
            csvDatei.close();
            this.logger.log("CSV-Datei erfolgreich erstellt: " + csvDatei.fsName);
            return csvDatei;
        } catch (fehler) {
            this.logger.logFehler("Fehler beim Schreiben der CSV: " + fehler.message, csvDatei.name);
            throw fehler;
        }
    };

    this._entferneDuplikate = function(inhalt) {
        var zeilen = inhalt.split('\n');
        var seen = {};
        var out = [];
        for (var i = 0; i < zeilen.length; i++) {
            var z = zeilen[i];
            if (z.length === 0) { continue; }
            if (!seen.hasOwnProperty(z)) {
                seen[z] = true;
                out.push(z);
            }
        }
        return out.join('\n');
    };
}


/**
 * DataProcessors - Sammlung von Objekten zur Extraktion von Daten (Graphics, Text, etc.)
 */
function DataProcessors(dependencies) {
    this.logger = dependencies.logger;

    this.verarbeiteGraphics = function(currentPage, page, settings) {
        var temp = "";
        if (!currentPage) return temp;
        for (var j = 0; j < currentPage.allPageItems.length; j++) {
            var pageItem = currentPage.allPageItems[j];
            // exakte Reihenfolge wie im Ursprung: auf PageItem-Ebene über alle Grafiken iterieren
            for (var i = 0; i < pageItem.allGraphics.length; i++) {
                var graphic = pageItem.allGraphics[i];
                if (graphic && graphic.itemLink && graphic.itemLink.name && graphic.visibleBounds) {
                    var parentObject = graphic.parent;
                    var graphicFileName = File.encode(graphic.itemLink.name).replace(/,/g, "_");
                    if (page && !isNaN(page.name) && page.name > 0) {
                        if (settings.pages) settings.pages.push(page.name);
                        var line = "G," + page.name + "," + page.index + "," + graphicFileName + "," + graphic.visibleBounds;
                        line = concatTempWithShapes(parentObject, line);
                        temp += line + "\n";
                    }
                }
            }
        }
        return temp;
    };
    
    this.verarbeiteTextFrames = function(currentPage, page, settings) {
        var temp = "";
        for (var i = 0; i < currentPage.allPageItems.length; i++) {
            var pageItem = currentPage.allPageItems[i];
            if (pageItem instanceof TextFrame) {
                if(page) {
                    temp = this._verarbeiteZeilen(pageItem.lines, temp, page.name, settings.regex.textFrame, "T", settings);
                }
            }
        }
        return temp;
    };
    
    this.verarbeiteStories = function(doc, currentPage, page, settings) {
        var temp = "";
        for (var i = 0; i < doc.stories.length; i++) {
            var currentStory = doc.stories.item(i);
            for (var j = 0; j < currentStory.textFrames.length; j++) {
                var currentTextFrame = currentStory.textFrames.item(j);
                if (currentTextFrame.parentPage && currentTextFrame.parentPage.name == currentPage.name) {
                    if (page) {
                         temp = this._verarbeiteZeilen(currentTextFrame.lines, temp, page.name, settings.regex.textFrame, "T", settings);
                    }
                }
            }
        }
        return temp;
    };

    this.verarbeiteTables = function(doc, currentPage, page, settings) {
        var temp = "";
        for (var i = 0; i < doc.stories.length; i++) {
            var currentStory = doc.stories.item(i);
            for (var j = 0; j < currentStory.textFrames.length; j++) {
                var currentTextFrame = currentStory.textFrames.item(j);
                if (currentTextFrame.parentPage && currentTextFrame.parentPage.name == currentPage.name) {
                    for (var k = 0; k < currentTextFrame.tables.length; k++) {
                        var table = currentTextFrame.tables.item(k);
                        if(page) {
                            temp = this._verarbeiteTabelle(table, temp, page.name, settings.regex.table, settings);
                        }
                    }
                }
            }
        }
        return temp;
    };

    this._verarbeiteTabelle = function(table, temp, pageName, regex, settings) {
        for (var i = 0; i < table.cells.length; i++) {
            var cell = table.cells.item(i);
            if (cell.textStyleRanges.length > 0) {
                temp = this._verarbeiteZeilen(cell.textStyleRanges.item(0).lines, temp, pageName, regex, "T", settings);
            }
        }
        return temp;
    };
    
    this.verarbeitePageItems = function(currentPage, page, settings) {
        var temp = "";
        for (var i = 0; i < currentPage.allPageItems.length; i++) {
            var pageItem = currentPage.allPageItems[i];
            if (pageItem.label !== "") {
                var visibleBounds = pageItem.visibleBounds;
                var linkType = (pageItem.allGraphics.length === 1 && pageItem.allGraphics[0].itemLink) ? pageItem.allGraphics[0].itemLink.name : "unknown";
                if (page && !isNaN(page.name) && page.name > 0) {
                    if (settings.pages) settings.pages.push(page.name);
                    var line = "W," + page.name + ",\"" + File.encode(linkType) + "\",\"" + File.encode(pageItem.label) + "\"," + visibleBounds;
                    line = concatTempWithPageItems(pageItem.parent, line);
                    temp += line + "\n";
                }
            }
        }
        return temp;
    };
    
    this._verarbeiteZeilen = function(lines, temp, pageName, regex, letter, settings) {
        for (var i = 0; i < lines.length; i++) {
            try {
                var line = lines.item(i);
                var lineContents = "" + line.contents;
                var characters = line.characters;
                var result = regex.exec(lineContents);
                if (result === undefined || result === null) {
                    continue;
                }
                var artNr = result[0];
                if (settings.pages) settings.pages.push(pageName);

                var xLeft = characters[0].horizontalOffset;
                var yTop = characters[0].baseline - line.ascent - line.descent;
                var yBottom = characters[0].baseline;
                var xRight = characters[characters.length - 1].horizontalOffset;

                temp += letter + "," + pageName + "," + artNr + "," + yTop + "," + xLeft + "," + yBottom + "," + xRight + "\n";
            } catch (e) { /* still weiter */ }
        }
        return temp;
    };
    
    /**
     * Ermittelt die Zielseite im geöffneten Quelldokument.
     * Bevorzugt die tatsächlich platzierte PDF-Seitennummer; fällt sonst auf eine Seiten-Heuristik zurück.
     * @param {Object} sourceItem - Export-Item mit Link und Katalog-Seite.
     * @param {Document} inddDoc - Geöffnetes InDesign-Dokument (Quelle).
     * @returns {number} 1-basierte Seitennummer im Quelldokument.
     */
    this.ermittleZielseite = function(sourceItem, inddDoc) {
        // Bevorzugt die tatsächlich platzierte PDF-Seitennummer, ansonsten heuristische Abbildung.
        try {
            if (inddDoc.pages.length === 1) return 1;

            // 1) Versuche, die PDF-Seite direkt zu lesen (nur im CONVERT-Zweig relevant).
            try {
                if (sourceItem.typ === EXPORT_MODE_CONVERT &&
                    sourceItem.link && sourceItem.link.parent &&
                    sourceItem.link.parent.hasOwnProperty("pdfAttributes") &&
                    sourceItem.link.parent.pdfAttributes) {
                    var pdfNum = sourceItem.link.parent.pdfAttributes.pageNumber; // 1-basiert
                    if (pdfNum >= 1 && pdfNum <= inddDoc.pages.length) {
                        this.logger.log("Zielseite aus PDF.pageNumber: " + pdfNum);
                        return pdfNum;
                    }
                }
            } catch (ePdf) {}

            // 2) Direkte Page-Referenz am platzierten Objekt (z. B. ImportedPage).
            var linkParent = sourceItem.link ? sourceItem.link.parent : null;
            if (linkParent && linkParent.hasOwnProperty("pageNumber")) {
                var pageNum = linkParent.pageNumber; // 1-basiert
                if (pageNum >= 1 && pageNum <= inddDoc.pages.length) {
                    this.logger.log("Zielseite aus ImportedPage.pageNumber: " + pageNum);
                    return pageNum;
                }
            }

            // 3) Heuristik für 4-seitige Quelldokumente im CONVERT-Zweig:
            //    Katalog RIGHT → Quelle Seite 2, Katalog LEFT → Quelle Seite 3.
            if (sourceItem.typ === EXPORT_MODE_CONVERT && inddDoc.pages.length === 4) {
                var isRight = (sourceItem.page.side === PageSideOptions.RIGHT_HAND);
                var ziel = isRight ? FOUR_PAGE_SIDE_MAP.RIGHT : FOUR_PAGE_SIDE_MAP.LEFT;
                this.logger.log("Zielseite aus Seitenlogik (4p): " + (isRight ? "RIGHT→" : "LEFT→") + ziel);
                return ziel;
            }

            // 4) Heuristik für 2-seitige Quelldokumente im CONVERT-Zweig (falls vorhanden):
            if (sourceItem.typ === EXPORT_MODE_CONVERT && inddDoc.pages.length === 2) {
                var isRight2 = (sourceItem.page.side === PageSideOptions.RIGHT_HAND);
                var ziel2 = isRight2 ? 2 : 1;
                this.logger.log("Zielseite aus Seitenlogik (2p): " + (isRight2 ? "RIGHT→2" : "LEFT→1"));
                return ziel2;
            }

        } catch (e) {
            this.logger.logWarnung("Konnte Zielseite nicht automatisch ermitteln, verwende Seite 1: " + e.message);
        }
        return 1;
    };

    this.verarbeiteDokument = function(inddDoc, settings, sourceItem) {
        var hauptDokPage = sourceItem.page;
        var pageNumber = this.ermittleZielseite(sourceItem, inddDoc);
        
        if (isNaN(pageNumber) || pageNumber < 1 || pageNumber > inddDoc.pages.length) {
            throw new Error("Ungültige Seitennummer: " + pageNumber + " in Dokument: " + inddDoc.name);
        }
        var targetPage = inddDoc.pages[pageNumber - 1];
        var daten = "";

        if (settings.exportOptionen.graphics) {
            daten += this.verarbeiteGraphics(targetPage, hauptDokPage, settings);
        }
        
        if (settings.exportOptionen.textInStories) {
            daten += this.verarbeiteStories(inddDoc, targetPage, hauptDokPage, settings);
        }
        if (settings.exportOptionen.textFrames) {
            daten += this.verarbeiteTextFrames(targetPage, hauptDokPage, settings);
        }
        if (settings.exportOptionen.tables) {
            daten += this.verarbeiteTables(inddDoc, targetPage, hauptDokPage, settings);
        }
        if (settings.exportOptionen.pageItems) {
            daten += this.verarbeitePageItems(targetPage, hauptDokPage, settings);
        }

        return erstelleCsvDataModel(inddDoc, sourceItem, daten, settings.layer.name);
    };
}


/**
 * DialogFactory - Erstellt alle UI-Dialoge.
 */
function DialogFactory() {
    this.erstelleCsvExportEinstellungenDialog = function(layerNamen, standardOrdner) {
        var dialog = new Window("dialog", "CSV Export Einstellungen");
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 15;
        dialog.margins = 15;

        var layerPanel = dialog.add("panel", undefined, "Ebene für den Export auswählen");
        layerPanel.alignChildren = "fill";
        var layerOptionen = layerNamen.concat(["** Alle Ebenen **"]);
        var layerDropdown = layerPanel.add("dropdownlist", undefined, layerOptionen);
        
        var standardAuswahlIndex = 0;
        for (var i = 0; i < layerNamen.length; i++) {
            if (layerNamen[i].indexOf("DA Seiten") === 0) {
                standardAuswahlIndex = i;
                break;
            }
        }
        layerDropdown.selection = standardAuswahlIndex;
        layerDropdown.preferredSize.width = 350;

        var zielPanel = dialog.add("panel", undefined, "Export-Verzeichnis");
        zielPanel.alignChildren = "fill";
        var zielGroup = zielPanel.add("group", undefined, { orientation: "row", alignChildren: ["fill", "center"] });
        zielGroup.add("statictext", undefined, "Ziel:").preferredSize.width = 40;
        var zielPfadText = zielGroup.add("edittext", undefined, standardOrdner);
        zielPfadText.preferredSize.width = 250;
        var durchsuchenBtn = zielGroup.add("button", undefined, "Durchsuchen...");

        durchsuchenBtn.onClick = function() {
            var folder = Folder.selectDialog("Bitte wähle das Export-Verzeichnis:", Folder(zielPfadText.text));
            if (folder) zielPfadText.text = folder.fsName;
        };

        var buttonGroup = dialog.add("group", undefined, { orientation: "row", alignment: "right" });
        buttonGroup.add("button", undefined, "Abbrechen", { name: "cancel" });
        buttonGroup.add("button", undefined, "Export starten", { name: "ok" });

        if (dialog.show() === 1) {
            if (!layerDropdown.selection) {
                alert("Bitte wähle eine Ebene aus.");
                return null;
            }
            
            var zielFolder = Folder(zielPfadText.text);
            if (!zielFolder.exists) {
                alert("Das angegebene Verzeichnis existiert nicht:\n" + zielPfadText.text);
                return null;
            }

            return {
                layerName: (layerDropdown.selection.text === "** Alle Ebenen **") ? null : layerDropdown.selection.text,
                zielOrdner: zielFolder,
                alleEbenen: layerDropdown.selection.text === "** Alle Ebenen **"
            };
        }
        return null;
    };

    this.erstelleAbschnittsOptionenDialog = function() {
        var dialog = new Window("dialog", "Fehler in der Seitenstruktur");
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        
        var text = "Das Dokument enthält mehrere Seiten mit der Seitenzahl '1' aufgrund von Abschnittsdefinitionen.\n\nSoll das Skript die Seitennummerierung automatisch korrigieren?";
        dialog.add("statictext", [15, 15, 435, 75], text, { multiline: true });

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        var abbrechenBtn = buttonGroup.add("button", undefined, "Prozess abbrechen");
        var korrigierenBtn = buttonGroup.add("button", undefined, "Automatisch korrigieren", { name: "ok" });
        
        abbrechenBtn.onClick = function() { dialog.close(0); };
        korrigierenBtn.onClick = function() { dialog.close(1); };

        return dialog.show() === 1;
    };

    this.erstelleFortschrittsDialog = function(nachricht, dateiUtils) {
        var w = new Window("palette", "Fortschritt", undefined, { closeButton: false });
        w.alignChildren = "fill";
        w.margins = 15;
        var t = w.add("statictext", undefined, nachricht, { properties: { preferredSize: [450, -1] }});
        var b = w.add("progressbar", undefined, 0, 100, { properties: { preferredSize: [450, 20] }});
        var timeGroup = w.add('group', undefined, { orientation: 'row', alignment: 'fill' });
        var elapsedLabel = timeGroup.add("statictext", undefined, "Verstrichene Zeit: 00:00:00");
        elapsedLabel.alignment = "left";
        var estimateLabel = timeGroup.add("statictext", undefined, "Verbleibende Zeit: Berechnung...");
        estimateLabel.alignment = "right";
        var abgebrochen = false;
        w.add("button", undefined, "Abbrechen", { alignment: "center" }).onClick = function() {
            abgebrochen = true;
            w.close();
        };
        var startTime = new Date();
        w.show();
        return {
            isCancelled: function() { return abgebrochen; },
            update: function(current, max, msg) {
                if (abgebrochen) return;
                b.maxvalue = max;
                b.value = current;
                if (msg) t.text = msg;
                var elapsedSeconds = (new Date().getTime() - startTime.getTime()) / 1000;
                elapsedLabel.text = "Verstrichene Zeit: " + dateiUtils.formatiereZeit(elapsedSeconds);
                if (current > 0) {
                    var remainingSeconds = (elapsedSeconds / current) * (max - current);
                    estimateLabel.text = "Verbleibende Zeit: " + dateiUtils.formatiereZeit(remainingSeconds);
                }
                w.update();
            },
            close: function() { w.close(); }
        };
    };
}

/**
 * ExportCoordinator - Koordiniert den gesamten Export-Prozess.
 */
function ExportCoordinator(dependencies) {
    this.documentService = dependencies.documentService;
    this.dataProcessors = dependencies.dataProcessors;
    this.dialogFactory = dependencies.dialogFactory;
    this.logger = dependencies.logger;
    this.errorHandler = dependencies.errorHandler;
    this.dateUtils = dependencies.dateUtils;
    this.fileService = dependencies.fileService;

    this.koordiniereExport = function(exportItems, settings) {
        this.logger.log("Starte Export-Koordination: " + exportItems.length + " Items");
        var progressDialog = this.dialogFactory.erstelleFortschrittsDialog("INDD Dateien werden verarbeitet...", this.dateUtils);
        var alleDaten = [];
        var fehlerListe = [];
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

        try {
            for (var i = 0; i < exportItems.length; i++) {
                if (progressDialog.isCancelled()) {
                    this.logger.log("Export wurde vom Benutzer abgebrochen");
                    break;
                }
                var item = exportItems[i];
                var dateiName = item.inddDatei ? item.inddDatei.name : (item.pdfName || "unbekannt");
                progressDialog.update(i + 1, exportItems.length, "Verarbeite: " + dateiName + " (" + (i + 1) + "/" + exportItems.length + ")");
                try {
                    var csvData = this._verarbeiteItem(item, settings);
                    if (csvData) alleDaten.push(csvData);
                } catch (fehler) {
                    this.errorHandler.sammleFehler(fehler, dateiName, fehlerListe);
                }
            }
        } finally {
            progressDialog.close();
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
        }
        return { daten: alleDaten, fehler: fehlerListe };
    };

    this._verarbeiteItem = function(item, settings) {
        var inddDatei = item.inddDatei;
        if (item.typ === EXPORT_MODE_CONVERT) {
            var basisName = this.fileService.holeBasisname(item.pdfName);
            var inddPfadOriginal = File(item.link.filePath).path + "/" + basisName + ".indd";
            inddDatei = this.fileService.findeDateiMitFallback(inddPfadOriginal);
            item.inddDatei = inddDatei; 
        }

        if (!inddDatei || !inddDatei.exists) {
            throw new Error("INDD-Datei nicht gefunden für: " + (item.pdfName || item.inddDatei.name));
        }

        var doc = null;
        try {
            doc = this.documentService.oeffneDokument(inddDatei, false);
            return this.dataProcessors.verarbeiteDokument(doc, settings, item);
        } finally {
            if (doc) this.documentService.schliesseDokument(doc, SaveOptions.NO);
        }
    };
}

/**
 * CsvExportWorkflow - Hauptlogik für den CSV-Export.
 */
function CsvExportWorkflow(dependencies) {
    this.exportCoordinator = dependencies.exportCoordinator;
    this.csvService = dependencies.csvService;
    this.dialogFactory = dependencies.dialogFactory;
    this.logger = dependencies.logger;
    this.documentService = dependencies.documentService;
    this.errorHandler = dependencies.errorHandler;
    this.pages = [];

    this.fuehreAus = function() {
        this.logger.log("Starte CSV-Export-Workflow");
        var doc = app.activeDocument;
        var layerNamen = this.documentService.holeAlleLayer(doc);
        if (layerNamen.length === 0) {
            alert("Keine Ebenen im Dokument gefunden.", SKRIPT_NAME);
            return;
        }

        var einstellungen = null;
        var prozessFortsetzen = false;

        while (!prozessFortsetzen) {
            einstellungen = this.dialogFactory.erstelleCsvExportEinstellungenDialog(layerNamen, doc.filePath);

            if (!einstellungen) {
                this.logger.log("Benutzer hat den Einstellungsdialog abgebrochen.");
                return;
            }

            if (einstellungen.alleEbenen) {
                var warnungText = "Achtung!\n\nDie Verarbeitung aller Ebenen kann den Prozess erheblich verlängern.\n\nMöchtest Du fortfahren?";
                var bestaetigt = confirm(warnungText, false, "Lange Prozesszeit");
                if (bestaetigt) {
                    prozessFortsetzen = true;
                }
            } else {
                prozessFortsetzen = true;
            }
        }

        var linksAufEbene = this._bestimmeLinksInDokument(doc, einstellungen.alleEbenen ? null : einstellungen.layerName);
        if (linksAufEbene.length === 0) {
            alert("Keine verarbeitbaren INDD- oder PDF-Verknüpfungen auf der gewählten Ebene(n) gefunden.", SKRIPT_NAME);
            return;
        }
        
        var exportItems = this._erstelleExportItemsAusLinks(linksAufEbene);
        if (exportItems.length === 0) {
            alert("Keine gültigen INDD- oder PDF-Dateien zum Verarbeiten gefunden.", SKRIPT_NAME);
            return;
        }

        var settings = this._erstelleStandardSettings();
        settings.layer.name = einstellungen.alleEbenen ? "Alle Ebenen" : einstellungen.layerName;
        
        var ergebnis = this.exportCoordinator.koordiniereExport(exportItems, settings);

        if (ergebnis.daten.length > 0) {
            try {
                var csvDatei = this._erstelleCsvDateiNamen(einstellungen.zielOrdner, doc);
                this.csvService.schreibeEinzelneCsv(ergebnis.daten, csvDatei);
                this._zeigeErgebnis(ergebnis.daten.length, ergebnis.fehler, csvDatei.fsName);
            } catch (e) {
                alert("Fehler beim Speichern der CSV-Datei:\n" + e.message);
                this.logger.logFehler("Fehler beim Speichern der CSV: " + e.message);
            }
        } else {
            this._zeigeErgebnis(0, ergebnis.fehler, null);
        }
    };

    this._bestimmeLinksInDokument = function(doc, layerName) {
        var linksImDokument = [];
        var zielLayer = layerName ? doc.layers.itemByName(layerName) : null;

        for (var p = 0; p < doc.pages.length; p++) {
            var currentPage = doc.pages[p];

            for (var k = 0; k < currentPage.allPageItems.length; k++) {
                var currentPageItem = currentPage.allPageItems[k];

                // Filter: nur gewünschte Ebene
                var ebeneOk = (!zielLayer) || (currentPageItem.itemLayer === zielLayer);

                // (1) INDD/PDF über importedPages (z.B. platzierte INDD-Seiten)
                if (currentPageItem instanceof Rectangle && currentPageItem.hasOwnProperty("importedPages")) {
                    for (var j = 0; j < currentPageItem.importedPages.length; j++) {
                        var importedPage = currentPageItem.importedPages[j];
                        try {
                            var link = importedPage.itemLink;
                            if (!link) continue;
                            var nameLower = ("" + link.name).toLowerCase();
                            if (nameLower.indexOf(".indd") === -1 && nameLower.indexOf(".pdf") === -1) continue;
                            if (!zielLayer || importedPage.itemLayer === zielLayer) {
                                linksImDokument.push({ itemLink: link, page: currentPage });
                            }
                        } catch (e) {}
                    }
                }

                // (2) PDF/INDD als normale Grafiklink (z.B. platzált PDF-ek grafikai hivatkozásként)
                // Reihenfolge: wie PageItem-Iteration, dann Grafiken
                if (currentPageItem.allGraphics && currentPageItem.allGraphics.length > 0) {
                    for (var g = 0; g < currentPageItem.allGraphics.length; g++) {
                        try {
                            var graphic = currentPageItem.allGraphics[g];
                            var gLink = graphic && graphic.itemLink ? graphic.itemLink : null;
                            if (!gLink) continue;
                            var gNameLower = ("" + gLink.name).toLowerCase();
                            if (gNameLower.indexOf(".indd") === -1 && gNameLower.indexOf(".pdf") === -1) continue;
                            if (ebeneOk) {
                                linksImDokument.push({ itemLink: gLink, page: currentPage });
                            }
                        } catch (e2) {}
                    }
                }
            }
        }

        this.logger.log("Gefundene Links auf Ebene '" + (layerName || "Alle") + "': " + linksImDokument.length);
        return linksImDokument;
    };

    this._erstelleExportItemsAusLinks = function(links) {
        var items = [];
        for (var i = 0; i < links.length; i++) {
            var linkInfo = links[i];
            var linkNameLower = linkInfo.itemLink.name.toLowerCase();
            var typ = (linkNameLower.indexOf(".indd") > -1) ? EXPORT_MODE_DIRECT : EXPORT_MODE_CONVERT;
            var inddDatei = (typ === EXPORT_MODE_DIRECT) ? File(linkInfo.itemLink.filePath) : null;
            var pdfName = (typ === EXPORT_MODE_CONVERT) ? linkInfo.itemLink.name : null;
            
            if (typ === EXPORT_MODE_DIRECT && !inddDatei.exists) {
                this.logger.logWarnung("INDD-Datei existiert nicht: " + inddDatei.fsName);
                continue;
            }
            items.push(erstelleExportItem(typ, inddDatei, pdfName, linkInfo.page, linkInfo.itemLink));
        }
        return items;
    };

    this._erstelleStandardSettings = function() {
        return {
            layer: { name: "" },
            exportOptionen: {
                graphics: STANDARD_EXPORT_GRAPHICS,
                textFrames: STANDARD_EXPORT_TEXTFRAMES,
                textInStories: STANDARD_EXPORT_TEXT_IN_STORIES,
                tables: STANDARD_EXPORT_TABLES,
                pageItems: STANDARD_EXPORT_PAGEITEMS
            },
            regex: {
                textFrame: REGEX_TEXTFRAME_IN_STORY,
                table: REGEX_TABLE_IN_STORY
            },
            pages: this.pages
        };
    };

    this._erstelleCsvDateiNamen = function(ordner, doc) {
        this.pages.sort(function(a, b) { return parseInt(a) - parseInt(b); });
    
        var dateiName;
        if (KEEP_CSV_NAMES) {
            dateiName = doc.name + ".csv";
        } else {
            if (this.pages.length === 0) {
                this.logger.logWarnung("Keine Seiten im 'pages'-Array gefunden. Fallback-Dateiname wird verwendet.");
                dateiName = CUSTOMER + "_" + doc.name.replace(/\.indd$/i, "") + "_no_pages_found.csv";
            } else {
                var erstePage = this.pages[0];
                var letztePage = this.pages[this.pages.length - 1];
                dateiName = CUSTOMER + "_" + erstePage + "-" + letztePage + ".csv";
            }
        }
        return new File(ordner.fsName + "/" + dateiName);
    };

    this._zeigeErgebnis = function(erfolgreich, fehler, csvPfad) {
        var gesamt = erfolgreich + fehler.length;
        var nachricht = "CSV-Export abgeschlossen" + (fehler.length > 0 ? " (mit Fehlern)" : "!") + "\n\n" +
                        "Verarbeitete Dateien: " + gesamt + "\n" +
                        "Erfolgreich: " + erfolgreich + "\n" +
                        "Fehlgeschlagen: " + fehler.length + "\n";
        if (csvPfad) {
            nachricht += "\nCSV-Datei gespeichert unter:\n" + csvPfad + "\n";
        }
        if (fehler.length > 0) {
            nachricht += "\n" + this.errorHandler.erstelleFehlerReport(fehler);
        }
        alert(nachricht, SKRIPT_NAME + " - Export Ergebnis");
    };
}


/**
 * CsvExportManager - Haupt-Controller für den Standalone-Export.
 */
function CsvExportManager(dependencies) {
    this.csvExportWorkflow = dependencies.csvExportWorkflow;
    this.logger = dependencies.logger;
    this.errorHandler = dependencies.errorHandler;
    this.dialogFactory = dependencies.dialogFactory;

    this.starte = function() {
        if (!this._initialisiere()) return;

        try {
            this.csvExportWorkflow.fuehreAus();
        } catch (fehler) {
            var fehlerMsg = this.errorHandler.formatiereFehler(fehler, "CsvExportManager");
            alert("Kritischer Fehler:\n" + fehlerMsg, SKRIPT_NAME + " - Fehler");
            this.logger.logFehler(fehlerMsg, "CsvExportManager.starte");
        } finally {
            this.logger.log("=== Script beendet ===");
        }
    };
    
    this.korrigiereAbschnitte = function(doc) {
        this.logger.logWarnung("Korrigiere Abschnittsoptionen für das Dokument.");
        // Starte die Schleife bei der ZWEITEN Seite (Index 1), da die erste Seite immer ein Abschnittsanfang ist.
        for (var i = 1; i < doc.pages.length; i++) {
            var page = doc.pages[i];
// Hinweis: Diese Kommentarzeile wurde ins Deutsche übernommen (du-Form).
            if (page.name === "1") {
                try {
                    // Wenn ja, stelle die Eigenschaft der Sektion so ein, dass die Nummerierung fortgesetzt wird.
                    // Dies entspricht dem Deaktivieren von "Neuen Abschnitt anfangen".
                    page.appliedSection.continueNumbering = true;
                    this.logger.log("Abschnittsanfang auf Dokumentseite " + (i + 1) + " (ursprünglich als '" + page.name + "' benannt) wurde korrigiert.");
                } catch(e) {
                    this.logger.logFehler("Fehler bei der Korrektur von Seite " + (i + 1) + ": " + e.message);
                }
            }
        }
    };

    this._initialisiere = function() {
        if (app.documents.length === 0) {
            alert("Bitte öffne ein Dokument, bevor Du dieses Skript ausführst.", SKRIPT_NAME);
            return false;
        }
        var doc = app.activeDocument;
        if (!doc.saved || !doc.filePath) {
            alert("Bitte speichere Dein Dokument zuerst.", SKRIPT_NAME);
            return false;
        }
        
        var pageOneCount = 0;
        for (var i = 0; i < doc.pages.length; i++) {
            if (doc.pages[i].name === "1") {
                pageOneCount++;
            }
        }

        if (pageOneCount > 1) {
            var korrigieren = this.dialogFactory.erstelleAbschnittsOptionenDialog();
            if (korrigieren) {
                this.korrigiereAbschnitte(doc);
                alert("Die Seitennummerierung wurde korrigiert. Der Export wird nun fortgesetzt.");
            } else {
                this.logger.logWarnung("Benutzer hat den Prozess wegen fehlerhafter Abschnittsoptionen abgebrochen.");
                return false;
            }
        }

        if (LOGGING_AKTIVIERT) {
            var logDatei = new File(doc.filePath + "/" + doc.name.replace(/\.indd$/i, "") + "_Umsetz_log.txt");
            this.logger.initialisiere(logDatei, true);
        }

        this.logger.log("=== " + SKRIPT_NAME + " " + SKRIPT_VERSION + " ===");
        this.logger.log("Dokument: " + doc.name);
        return true;
    };
}


(function main() {
    /**
     * Erstellt den Dependency Container
     */
    function erstelleDependencyContainer() {
        var container = {};
        
        container.errorHandler = new ErrorHandler();
        container.logger = new Logger();
        container.dateUtils = new DateUtils();
        container.stringUtils = new StringUtils();
        
        container.fileService = new FileService({ logger: container.logger });
        container.documentService = new DocumentService({ logger: container.logger, errorHandler: container.errorHandler });
        container.csvService = new CsvService({ logger: container.logger });
        
        container.dataProcessors = new DataProcessors({ logger: container.logger });
        container.dialogFactory = new DialogFactory();
        
        container.exportCoordinator = new ExportCoordinator({
            documentService: container.documentService,
            dataProcessors: container.dataProcessors,
            dialogFactory: container.dialogFactory,
            logger: container.logger,
            errorHandler: container.errorHandler,
            dateUtils: container.dateUtils,
            fileService: container.fileService
        });

        container.csvExportWorkflow = new CsvExportWorkflow({
            exportCoordinator: container.exportCoordinator,
            csvService: container.csvService,
            dialogFactory: container.dialogFactory,
            logger: container.logger,
            documentService: container.documentService,
            errorHandler: container.errorHandler
        });

        container.csvExportManager = new CsvExportManager({
            csvExportWorkflow: container.csvExportWorkflow,
            logger: container.logger,
            errorHandler: container.errorHandler,
            dialogFactory: container.dialogFactory
        });

        return container;
    }

    try {
        var container = erstelleDependencyContainer();
        container.csvExportManager.starte();
    } catch (e) {
        alert("Kritischer Fehler beim Starten des Scripts:\n" + e.message + " (Zeile: " + e.line + ")", SKRIPT_NAME + " - Kritischer Fehler");
    }
})();
