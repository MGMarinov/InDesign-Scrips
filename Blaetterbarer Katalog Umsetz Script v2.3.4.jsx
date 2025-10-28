/**
 * Blaetterbarer Katalog Umsetz Script
 * Version: v2.3.3
 * Zweck: Exportiert Seiten- und Objektdaten aus InDesign für den blätterbaren Katalog.
 * Nutzt ein intelligentes Caching-System, um die Verarbeitung zu beschleunigen,
 * indem bereits gescannte und unveränderte Dateien nicht erneut verarbeitet werden.
 *
 * Änderungen v2.3.3:
 * - Neues Feature: Fortschrittsdialog während der Cache-Prüfung
 *   - Zeigt aktuell geprüfte Datei an
 *   - Zeigt verstrichene und verbleibende Zeit
 *   - Erlaubt Abbruch während der Prüfung
 *
 * Änderungen v2.3.2:
 * - Bugfix: Cache-Speicherung verwendet nun konsistent den ursprünglichen Link-Namen
 *   als Schlüssel, auch wenn die Datei an einem anderen Ort gefunden wurde
 * - Bugfix: JSON Polyfill eingebaut für ExtendScript-Kompatibilität
 *
 * Hinweise:
 * - Schreibe CSV-Header mit Versionswert "5" (nicht "5.0"), damit der Output bytegenau kompatibel ist.
 * - Kein Deduplizieren vor dem Schreiben der CSV.
 * - Für Text-Zeilen verwende die Legacy-Koordinatenberechnung (erster/letzter sichtbarer Buchstabe).
 * - Verwende überall dieselbe Seitenreferenz (targetPage) für Verarbeitung und Koordinaten.
 */

// JSON Polyfill für ExtendScript (Minified) - Public Domain
if(typeof JSON!=="object"){JSON={}}(function(){"use strict";var rx_one=/^[\],:{}\s]*$/;var rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g;var rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g;var rx_four=/(?:^|:|,)(?:\s*\[)+/g;var rx_escapable=/[\\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;var rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;function f(n){return n<10?"0"+n:n}function this_value(){return this.valueOf()}if(typeof Date.prototype.toJSON!=="function"){Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z":null};Boolean.prototype.toJSON=this_value;Number.prototype.toJSON=this_value;String.prototype.toJSON=this_value}var gap;var indent;var meta={"\b":"\\b","\t":"\\t","\n":"\\n","\f":"\\f","\r":"\\r",'"':'\\"',"\\":"\\\\"};var rep;function quote(string){rx_escapable.lastIndex=0;return rx_escapable.test(string)?'"'+string.replace(rx_escapable,function(a){var c=meta[a];return typeof c==="string"?c:"\\u"+("0000"+a.charCodeAt(0).toString(16)).slice(-4)})+'"':'"'+string+'"'}function str(key,holder){var i;var k;var v;var length;var mind=gap;var partial;var value=holder[key];if(value&&typeof value==="object"&&typeof value.toJSON==="function"){value=value.toJSON(key)}if(typeof rep==="function"){value=rep.call(holder,key,value)}switch(typeof value){case"string":return quote(value);case"number":return isFinite(value)?String(value):"null";case"boolean":case"null":return String(value);case"object":if(!value){return"null"}gap+=indent;partial=[];if(Object.prototype.toString.apply(value)==="[object Array]"){length=value.length;for(i=0;i<length;i+=1){partial[i]=str(i,value)||"null"}v=partial.length===0?"[]":gap?"[\n"+gap+partial.join(",\n"+gap)+"\n"+mind+"]":"["+partial.join(",")+"]";gap=mind;return v}if(rep&&typeof rep==="object"){length=rep.length;for(i=0;i<length;i+=1){if(typeof rep[i]==="string"){k=rep[i];v=str(k,value);if(v){partial.push(quote(k)+(gap?": ":":")+v)}}}}else{for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=str(k,value);if(v){partial.push(quote(k)+(gap?": ":":")+v)}}}}v=partial.length===0?"{}":gap?"{\n"+gap+partial.join(",\n"+gap)+"\n"+mind+"}":"{"+partial.join(",")+"}";gap=mind;return v}}if(typeof JSON.stringify!=="function"){JSON.stringify=function(value,replacer,space){var i;gap="";indent="";if(typeof space==="number"){for(i=0;i<space;i+=1){indent+=" "}}else if(typeof space==="string"){indent=space}rep=replacer;if(replacer&&typeof replacer!=="function"&&(typeof replacer!=="object"||typeof replacer.length!=="number")){throw new Error("JSON.stringify")}return str("",{"":value})}}if(typeof JSON.parse!=="function"){JSON.parse=function(text,reviver){var j;function walk(holder,key){var k;var v;var value=holder[key];if(value&&typeof value==="object"){for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=walk(value,k);if(v!==undefined){value[k]=v}else{delete value[k]}}}}return reviver.call(holder,key,value)}text=String(text);rx_dangerous.lastIndex=0;if(rx_dangerous.test(text)){text=text.replace(rx_dangerous,function(a){return"\\u"+("0000"+a.charCodeAt(0).toString(16)).slice(-4)})}if(rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,""))){j=eval("("+text+")");return typeof reviver==="function"?walk({"":j},""):j}throw new SyntaxError("JSON.parse")}}}());
var SKRIPT_NAME = "Blaetterbarer Katalog Umsetz Script";
var SKRIPT_VERSION = "v2.3.3";
var SKRIPT_DATUM = "2025-10-28";
var CUSTOMER = "personalshop";
var KEEP_CSV_NAMES = false;

var FALLBACK_SUCHPFAD = "X:/_Grafik 2024/10_Katalogseiten/";
var FALLBACK_SUCHTIEFE = 3;

var LOGGING_AKTIVIERT = false;
var CSV_VERSION = "5";

var EXPORT_MODE_DIRECT = "DIRECT";
var EXPORT_MODE_CONVERT = "CONVERT";

// Zuordnung für 4-seitige Quelldokumente (Katalog-Rechts/Links → Quelle 2/3)
var FOUR_PAGE_SIDE_MAP = {
    RIGHT: 2,
    LEFT: 3
};

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
function concatTempWithShapes(myParentObject, temp) {
    var parentType = "" + myParentObject;
    while (parentType.indexOf("Polygon") > 0 || parentType.indexOf("Rectangle") > 0 || parentType.indexOf("Oval") > 0) {
        temp = temp + "," + myParentObject.visibleBounds;
        myParentObject = myParentObject.parent;
        parentType = "" + myParentObject;
    }
    return temp;
}

function concatTempWithPageItems(myParentObject, temp) {
    var parentType = "" + myParentObject;
    while (parentType.indexOf("Polygon") > 0 || parentType.indexOf("Rectangle") > 0 || parentType.indexOf("Oval") > 0 || parentType.indexOf("group") > 0) {
        temp = temp + "," + myParentObject.visibleBounds;
        myParentObject = myParentObject.parent;
        parentType = "" + myParentObject;
    }
    return temp;
}


/**
 * Erstellt ein Export-Item.
 * @param {string} typ - "DIRECT"|"CONVERT"
 * @param {File} inddDatei - Die zu verarbeitende INDD-Datei.
 * @param {string|null} pdfName - Ursprünglicher PDF-Name (optional).
 * @param {Page} page - Page-Objekt aus dem Hauptdokument.
 * @param {Link} link - Original Link-Objekt.
 * @returns {Object} ExportItem.
 */
function erstelleExportItem(typ, inddDatei, pdfName, page, link) {
    return {
        typ: typ,
        inddDatei: inddDatei,
        pdfName: pdfName || null,
        page: page,
        link: link,
        getCacheKey: function() {
            if (this.typ === EXPORT_MODE_DIRECT && this.link) {
                return this.link.name;
            }
            if (this.typ === EXPORT_MODE_CONVERT && this.pdfName) {
                var basisName = (new StringUtils()).entferneExtension(this.pdfName);
                return basisName + ".indd";
            }
            return null;
        }
    };
}

/**
 * Erstellt ein CSV-Daten-Model.
 * @param {Document} inddDoc - Das verarbeitete InDesign-Dokument.
 * @param {Object} sourceItem - Das Quell-Item (ExportItem).
 * @param {string} zeilenDaten - Die gesammelten CSV-Zeilen.
 * @param {string} verwendeterLayer - Der verwendete Layer-Name.
 * @returns {Object} CsvDataModel.
 */
function erstelleCsvDataModel(inddDoc, sourceItem, zeilenDaten, verwendeterLayer) {
    /**
     * Erstellt den CSV-Header aus einem Dokument.
     * @param {Document} doc - Das InDesign-Dokument.
     * @returns {string} CSV-Header-Zeile.
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
            verarbeitetAm: new Date().toUTCString(),
            zeilenAnzahl: zeilenAnzahl,
            dokumentName: inddDoc.name,
            verwendeterLayer: verwendeterLayer || "unbekannt"
        }
    };
}


/**
 * @class StringUtils
 * @description Hilfsfunktionen für String-Manipulation.
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
 * @class DateUtils
 * @description Hilfsfunktionen für Datum und Zeit.
 */
function DateUtils() {
    this.formatiereZeit = function(sekunden) {
        var h = Math.floor(sekunden / 3600);
        var m = Math.floor((sekunden % 3600) / 60);
        var s = Math.floor(sekunden % 60);
        var pad = function(num) {
            return (num < 10 ? '0' : '') + num;
        };
        return pad(h) + ":" + pad(m) + ":" + pad(s);
    };
}

/**
 * @class ErrorHandler
 * @description Fehlerbehandlung.
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
        fehlerListe.push({
            betroffeneDatei: kontext,
            grund: fehlerMsg
        });
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
 * @class Logger
 * @description Logging-System.
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
 * @class FileService
 * @description Dateioperationen und Fallback-Suche.
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
            alleTreffer.sort(function(a, b) {
                return b.modified.getTime() - a.modified.getTime();
            });
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
 * @class DocumentService
 * @description InDesign-Dokumentoperationen.
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
            throw fehler;
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
 * @class CsvService
 * @description Verantwortlich für das Schreiben von CSV-Dateien.
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
            if (z.length === 0) {
                continue;
            }
            if (!seen.hasOwnProperty(z)) {
                seen[z] = true;
                out.push(z);
            }
        }
        return out.join('\n');
    };
}

/**
 * @class CacheService
 * @description Verwaltet das Lesen und Schreiben des JSON-Caches zur Beschleunigung der Verarbeitung.
 */
function CacheService(dependencies) {
    this.logger = dependencies.logger;
    this.CACHE_ORDNER_PFAD = "X:/00 Entwürfe/Miklos/Katalogseiten Info/";
    this.CACHE_DATEI_PREFIX = "blatterbarer_katalog_csv_";

    /**
     * Findet die neueste Cache-Datei im Cache-Ordner.
     * @returns {File|null} Die neueste Cache-Datei oder null, wenn keine gefunden wurde.
     * @private
     */
    this._findeNeuesteCacheDatei = function() {
        var cacheOrdner = new Folder(this.CACHE_ORDNER_PFAD);
        if (!cacheOrdner.exists) {
            this.logger.logWarnung("Cache-Ordner existiert nicht: " + this.CACHE_ORDNER_PFAD);
            return null;
        }

        var cacheDateien = cacheOrdner.getFiles(this.CACHE_DATEI_PREFIX + "*.json");
        if (cacheDateien.length === 0) {
            return null;
        }

        cacheDateien.sort();
        return cacheDateien[cacheDateien.length - 1];
    };

    /**
     * Lädt die Daten aus der neuesten Cache-Datei.
     * @returns {Object} Das geladene Cache-Objekt oder ein leeres Objekt bei Fehlern.
     */
    this.ladeCache = function() {
        var cacheDatei = this._findeNeuesteCacheDatei();
        if (!cacheDatei || !cacheDatei.exists) {
            this.logger.log("Keine Cache-Datei gefunden. Starte mit leerem Cache.");
            return {};
        }

        try {
            this.logger.log("Lade Cache aus: " + cacheDatei.name);
            cacheDatei.open("r");
            var inhalt = cacheDatei.read();
            cacheDatei.close();
            if (inhalt.length > 0) {
                return JSON.parse(inhalt);
            }
            return {};
        } catch (e) {
            this.logger.logFehler("Cache-Datei konnte nicht gelesen oder geparst werden: " + cacheDatei.name + ". Fehler: " + e.message);
            return {};
        }
    };

    /**
     * Speichert die übergebenen Daten im Cache. Erstellt täglich eine neue Datei oder überschreibt die heutige.
     * @param {Object} daten - Das zu speichernde Cache-Objekt.
     */
    this.speichereCache = function(daten) {
        var cacheOrdner = new Folder(this.CACHE_ORDNER_PFAD);
        if (!cacheOrdner.exists) {
            if (!cacheOrdner.create()) {
                this.logger.logFehler("Cache-Ordner konnte nicht erstellt werden: " + this.CACHE_ORDNER_PFAD);
                return;
            }
        }

        var jetzt = new Date();
        var jahr = jetzt.getFullYear();
        var monat = ("0" + (jetzt.getMonth() + 1)).slice(-2);
        var tag = ("0" + jetzt.getDate()).slice(-2);
        var datumsstempel = jahr + "-" + monat + "-" + tag;

        var zielDatei = new File(this.CACHE_ORDNER_PFAD + "/" + this.CACHE_DATEI_PREFIX + datumsstempel + ".json");
        try {
            this.logger.log("Speichere Cache in: " + zielDatei.name);
            zielDatei.encoding = "UTF-8";
            zielDatei.open("w");
            zielDatei.write(JSON.stringify(daten, null, 2));
            zielDatei.close();
        } catch (e) {
            this.logger.logFehler("Cache-Datei konnte nicht geschrieben werden: " + zielDatei.name + ". Fehler: " + e.message);
        }
    };
}


/**
 * @class DataProcessors
 * @description Sammlung von Objekten zur Extraktion von Daten (Graphics, Text, etc.)
 */
function DataProcessors(dependencies) {
    this.logger = dependencies.logger;

    this.verarbeiteGraphics = function(currentPage, page, settings) {
        var temp = "";
        if (!currentPage) return temp;
        for (var j = 0; j < currentPage.allPageItems.length; j++) {
            var pageItem = currentPage.allPageItems[j];
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
                if (page) {
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
                        if (page) {
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
            } catch (e) {}
        }
        return temp;
    };

    /**
     * Ermittelt die Zielseite im geöffneten Quelldokument.
     * @param {Object} sourceItem - Export-Item mit Link und Katalog-Seite.
     * @param {Document} inddDoc - Geöffnetes InDesign-Dokument (Quelle).
     * @returns {number} 1-basierte Seitennummer im Quelldokument.
     */
    this.ermittleZielseite = function(sourceItem, inddDoc) {
        try {
            if (inddDoc.pages.length === 1) return 1;

            try {
                if (sourceItem.typ === EXPORT_MODE_CONVERT &&
                    sourceItem.link && sourceItem.link.parent &&
                    sourceItem.link.parent.hasOwnProperty("pdfAttributes") &&
                    sourceItem.link.parent.pdfAttributes) {
                    var pdfNum = sourceItem.link.parent.pdfAttributes.pageNumber;
                    if (pdfNum >= 1 && pdfNum <= inddDoc.pages.length) {
                        this.logger.log("Zielseite aus PDF.pageNumber: " + pdfNum);
                        return pdfNum;
                    }
                }
            } catch (ePdf) {}

            var linkParent = sourceItem.link ? sourceItem.link.parent : null;
            if (linkParent && linkParent.hasOwnProperty("pageNumber")) {
                var pageNum = linkParent.pageNumber;
                if (pageNum >= 1 && pageNum <= inddDoc.pages.length) {
                    this.logger.log("Zielseite aus ImportedPage.pageNumber: " + pageNum);
                    return pageNum;
                }
            }

            if (sourceItem.typ === EXPORT_MODE_CONVERT && inddDoc.pages.length === 4) {
                var isRight = (sourceItem.page.side === PageSideOptions.RIGHT_HAND);
                var ziel = isRight ? FOUR_PAGE_SIDE_MAP.RIGHT : FOUR_PAGE_SIDE_MAP.LEFT;
                this.logger.log("Zielseite aus Seitenlogik (4p): " + (isRight ? "RIGHT→" : "LEFT→") + ziel);
                return ziel;
            }

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
 * @class DialogFactory
 * @description Erstellt alle UI-Dialoge.
 */
function DialogFactory() {
    this.erstelleCsvExportEinstellungenDialog = function(layerNamen, standardOrdner) {
        var dialog = new Window("dialog", "CSV Export Einstellungen");
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 10;
        dialog.margins = 15;

        var layerPanel = dialog.add("panel", undefined, "Ebene für den Export auswählen");
        layerPanel.alignChildren = "fill";
        var layerOptionen = layerNamen.concat(["** Alle Ebenen **"]);
        var layerDropdown = layerPanel.add("dropdownlist", undefined, layerOptionen);

        var standardAuswahlIndex = 0;
        for (var i = 0; i < layerNamen.length; i++) {
            if (layerNamen[i].indexOf("DA") > -1 && layerNamen[i].indexOf("Seiten") > -1) {
                standardAuswahlIndex = i;
                break;
            }
        }
        layerDropdown.selection = standardAuswahlIndex;
        layerDropdown.preferredSize.width = 350;

        var zielPanel = dialog.add("panel", undefined, "Export-Verzeichnis");
        zielPanel.alignChildren = "fill";
        var zielGroup = zielPanel.add("group", undefined, {
            orientation: "row",
            alignChildren: ["fill", "center"]
        });
        zielGroup.add("statictext", undefined, "Ziel:").preferredSize.width = 40;
        var zielPfadText = zielGroup.add("edittext", undefined, standardOrdner);
        zielPfadText.preferredSize.width = 250;
        var durchsuchenBtn = zielGroup.add("button", undefined, "Durchsuchen...");

        durchsuchenBtn.onClick = function() {
            var folder = Folder.selectDialog("Bitte wähle das Export-Verzeichnis:", Folder(zielPfadText.text));
            if (folder) zielPfadText.text = folder.fsName;
        };

        var cachePanel = dialog.add("panel", undefined, "Optionen");
        cachePanel.alignChildren = "fill";
        var cacheCheckbox = cachePanel.add("checkbox", undefined, "Schnell-Export über Cache verwenden (empfohlen)");
        cacheCheckbox.value = true;

        var buttonGroup = dialog.add("group", undefined, {
            orientation: "row",
            alignment: "right"
        });
        buttonGroup.add("button", undefined, "Abbrechen", {
            name: "cancel"
        });
        buttonGroup.add("button", undefined, "Export starten", {
            name: "ok"
        });

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
                alleEbenen: layerDropdown.selection.text === "** Alle Ebenen **",
                cacheVerwenden: cacheCheckbox.value
            };
        }
        return null;
    };

    this.erstelleAbschnittsOptionenDialog = function() {
        var dialog = new Window("dialog", "Fehler in der Seitenstruktur");
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        var text = "Das Dokument enthält mehrere Seiten mit der Seitenzahl '1' aufgrund von Abschnittsdefinitionen.\n\nSoll das Skript die Seitennummerierung automatisch korrigieren?";
        dialog.add("statictext", [15, 15, 435, 75], text, {
            multiline: true
        });

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        var abbrechenBtn = buttonGroup.add("button", undefined, "Prozess abbrechen");
        var korrigierenBtn = buttonGroup.add("button", undefined, "Automatisch korrigieren", {
            name: "ok"
        });

        abbrechenBtn.onClick = function() {
            dialog.close(0);
        };
        korrigierenBtn.onClick = function() {
            dialog.close(1);
        };

        return dialog.show() === 1;
    };

    this.erstelleFortschrittsDialog = function(nachricht, dateiUtils) {
        var w = new Window("palette", "Fortschritt", undefined, {
            closeButton: false
        });
        w.alignChildren = "fill";
        w.margins = 15;
        var t = w.add("statictext", undefined, nachricht, {
            properties: {
                preferredSize: [450, -1]
            }
        });
        var b = w.add("progressbar", undefined, 0, 100, {
            properties: {
                preferredSize: [450, 20]
            }
        });
        var timeGroup = w.add('group', undefined, {
            orientation: 'row',
            alignment: 'fill'
        });
        var elapsedLabel = timeGroup.add("statictext", undefined, "Verstrichene Zeit: 00:00:00");
        elapsedLabel.alignment = "left";
        var estimateLabel = timeGroup.add("statictext", undefined, "Verbleibende Zeit: Berechnung...");
        estimateLabel.alignment = "right";
        var abgebrochen = false;
        w.add("button", undefined, "Abbrechen", {
            alignment: "center"
        }).onClick = function() {
            abgebrochen = true;
            w.close();
        };
        var startTime = new Date();
        w.show();
        return {
            isCancelled: function() {
                return abgebrochen;
            },
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
            close: function() {
                w.close();
            }
        };
    };
}

/**
 * @class ExportCoordinator
 * @description Koordiniert den gesamten Export-Prozess für eine gegebene Liste von Items.
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
        if (exportItems.length === 0) {
            return {
                daten: [],
                fehler: []
            };
        }
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
                var dateiName = item.getCacheKey() || "unbekannt";
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
        return {
            daten: alleDaten,
            fehler: fehlerListe
        };
    };

    this._verarbeiteItem = function(item, settings) {
        if (!item.inddDatei || !item.inddDatei.exists) {
            throw new Error("INDD-Datei für Item nicht gefunden: " + item.getCacheKey());
        }

        var doc = null;
        try {
            doc = this.documentService.oeffneDokument(item.inddDatei, false);
            return this.dataProcessors.verarbeiteDokument(doc, settings, item);

        } catch (e) {
            var isLockedError = (e.message.toLowerCase().indexOf("cannot open") > -1) ||
                (e.message.toLowerCase().indexOf("kann die datei nicht öffnen") > -1) ||
                (e.message.toLowerCase().indexOf("is already open") > -1);

            if (isLockedError) {
                this.logger.logWarnung("Datei '" + item.inddDatei.name + "' ist gesperrt. Erstelle temporäre Kopie in 'Ablage'.");
                return this._verarbeiteGesperrteDatei(item.inddDatei, settings, item);
            } else {
                throw e;
            }
        } finally {
            if (doc) {
                this.documentService.schliesseDokument(doc, SaveOptions.NO);
            }
        }
    };

    this.ermittleInddDateiFuerItem = function(item) {
        if (item.typ === EXPORT_MODE_DIRECT) {
            return File(item.link.filePath);
        }
        if (item.typ === EXPORT_MODE_CONVERT) {
            var basisName = (new StringUtils()).entferneExtension(item.pdfName);
            var inddPfadOriginal = File(item.link.filePath).path + "/" + basisName + ".indd";
            return this.fileService.findeDateiMitFallback(inddPfadOriginal);
        }
        return null;
    };


    this._verarbeiteGesperrteDatei = function(originalDatei, settings, item) {
        var tempDoc = null;
        var tempDatei = null;
        try {
            var parentFolder = originalDatei.parent;
            var ablageFolder = new Folder(parentFolder.fsName + "/Ablage");
            if (!ablageFolder.exists) {
                if (!ablageFolder.create()) {
                    throw new Error("Konnte 'Ablage'-Ordner nicht erstellen: " + ablageFolder.fsName);
                }
            }

            tempDatei = new File(ablageFolder.fsName + "/" + originalDatei.name);
            if (tempDatei.exists) {
                tempDatei.remove();
            }

            if (!originalDatei.copy(tempDatei)) {
                throw new Error("Konnte temporäre Kopie nicht erstellen: " + tempDatei.fsName);
            }
            this.logger.log("Temporäre Kopie erstellt: " + tempDatei.name);

            tempDoc = this.documentService.oeffneDokument(tempDatei, false);
            return this.dataProcessors.verarbeiteDokument(tempDoc, settings, item);

        } finally {
            if (tempDoc) {
                this.documentService.schliesseDokument(tempDoc, SaveOptions.NO);
            }
            if (tempDatei && tempDatei.exists) {
                this.logger.log("Lösche temporäre Kopie: " + tempDatei.name);
                tempDatei.remove();
            }
        }
    };
}


/**
 * @class CsvExportWorkflow
 * @description Hauptlogik für den CSV-Export, inklusive Cache-Verwaltung.
 */
function CsvExportWorkflow(dependencies) {
    this.exportCoordinator = dependencies.exportCoordinator;
    this.csvService = dependencies.csvService;
    this.cacheService = dependencies.cacheService;
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

        var einstellungen = this._holeBenutzerEinstellungen(doc, layerNamen);
        if (!einstellungen) {
            this.logger.log("Benutzer hat den Einstellungsdialog abgebrochen.");
            return;
        }

        var linksAufEbene = this._bestimmeLinksInDokument(doc, einstellungen.alleEbenen ? null : einstellungen.layerName);
        if (linksAufEbene.length === 0) {
            alert("Keine verarbeitbaren INDD- oder PDF-Verknüpfungen auf der gewählten Ebene(n) gefunden.", SKRIPT_NAME);
            return;
        }

        var exportItems = this._erstelleExportItemsAusLinks(linksAufEbene);

        var settings = this._erstelleStandardSettings();
        settings.layer.name = einstellungen.alleEbenen ? "Alle Ebenen" : einstellungen.layerName;

        var alleCsvDaten;
        var fehlerListe = [];

        if (einstellungen.cacheVerwenden) {
            var cacheErgebnis = this._verarbeiteMitCache(exportItems, settings);
            alleCsvDaten = cacheErgebnis.daten;
            fehlerListe = cacheErgebnis.fehler;
        } else {
            this.logger.log("Cache-Verwendung deaktiviert. Führe vollständigen Scan durch.");
            for (var i = 0; i < exportItems.length; i++) {
                exportItems[i].inddDatei = this.exportCoordinator.ermittleInddDateiFuerItem(exportItems[i]);
            }
            var vollscanErgebnis = this.exportCoordinator.koordiniereExport(exportItems, settings);
            alleCsvDaten = vollscanErgebnis.daten;
            fehlerListe = vollscanErgebnis.fehler;
        }

        if (alleCsvDaten.length > 0) {
            try {
                for (var j = 0; j < alleCsvDaten.length; j++) {
                    var seitenNummer = alleCsvDaten[j].elternSeite;
                    if (seitenNummer > -1 && this.pages.join(",").indexOf(seitenNummer) === -1) {
                        this.pages.push(seitenNummer);
                    }
                }
                var csvDatei = this._erstelleCsvDateiNamen(einstellungen.zielOrdner, doc);
                this.csvService.schreibeEinzelneCsv(alleCsvDaten, csvDatei);
                this._zeigeErgebnis(alleCsvDaten.length, fehlerListe, csvDatei.fsName);
            } catch (e) {
                alert("Fehler beim Speichern der CSV-Datei:\n" + e.message);
                this.logger.logFehler("Fehler beim Speichern der CSV: " + e.message);
            }
        } else {
            this._zeigeErgebnis(0, fehlerListe, null);
        }
    };

    this._holeBenutzerEinstellungen = function(doc, layerNamen) {
        var einstellungen = null;
        var prozessFortsetzen = false;
        while (!prozessFortsetzen) {
            einstellungen = this.dialogFactory.erstelleCsvExportEinstellungenDialog(layerNamen, doc.filePath);
            if (!einstellungen) return null;

            if (einstellungen.alleEbenen) {
                var warnungText = "Achtung!\n\nDie Verarbeitung aller Ebenen kann den Prozess erheblich verlängern.\n\nMöchtest Du fortfahren?";
                if (confirm(warnungText, false, "Lange Prozesszeit")) {
                    prozessFortsetzen = true;
                }
            } else {
                prozessFortsetzen = true;
            }
        }
        return einstellungen;
    };

    /**
     * Verarbeitet Export-Items mit Cache-Unterstützung zur Beschleunigung.
     * 
     * Diese Methode implementiert die intelligente Cache-Logik:
     * - Zeigt einen Fortschrittsdialog während der Cache-Prüfung an
     * - Lädt den vorhandenen Cache und vergleicht Änderungsdaten
     * - Verwendet gecachte Daten für unveränderte Dateien
     * - Verarbeitet nur neue oder geänderte Dateien
     * - Speichert neue Ergebnisse im Cache mit dem ursprünglichen Link-Namen als Schlüssel
     * - Unterstützt Benutzerabbruch während der Cache-Prüfung
     * 
     * @param {Array} exportItems - Array von Export-Items (erstellt aus Links)
     * @param {Object} settings - Export-Einstellungen
     * @returns {Object} Objekt mit {daten: Array, fehler: Array}
     * @private
     */
    this._verarbeiteMitCache = function(exportItems, settings) {
        this.logger.log("Verwende Cache-Logik für die Verarbeitung.");
        var cacheDaten = this.cacheService.ladeCache();
        var zuVerarbeitendeItems = [];
        var finaleCsvDaten = [];
        
        var cacheCheckDialog = this.dialogFactory.erstelleFortschrittsDialog(
            "Cache-Prüfung läuft...", 
            this.exportCoordinator.dateUtils
        );

        try {
            for (var i = 0; i < exportItems.length; i++) {
                if (cacheCheckDialog.isCancelled()) {
                    this.logger.log("Cache-Prüfung wurde vom Benutzer abgebrochen");
                    return {
                        daten: [],
                        fehler: []
                    };
                }

                var item = exportItems[i];
                var tempCacheKey = item.getCacheKey();
                var anzeigeNachricht = "Prüfe Datei " + (i + 1) + " von " + exportItems.length + ": " + (tempCacheKey || "unbekannt");
                cacheCheckDialog.update(i + 1, exportItems.length, anzeigeNachricht);

                var cacheKey = item.getCacheKey();
                var cacheEintrag = cacheDaten[cacheKey];
                
                if (cacheEintrag && cacheEintrag.lastKnownPath) {
                    var cachedFile = new File(cacheEintrag.lastKnownPath);
                    if (cachedFile.exists) {
                        item.inddDatei = cachedFile;
                        this.logger.log("Datei über Cache-Pfad gefunden: " + cacheEintrag.lastKnownPath);
                    } else {
                        item.inddDatei = this.exportCoordinator.ermittleInddDateiFuerItem(item);
                        this.logger.log("Cache-Pfad ungültig, Fallback-Suche durchgeführt für: " + cacheKey);
                    }
                } else {
                    item.inddDatei = this.exportCoordinator.ermittleInddDateiFuerItem(item);
                }
                
                var inddDatei = item.inddDatei;

                if (!cacheKey || !inddDatei || !inddDatei.exists) {
                    this.logger.logWarnung("Quelldatei für Item nicht gefunden, überspringe: " + (cacheKey || "unbekannt"));
                    continue;
                }

                var dateiModifiziert = new Date(inddDatei.modified);

                if (cacheEintrag && new Date(cacheEintrag.lastModified) >= dateiModifiziert) {
                    this.logger.log("'" + cacheKey + "' aus dem Cache geladen (unverändert).");
                    finaleCsvDaten.push(cacheEintrag.csvDatenModel);
                } else {
                    var grund = cacheEintrag ? "wurde geändert" : "ist neu";
                    this.logger.log("'" + cacheKey + "' wird verarbeitet (" + grund + ").");
                    zuVerarbeitendeItems.push(item);
                }
            }
        } finally {
            cacheCheckDialog.close();
        }

        var prozessErgebnis = this.exportCoordinator.koordiniereExport(zuVerarbeitendeItems, settings);

        // Die neuen Ergebnisse in das Cache-Objekt eintragen
        for (var k = 0; k < prozessErgebnis.daten.length; k++) {
            var neuesModell = prozessErgebnis.daten[k];
            var quellInddName = neuesModell.quellIndd;

            // Finde das zugehörige Export-Item anhand des tatsächlichen Dateinamens
            var zugehoerigesItem = null;
            for (var m = 0; m < zuVerarbeitendeItems.length; m++) {
                if (zuVerarbeitendeItems[m].inddDatei.name === quellInddName) {
                    zugehoerigesItem = zuVerarbeitendeItems[m];
                    break;
                }
            }

            if (zugehoerigesItem) {
                // Verwende den ursprünglichen Cache-Schlüssel (Link-Name) für konsistente Speicherung
                var cacheSchluessel = zugehoerigesItem.getCacheKey();
                cacheDaten[cacheSchluessel] = {
                    lastModified: new Date(zugehoerigesItem.inddDatei.modified).toUTCString(),
                    lastKnownPath: zugehoerigesItem.inddDatei.fsName,
                    csvDatenModel: neuesModell
                };
                this.logger.log("Cache-Eintrag erstellt: '" + cacheSchluessel + "' (Quelldatei: " + quellInddName + ")");
                finaleCsvDaten.push(neuesModell);
            } else {
                this.logger.logWarnung("Kein zugehöriges Export-Item für Quelldatei gefunden: " + quellInddName);
            }
        }

        this.cacheService.speichereCache(cacheDaten);

        return {
            daten: finaleCsvDaten,
            fehler: prozessErgebnis.fehler
        };
    };

    this._bestimmeLinksInDokument = function(doc, layerName) {
        var linksImDokument = [];
        var zielLayer = layerName ? doc.layers.itemByName(layerName) : null;

        for (var p = 0; p < doc.pages.length; p++) {
            var currentPage = doc.pages[p];
            for (var k = 0; k < currentPage.allPageItems.length; k++) {
                var currentPageItem = currentPage.allPageItems[k];
                var ebeneOk = (!zielLayer) || (currentPageItem.itemLayer === zielLayer);

                if (currentPageItem instanceof Rectangle && currentPageItem.hasOwnProperty("importedPages")) {
                    for (var j = 0; j < currentPageItem.importedPages.length; j++) {
                        var importedPage = currentPageItem.importedPages[j];
                        try {
                            var link = importedPage.itemLink;
                            if (!link) continue;
                            var nameLower = ("" + link.name).toLowerCase();
                            if (nameLower.indexOf(".indd") === -1 && nameLower.indexOf(".pdf") === -1) continue;
                            if (!zielLayer || importedPage.itemLayer === zielLayer) {
                                linksImDokument.push({
                                    itemLink: link,
                                    page: currentPage
                                });
                            }
                        } catch (e) {}
                    }
                }

                if (currentPageItem.allGraphics && currentPageItem.allGraphics.length > 0) {
                    for (var g = 0; g < currentPageItem.allGraphics.length; g++) {
                        try {
                            var graphic = currentPageItem.allGraphics[g];
                            var gLink = graphic && graphic.itemLink ? graphic.itemLink : null;
                            if (!gLink) continue;
                            var gNameLower = ("" + gLink.name).toLowerCase();
                            if (gNameLower.indexOf(".indd") === -1 && gNameLower.indexOf(".pdf") === -1) continue;
                            if (ebeneOk) {
                                linksImDokument.push({
                                    itemLink: gLink,
                                    page: currentPage
                                });
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
            var pdfName = (typ === EXPORT_MODE_CONVERT) ? linkInfo.itemLink.name : null;

            items.push(erstelleExportItem(typ, null, pdfName, linkInfo.page, linkInfo.itemLink));
        }
        return items;
    };

    this._erstelleStandardSettings = function() {
        return {
            layer: {
                name: ""
            },
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
            pages: []
        };
    };

    this._erstelleCsvDateiNamen = function(ordner, doc) {
        this.pages.sort(function(a, b) {
            return parseInt(a) - parseInt(b);
        });

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
 * @class CsvExportManager
 * @description Haupt-Controller für den Standalone-Export.
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
        for (var i = 1; i < doc.pages.length; i++) {
            var page = doc.pages[i];
            if (page.name === "1") {
                try {
                    page.appliedSection.continueNumbering = true;
                    this.logger.log("Abschnittsanfang auf Dokumentseite " + (i + 1) + " (ursprünglich als '" + page.name + "' benannt) wurde korrigiert.");
                } catch (e) {
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
     * Erstellt den Dependency Container.
     * @returns {Object} Der Dependency Container.
     */
    function erstelleDependencyContainer() {
        var container = {};

        container.errorHandler = new ErrorHandler();
        container.logger = new Logger();
        container.dateUtils = new DateUtils();
        container.stringUtils = new StringUtils();
        container.cacheService = new CacheService({
            logger: container.logger
        });

        container.fileService = new FileService({
            logger: container.logger
        });
        container.documentService = new DocumentService({
            logger: container.logger,
            errorHandler: container.errorHandler
        });
        container.csvService = new CsvService({
            logger: container.logger
        });

        container.dataProcessors = new DataProcessors({
            logger: container.logger
        });
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
            cacheService: container.cacheService,
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