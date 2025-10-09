/**
 * Globale Konstanten für das INDD/PDF Link Manager & CSV Exporter System
 *
 * Diese Datei enthält alle grundlegenden Konfigurationswerte,
 * die im gesamten System verwendet werden.
 */

// Skript-Informationen
var SKRIPT_NAME = "INDD/PDF Link Manager & CSV Exporter";
var SKRIPT_VERSION = "v2.0.0";
var SKRIPT_DATUM = "2025-01-15";

// Fallback-Dateisuche
var FALLBACK_SUCHPFAD = "X:/_Grafik 2024/10_Katalogseiten/";
var FALLBACK_SUCHTIEFE = 3;

// Validierung
var DATUM_TOLERANZ_MINUTEN = 20;

// Logging
var LOGGING_AKTIVIERT = true;

// CSV-Format
var CSV_VERSION = "5.0";

// Export-Modi
var EXPORT_MODE_DIRECT = "DIRECT";
var EXPORT_MODE_CONVERT = "CONVERT";
var EXPORT_MODE_MIXED = "MIXED";
var EXPORT_MODE_ERROR = "ERROR";

// CSV-Zeilen-Typen
var CSV_ZEILE_GRAPHIC = "G";
var CSV_ZEILE_TEXT = "T";
var CSV_ZEILE_PAGEITEM = "W";

// Performance-Einstellungen
var PROGRESS_UPDATE_INTERVALL = 10;
var MAX_REKURSIONS_TIEFE = 3;
/**
 * Standard-Einstellungen für Export-Optionen
 *
 * Diese Werte können durch Benutzer-Dialoge überschrieben werden.
 */

// Standard Export-Optionen
var STANDARD_EXPORT_GRAPHICS = true;
var STANDARD_EXPORT_TEXTFRAMES = true;
var STANDARD_EXPORT_TEXT_IN_STORIES = true;
var STANDARD_EXPORT_TABLES = false;
var STANDARD_EXPORT_PAGEITEMS = false;

// Standard CSV-Format
var STANDARD_CSV_EINZELN = true;
var STANDARD_CSV_MEHRFACH = false;

// Standard Layer-Auswahl
var STANDARD_LAYER_AUTO = true;
var STANDARD_LAYER_MANUAL = false;
var STANDARD_LAYER_NAME = "";

// Standard Source-Info
var STANDARD_SOURCE_INFO = true;

// Regex-Patterns für Artikelnummern
var REGEX_TABLE_IN_STORY = /\d{6}/;
var REGEX_TEXTFRAME_IN_STORY = /(\d{2}[.])?\d{3}[.]\d{3}/;

// Customer-spezifische Einstellungen (falls benötigt)
var CUSTOMER_NAME = "default";
var CUSTOMER_PREFIX = "";
/**
 * LinkModel - Datenstruktur für Link-Informationen
 *
 * Kapselt alle relevanten Informationen über einen InDesign-Link.
 */

/**
 * Erstellt ein Link-Model aus einem InDesign Link-Objekt
 *
 * @param {Link} link - Das InDesign Link-Objekt
 * @returns {Object} LinkModel mit allen relevanten Informationen
 *
 * @typedef {Object} LinkModel
 * @property {string} typ - "INDD"|"PDF"|"OTHER"
 * @property {Link} link - InDesign Link-Objekt
 * @property {string} dateiPfad - Vollständiger Pfad
 * @property {string} dateiName - Nur Dateiname
 * @property {boolean} existiert - Datei existiert?
 * @property {number} elternSeite - Seitennummer im Hauptdokument
 * @property {string|null} fehler - Fehlermeldung (falls vorhanden)
 * @property {Object} metadaten - Zusätzliche Informationen
 */
function erstelleLinkModel(link) {
    var datei = File(link.filePath);
    var typ = "OTHER";

    var linkNameLower = link.name.toLowerCase();
    if (linkNameLower.indexOf(".indd") > -1) {
        typ = "INDD";
    } else if (linkNameLower.indexOf(".pdf") > -1) {
        typ = "PDF";
    }

    var elternSeite = -1;
    try {
        if (link.parent && link.parent.parentPage) {
            var seitenName = link.parent.parentPage.name;
            var seitenNummer = parseInt(seitenName, 10);
            if (!isNaN(seitenNummer)) {
                elternSeite = seitenNummer;
            }
        }
    } catch (e) {
        // Fehler beim Ermitteln der Elternseite ignorieren
    }

    return {
        typ: typ,
        link: link,
        dateiPfad: datei.fsName,
        dateiName: datei.name,
        existiert: datei.exists,
        elternSeite: elternSeite,
        fehler: null,
        metadaten: {
            status: link.status,
            groesse: datei.exists ? datei.length : 0,
            geaendert: datei.exists ? datei.modified : null
        }
    };
}
/**
 * ExportPlanModel - Datenstruktur für Export-Planung
 *
 * Enthält alle Informationen über den geplanten Export-Vorgang.
 */

/**
 * Erstellt einen Export-Plan
 *
 * @param {string} modus - "DIRECT"|"CONVERT"|"MIXED"
 * @param {Array} items - Array von ExportItem-Objekten
 * @returns {Object} ExportPlanModel
 *
 * @typedef {Object} ExportPlanModel
 * @property {string} modus - "DIRECT"|"CONVERT"|"MIXED"
 * @property {Array<ExportItem>} items - Zu verarbeitende Items
 * @property {boolean} abgebrochen - User hat abgebrochen
 * @property {number} gesamtAnzahl - Gesamtzahl Items
 * @property {number} gueltigeAnzahl - Anzahl gültiger Items
 * @property {number} fehlendeAnzahl - Anzahl fehlender Items
 * @property {Object} statistik - Weitere Statistiken
 *
 * @typedef {Object} ExportItem
 * @property {string} typ - "DIRECT"|"CONVERT"
 * @property {File} inddDatei - INDD-Datei Objekt
 * @property {string|null} pdfName - Ursprünglicher PDF-Name (bei CONVERT)
 * @property {number} elternSeite - Seitennummer im Hauptdokument
 * @property {Link} link - Original Link-Objekt
 */
function erstelleExportPlan(modus, items) {
    var gueltig = 0;
    var fehlend = 0;
    var direkteINDDs = 0;
    var konvertierteINDDs = 0;

    for (var i = 0; i < items.length; i++) {
        if (items[i].inddDatei && items[i].inddDatei.exists) {
            gueltig++;
        } else {
            fehlend++;
        }

        if (items[i].typ === EXPORT_MODE_DIRECT) {
            direkteINDDs++;
        } else if (items[i].typ === EXPORT_MODE_CONVERT) {
            konvertierteINDDs++;
        }
    }

    return {
        modus: modus,
        items: items,
        abgebrochen: false,
        gesamtAnzahl: items.length,
        gueltigeAnzahl: gueltig,
        fehlendeAnzahl: fehlend,
        statistik: {
            direkteINDDs: direkteINDDs,
            konvertierteINDDs: konvertierteINDDs,
            erstelltAm: new Date()
        }
    };
}

/**
 * Erstellt ein einzelnes Export-Item
 *
 * @param {string} typ - "DIRECT"|"CONVERT"
 * @param {File} inddDatei - Die zu verarbeitende INDD-Datei
 * @param {string|null} pdfName - Ursprünglicher PDF-Name (optional)
 * @param {number} elternSeite - Seitennummer im Hauptdokument
 * @param {Link} link - Original Link-Objekt
 * @returns {Object} ExportItem
 */
function erstelleExportItem(typ, inddDatei, pdfName, elternSeite, link) {
    return {
        typ: typ,
        inddDatei: inddDatei,
        pdfName: pdfName || null,
        elternSeite: elternSeite,
        link: link
    };
}
/**
 * CsvDataModel - Datenstruktur für CSV-Export-Daten
 *
 * Enthält alle gesammelten Daten aus einem INDD-Dokument für den CSV-Export.
 */

/**
 * Erstellt ein CSV-Daten-Model
 *
 * @param {Document} inddDoc - Das verarbeitete InDesign-Dokument
 * @param {Object} sourceItem - Das Quell-Item (ExportItem)
 * @param {string} zeilenDaten - Die gesammelten CSV-Zeilen
 * @param {string} verwendeterLayer - Der verwendete Layer-Name
 * @returns {Object} CsvDataModel
 *
 * @typedef {Object} CsvDataModel
 * @property {string} quellIndd - Quell-INDD Dateiname
 * @property {string|null} quellPdf - Quell-PDF Dateiname (bei CONVERT)
 * @property {number} elternSeite - Seite im Hauptdokument
 * @property {string} header - CSV-Header (zeroPoint, width, height, version)
 * @property {string} zeilen - CSV-Datenzeilen
 * @property {Object} metadaten - Verarbeitungsmetadaten
 */
function erstelleCsvDataModel(inddDoc, sourceItem, zeilenDaten, verwendeterLayer) {
    var header = erstelleCsvHeader(inddDoc);

    var zeilenAnzahl = 0;
    if (zeilenDaten && zeilenDaten.length > 0) {
        var zeilenArray = zeilenDaten.split('\n');
        zeilenAnzahl = zeilenArray.length;
    }

    return {
        quellIndd: sourceItem.inddDatei.name,
        quellPdf: sourceItem.pdfName || null,
        elternSeite: sourceItem.elternSeite || -1,
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
 * Erstellt den CSV-Header aus einem Dokument
 *
 * @param {Document} doc - Das InDesign-Dokument
 * @returns {string} CSV-Header-Zeile
 */
function erstelleCsvHeader(doc) {
    var zeroPoint = doc.zeroPoint;
    var pageWidth = doc.documentPreferences.pageWidth;
    var pageHeight = doc.documentPreferences.pageHeight;

    return zeroPoint + "," + pageWidth + "," + pageHeight + "," + CSV_VERSION + "\n";
}
/**
 * StringUtils - String-Operationen
 *
 * Hilfsfunktionen für String-Manipulation.
 */

function StringUtils() {
    /**
     * Kodiert einen Dateinamen für CSV (entfernt Kommas)
     *
     * @param {string} dateiName - Der zu kodierende Dateiname
     * @returns {string} Kodierter Dateiname
     */
    this.kodiereDatainame = function(dateiName) {
        if (!dateiName) return "";
        var kodiert = File.encode(dateiName);
        kodiert = kodiert.replace(/,/g, "_");
        return kodiert;
    };

    /**
     * Entfernt die Extension von einem Dateinamen
     *
     * @param {string} dateiName - Der Dateiname
     * @returns {string} Dateiname ohne Extension
     */
    this.entferneExtension = function(dateiName) {
        if (!dateiName) return "";
        var letzterpunkt = dateiName.lastIndexOf(".");
        if (letzterpunkt === -1) return dateiName;
        return dateiName.substring(0, letzterpunkt);
    };

    /**
     * Ersetzt die Extension eines Dateinamens
     *
     * @param {string} dateiName - Der Dateiname
     * @param {string} neueExt - Die neue Extension (mit Punkt)
     * @returns {string} Dateiname mit neuer Extension
     */
    this.ersetzeExtension = function(dateiName, neueExt) {
        var basis = this.entferneExtension(dateiName);
        return basis + neueExt;
    };

    /**
     * Entfernt Whitespace am Anfang und Ende
     *
     * @param {string} text - Der zu trimmende Text
     * @returns {string} Getrimmter Text
     */
    this.trimme = function(text) {
        if (!text) return "";
        return text.replace(/^\s+|\s+$/g, '');
    };

    /**
     * Füllt einen String links mit Zeichen auf
     *
     * @param {string} text - Der Text
     * @param {number} laenge - Die Ziel-Länge
     * @param {string} zeichen - Das Füllzeichen
     * @returns {string} Aufgefüllter Text
     */
    this.fuelleLinks = function(text, laenge, zeichen) {
        text = String(text);
        while (text.length < laenge) {
            text = zeichen + text;
        }
        return text;
    };
}
/**
 * DateUtils - Datums- und Zeit-Operationen
 *
 * Hilfsfunktionen für Datum und Zeit.
 */

function DateUtils() {
    /**
     * Formatiert einen Zeitstempel (YYYYMMDD-HHMMSS)
     *
     * @param {Date} datum - Das Datum
     * @returns {string} Formatierter Zeitstempel
     */
    this.formatiereZeitstempel = function(datum) {
        if (!datum) datum = new Date();

        var jahr = datum.getFullYear();
        var monat = this.fuelleNull(datum.getMonth() + 1);
        var tag = this.fuelleNull(datum.getDate());
        var stunden = this.fuelleNull(datum.getHours());
        var minuten = this.fuelleNull(datum.getMinutes());
        var sekunden = this.fuelleNull(datum.getSeconds());

        return jahr + monat + tag + "-" + stunden + minuten + sekunden;
    };

    /**
     * Formatiert Sekunden als Zeit (HH:MM:SS)
     *
     * @param {number} sekunden - Anzahl der Sekunden
     * @returns {string} Formatierte Zeit
     */
    this.formatiereZeit = function(sekunden) {
        var h = Math.floor(sekunden / 3600);
        var m = Math.floor((sekunden % 3600) / 60);
        var s = Math.floor(sekunden % 60);

        h = this.fuelleNull(h);
        m = this.fuelleNull(m);
        s = this.fuelleNull(s);

        return h + ":" + m + ":" + s;
    };

    /**
     * Hilfsfunktion zum Auffüllen mit Nullen
     *
     * @param {number} zahl - Die Zahl
     * @returns {string} Mit Null aufgefüllte Zahl
     */
    this.fuelleNull = function(zahl) {
        return (zahl < 10 ? "0" : "") + zahl;
    };

    /**
     * Gibt die aktuelle Zeit als String zurück (HH:MM:SS)
     *
     * @returns {string} Aktuelle Zeit
     */
    this.holeAktuelleZeit = function() {
        var jetzt = new Date();
        var zeitString = jetzt.toTimeString();
        return zeitString.substr(0, 8);
    };
}
/**
 * RegexHelper - Regex-Operationen
 *
 * Hilfsfunktionen für Regular Expressions.
 */

function RegexHelper() {
    /**
     * Validiert einen Regex-String
     *
     * @param {string} regexString - Der Regex-String
     * @returns {boolean} True wenn gültig
     */
    this.validiereRegex = function(regexString) {
        try {
            new RegExp(regexString);
            return true;
        } catch (e) {
            return false;
        }
    };

    /**
     * Erstellt ein RegExp-Objekt aus einem String
     *
     * @param {string} regexString - Der Regex-String
     * @returns {RegExp|null} RegExp-Objekt oder null bei Fehler
     */
    this.erstelleRegex = function(regexString) {
        try {
            return new RegExp(regexString);
        } catch (e) {
            return null;
        }
    };

    /**
     * Testet ob ein Text auf einen Regex passt
     *
     * @param {string} inhalt - Der zu testende Text
     * @param {RegExp} regex - Das Regex-Objekt
     * @returns {boolean} True wenn Match gefunden
     */
    this.testeMatch = function(inhalt, regex) {
        if (!inhalt || !regex) return false;
        var result = regex.exec(inhalt);
        return result !== null && result !== undefined;
    };

    /**
     * Extrahiert den ersten Match
     *
     * @param {string} inhalt - Der Text
     * @param {RegExp} regex - Das Regex-Objekt
     * @returns {string|null} Der Match oder null
     */
    this.extrahiereMatch = function(inhalt, regex) {
        if (!inhalt || !regex) return null;
        var result = regex.exec(inhalt);
        if (result === null || result === undefined) return null;
        return result[0];
    };
}
/**
 * ErrorHandler - Fehlerbehandlung
 *
 * Formatiert und sammelt Fehler.
 */

function ErrorHandler() {
    /**
     * Formatiert ein Fehler-Objekt
     *
     * @param {Error} fehlerObj - Das Fehler-Objekt
     * @param {string} dateiInfo - Kontextinformation
     * @returns {string} Formatierte Fehlermeldung
     */
    this.formatiereFehler = function(fehlerObj, dateiInfo) {
        var basisNachricht = "Unbekannter Fehler.";

        if (fehlerObj.message && fehlerObj.message.length > 0) {
            basisNachricht = fehlerObj.message;
        } else {
            basisNachricht = "Ein interner Anwendungsfehler ist aufgetreten (keine Detailbeschreibung). " +
                "Dies kann passieren, wenn eine Datei (z.B. '" + (dateiInfo || 'unbekannt') + "') " +
                "beim Öffnen/Platzieren eine Dialogbox (z.B. 'Fehlende Schriftarten') auslösen würde.";
        }

        var zusatz = "";
        if (fehlerObj.line) {
            zusatz += " (Zeile: " + fehlerObj.line;
        }
        if (fehlerObj.number) {
            if (zusatz) {
                zusatz += ", Fehlercode: " + fehlerObj.number + ")";
            } else {
                zusatz += " (Fehlercode: " + fehlerObj.number + ")";
            }
        } else if (zusatz) {
            zusatz += ")";
        }

        return basisNachricht + zusatz;
    };

    /**
     * Sammelt einen Fehler in einer Liste
     *
     * @param {Error} fehler - Das Fehler-Objekt
     * @param {string} kontext - Kontextinformation
     * @param {Array} fehlerListe - Die Fehler-Liste
     */
    this.sammleFehler = function(fehler, kontext, fehlerListe) {
        var fehlerMsg = this.formatiereFehler(fehler, kontext);
        fehlerListe.push({
            betroffeneDatei: kontext,
            grund: fehlerMsg
        });
    };

    /**
     * Erstellt einen Fehler-Report
     *
     * @param {Array} fehlerListe - Liste der Fehler
     * @returns {string} Formatierter Report
     */
    this.erstelleFehlerReport = function(fehlerListe) {
        if (!fehlerListe || fehlerListe.length === 0) {
            return "Keine Fehler aufgetreten.";
        }

        var report = "Bei den folgenden Dateien ist ein Problem aufgetreten:\n\n";
        for (var i = 0; i < fehlerListe.length; i++) {
            report += "- " + fehlerListe[i].betroffeneDatei + "\n";
            report += "  Grund: " + fehlerListe[i].grund + "\n\n";
        }

        return report;
    };
}
/**
 * Logger - Logging-System
 *
 * Schreibt Log-Einträge in eine Datei.
 */

function Logger() {
    this.logDatei = null;
    this.aktiviert = false;

    /**
     * Initialisiert den Logger
     *
     * @param {File} logDatei - Die Log-Datei
     * @param {boolean} aktiviert - Logging aktiviert?
     */
    this.initialisiere = function(logDatei, aktiviert) {
        this.logDatei = logDatei;
        this.aktiviert = aktiviert;

        if (this.aktiviert && this.logDatei) {
            this.logDatei.encoding = "UTF-8";
        }
    };

    /**
     * Schreibt eine Log-Nachricht
     *
     * @param {string} nachricht - Die Nachricht
     */
    this.log = function(nachricht) {
        if (!this.aktiviert || !this.logDatei) return;

        try {
            this.logDatei.open("a");
            var zeitstempel = new Date().toTimeString().substr(0, 8);
            this.logDatei.writeln(zeitstempel + " - [INFO] " + nachricht);
            this.logDatei.close();
        } catch (e) {
            // Fehler beim Schreiben ignorieren
        }
    };

    /**
     * Schreibt eine Fehler-Nachricht
     *
     * @param {string} fehler - Die Fehlermeldung
     * @param {string} kontext - Der Kontext
     */
    this.logFehler = function(fehler, kontext) {
        if (!this.aktiviert || !this.logDatei) return;

        try {
            this.logDatei.open("a");
            var zeitstempel = new Date().toTimeString().substr(0, 8);
            var nachricht = zeitstempel + " - [ERROR] " + (kontext ? kontext + ": " : "") + fehler;
            this.logDatei.writeln(nachricht);
            this.logDatei.close();
        } catch (e) {
            // Fehler beim Schreiben ignorieren
        }
    };

    /**
     * Schreibt eine Warnung
     *
     * @param {string} nachricht - Die Warnung
     */
    this.logWarnung = function(nachricht) {
        if (!this.aktiviert || !this.logDatei) return;

        try {
            this.logDatei.open("a");
            var zeitstempel = new Date().toTimeString().substr(0, 8);
            this.logDatei.writeln(zeitstempel + " - [WARN] " + nachricht);
            this.logDatei.close();
        } catch (e) {
            // Fehler beim Schreiben ignorieren
        }
    };

    /**
     * Schließt den Logger
     */
    this.schliesse = function() {
        if (this.logDatei) {
            try {
                this.logDatei.close();
            } catch (e) {
                // Ignorieren
            }
        }
    };
}
/**
 * FileService - Dateioperationen und Fallback-Suche
 *
 * Verantwortlich für alle Datei-bezogenen Operationen.
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {Logger} dependencies.logger - Logger-Instanz
 */
function FileService(dependencies) {
    this.logger = dependencies.logger;

    /**
     * Sucht eine Datei mit Fallback-Mechanismus
     *
     * @param {string} originalPfad - Ursprünglicher Dateipfad
     * @returns {File|null} Gefundene Datei oder null
     */
    this.findeDateiMitFallback = function(originalPfad) {
        var datei = File(originalPfad);
        if (datei.exists) {
            return datei;
        }

        this.logger.log("Datei nicht am Originalpfad gefunden, starte Fallback-Suche: " + datei.name);

        var fallbackRoot = Folder(FALLBACK_SUCHPFAD);
        if (!fallbackRoot.exists) {
            this.logger.logWarnung("Fallback-Ordner existiert nicht: " + FALLBACK_SUCHPFAD);
            return null;
        }

        var dateiName = datei.name;
        var alleTreffer = [];
        this.sucheRekursiv(fallbackRoot, dateiName, alleTreffer, FALLBACK_SUCHTIEFE, 0);

        if (alleTreffer.length === 0) {
            this.logger.logWarnung("Datei nicht gefunden: " + dateiName);
            return null;
        }

        if (alleTreffer.length === 1) {
            this.logger.log("Datei via Fallback gefunden: " + alleTreffer[0].fsName);
            return alleTreffer[0];
        }

        alleTreffer.sort(function(a, b) {
            return b.modified.getTime() - a.modified.getTime();
        });

        this.logger.log("Mehrere Treffer gefunden, verwende neueste: " + alleTreffer[0].fsName);
        return alleTreffer[0];
    };

    /**
     * Sucht rekursiv nach einer Datei
     *
     * @param {Folder} ordner - Startordner
     * @param {string} dateiName - Gesuchter Dateiname
     * @param {Array} alleTreffer - Array für Treffer
     * @param {number} maxTiefe - Maximale Tiefe
     * @param {number} aktuelleTiefe - Aktuelle Tiefe
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
     * Prüft ob eine Datei existiert
     *
     * @param {string} pfad - Dateipfad
     * @returns {boolean} True wenn Datei existiert
     */
    this.dateiExistiert = function(pfad) {
        var datei = File(pfad);
        return datei.exists;
    };

    /**
     * Holt den Basisnamen einer Datei (ohne Extension)
     *
     * @param {string} dateiName - Der Dateiname
     * @returns {string} Basisname ohne Extension
     */
    this.holeBasisname = function(dateiName) {
        var letzterpunkt = dateiName.lastIndexOf(".");
        if (letzterpunkt === -1) return dateiName;
        return dateiName.substring(0, letzterpunkt);
    };

    /**
     * Ersetzt die Extension einer Datei
     *
     * @param {string} dateiName - Der Dateiname
     * @param {string} neueExt - Neue Extension (mit Punkt)
     * @returns {string} Dateiname mit neuer Extension
     */
    this.ersetzeExtension = function(dateiName, neueExt) {
        var basis = this.holeBasisname(dateiName);
        return basis + neueExt;
    };
}
/**
 * DocumentService - InDesign-Dokumentoperationen
 *
 * Verantwortlich für alle Dokument-bezogenen Operationen.
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {Logger} dependencies.logger - Logger-Instanz
 * @param {ErrorHandler} dependencies.errorHandler - ErrorHandler-Instanz
 */
function DocumentService(dependencies) {
    this.logger = dependencies.logger;
    this.errorHandler = dependencies.errorHandler;

    /**
     * Öffnet ein InDesign-Dokument
     *
     * @param {File} datei - Die zu öffnende Datei
     * @param {boolean} userInteraction - User-Interaktion erlauben?
     * @returns {Document} Das geöffnete Dokument
     * @throws {Error} Bei Öffnungsfehlern
     */
    this.oeffneDokument = function(datei, userInteraction) {
        if (!datei.exists) {
            throw new Error("Datei existiert nicht: " + datei.fsName);
        }

        try {
            this.logger.log("Öffne Dokument: " + datei.name);
            var doc = app.open(datei, !userInteraction);
            return doc;
        } catch (fehler) {
            var fehlermeldung = this.errorHandler.formatiereFehler(fehler, datei.name);
            this.logger.logFehler("Fehler beim Öffnen: " + fehlermeldung, datei.name);
            throw new Error("Fehler beim Öffnen: " + fehlermeldung);
        }
    };

    /**
     * Schließt ein Dokument
     *
     * @param {Document} doc - Das zu schließende Dokument
     * @param {SaveOptions} saveOption - Speicher-Option
     */
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

    /**
     * Holt alle Links eines Dokuments mit optionalem Filter
     *
     * @param {Document} doc - Das Dokument
     * @param {string} filterExtension - Optional: Filter nach Extension (z.B. ".indd")
     * @returns {Array} Array von Links
     */
    this.holeLinks = function(doc, filterExtension) {
        var alleLinks = doc.links;
        if (!filterExtension) {
            return alleLinks;
        }

        var gefilterteLinks = [];
        for (var i = 0; i < alleLinks.length; i++) {
            if (alleLinks[i].name.toLowerCase().indexOf(filterExtension.toLowerCase()) > -1) {
                gefilterteLinks.push(alleLinks[i]);
            }
        }

        return gefilterteLinks;
    };

    /**
     * Holt alle Layer eines Dokuments
     *
     * @param {Document} doc - Das Dokument
     * @returns {Array} Array von Layer-Namen
     */
    this.holeAlleLayer = function(doc) {
        var layerNamen = [];
        for (var i = 0; i < doc.layers.length; i++) {
            layerNamen.push(doc.layers[i].name);
        }
        return layerNamen;
    };

    /**
     * Holt einen Layer nach Name
     *
     * @param {Document} doc - Das Dokument
     * @param {string} name - Layer-Name
     * @returns {Layer|null} Der Layer oder null
     */
    this.holeLayerNachName = function(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        return null;
    };

    /**
     * Stellt sicher, dass die TrimBox verwendet wird
     *
     * @param {Link} link - Der Link
     */
    this.stelleTrimBoxEin = function(link) {
        if (link.parent && (link.parent.constructor.name === "Image" ||
                           link.parent.constructor.name === "PDF" ||
                           link.parent.constructor.name === "ImportedPage")) {
            try {
                link.parent.importedPageAttributes.pageBoundingBox = BoundingBoxOptions.TRIM_BOX;
            } catch (e) {
                // Fehler ignorieren wenn Eigenschaft nicht existiert
            }
        }
    };

    /**
     * Holt die Seitennummer eines Page-Objekts
     *
     * @param {Page} page - Das Page-Objekt
     * @returns {number} Die Seitennummer
     */
    this.holeSeitennummer = function(page) {
        if (!page || !page.name) return -1;
        var seitenNummer = parseInt(page.name, 10);
        return isNaN(seitenNummer) ? -1 : seitenNummer;
    };
}
/**
 * CsvService - CSV-Datei-Operationen
 *
 * Verantwortlich für Schreiben von CSV-Dateien.
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {DateUtils} dependencies.dateUtils - DateUtils-Instanz
 * @param {Logger} dependencies.logger - Logger-Instanz
 */
function CsvService(dependencies) {
    this.dateUtils = dependencies.dateUtils;
    this.logger = dependencies.logger;

    /**
     * Schreibt eine einzelne CSV-Datei
     *
     * @param {Array} datenArray - Array von CsvDataModel-Objekten
     * @param {Folder} ordner - Zielordner
     * @param {string} dateiName - Dateiname
     * @returns {File} Die erstellte CSV-Datei
     */
    this.schreibeEinzelneCsv = function(datenArray, ordner, dateiName) {
        this.logger.log("Erstelle CSV-Datei: " + dateiName);

        var inhalt = "";

        if (datenArray.length > 0) {
            inhalt += datenArray[0].header;
        }

        for (var i = 0; i < datenArray.length; i++) {
            if (datenArray[i].zeilen) {
                inhalt += datenArray[i].zeilen;
            }
        }

        inhalt = this.entferneDuplikate(inhalt);

        var csvDatei = new File(ordner.fsName + "/" + dateiName);
        csvDatei.encoding = "UTF-8";

        try {
            csvDatei.open("w");
            var erfolg = csvDatei.write(inhalt);
            csvDatei.close();

            if (!erfolg) {
                throw new Error("Fehler beim Schreiben der CSV-Datei");
            }

            this.logger.log("CSV-Datei erfolgreich erstellt: " + csvDatei.fsName);
            return csvDatei;
        } catch (fehler) {
            this.logger.logFehler("Fehler beim Schreiben der CSV: " + fehler.message, dateiName);
            throw fehler;
        }
    };

    /**
     * Schreibt mehrere CSV-Dateien (eine pro INDD)
     *
     * @param {Array} datenArray - Array von CsvDataModel-Objekten
     * @param {Folder} ordner - Zielordner
     * @returns {Array} Array der erstellten CSV-Dateien
     */
    this.schreibeMehrereCsvs = function(datenArray, ordner) {
        this.logger.log("Erstelle " + datenArray.length + " CSV-Dateien");

        var erstellteDateien = [];

        for (var i = 0; i < datenArray.length; i++) {
            var daten = datenArray[i];
            var dateiName = this.erstelleDatainame(daten);

            var inhalt = daten.header;
            if (daten.zeilen) {
                inhalt += daten.zeilen;
            }

            inhalt = this.entferneDuplikate(inhalt);

            var csvDatei = new File(ordner.fsName + "/" + dateiName);
            csvDatei.encoding = "UTF-8";

            try {
                csvDatei.open("w");
                csvDatei.write(inhalt);
                csvDatei.close();

                erstellteDateien.push(csvDatei);
                this.logger.log("CSV-Datei erstellt: " + dateiName);
            } catch (fehler) {
                this.logger.logFehler("Fehler beim Schreiben: " + fehler.message, dateiName);
            }
        }

        return erstellteDateien;
    };

    /**
     * Erstellt einen Dateinamen für CSV
     *
     * @param {Object} csvData - CsvDataModel
     * @returns {string} Dateiname
     */
    this.erstelleDatainame = function(csvData) {
        var basis = csvData.quellIndd.replace(/\.indd$/i, "");
        var zeitstempel = this.dateUtils.formatiereZeitstempel(new Date());
        return basis + "_export_" + zeitstempel + ".csv";
    };

    /**
     * Entfernt doppelte Zeilen aus CSV-Inhalt
     *
     * @param {string} inhalt - CSV-Inhalt
     * @returns {string} Bereinigter Inhalt
     */
    this.entferneDuplikate = function(inhalt) {
        var zeilen = inhalt.split('\n');
        var eindeutigeZeilen = {};
        var ergebnis = [];

        for (var i = 0; i < zeilen.length; i++) {
            var zeile = zeilen[i];
            if (zeile && !eindeutigeZeilen.hasOwnProperty(zeile)) {
                eindeutigeZeilen[zeile] = true;
                ergebnis.push(zeile);
            }
        }

        return ergebnis.join('\n');
    };
}
/**
 * GraphicsProcessor - Graphics-Extraktion
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {StringUtils} dependencies.stringUtils
 * @param {Logger} dependencies.logger
 */
function GraphicsProcessor(dependencies) {
    this.stringUtils = dependencies.stringUtils;
    this.logger = dependencies.logger;

    /**
     * Verarbeitet Graphics auf einer Seite
     *
     * @param {Page} seite - Die zu verarbeitende Seite
     * @param {Layer} layer - Der zu verwendende Layer
     * @param {number} parentPage - Elternseite
     * @returns {string} CSV-Zeilen für Graphics
     */
    this.verarbeiteGraphics = function(seite, layer, parentPage) {
        var zeilen = [];
        var seitenName = seite.name;
        var seitenIndex = seite.index;

        for (var i = 0; i < seite.allPageItems.length; i++) {
            var pageItem = seite.allPageItems[i];

            for (var j = 0; j < pageItem.allGraphics.length; j++) {
                var graphic = pageItem.allGraphics[j];

                if (graphic.itemLayer !== layer) continue;
                if (!graphic.itemLink || !graphic.visibleBounds) continue;

                var dateiName = this.stringUtils.kodiereDatainame(graphic.itemLink.name);
                var bounds = graphic.visibleBounds;

                var zeile = CSV_ZEILE_GRAPHIC + "," + seitenName + "," + seitenIndex + "," +
                           dateiName + "," + bounds;

                var parentObject = graphic.parent;
                zeile += this.sammleParentShapes(parentObject);

                zeilen.push(zeile);
            }
        }

        return zeilen.join("\n") + (zeilen.length > 0 ? "\n" : "");
    };

    /**
     * Sammelt Parent-Shape-Bounds
     *
     * @param {Object} parentObject - Parent-Objekt
     * @returns {string} Komma-separierte Bounds
     */
    this.sammleParentShapes = function(parentObject) {
        var bounds = "";
        var parentType = String(parentObject);

        while (parentType.indexOf("Polygon") > -1 ||
               parentType.indexOf("Rectangle") > -1 ||
               parentType.indexOf("Oval") > -1) {
            bounds += "," + parentObject.visibleBounds;
            parentObject = parentObject.parent;
            parentType = String(parentObject);
        }

        return bounds;
    };
}
/**
 * TextProcessor - Text-Extraktion
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {RegexHelper} dependencies.regexHelper
 * @param {Logger} dependencies.logger
 */
function TextProcessor(dependencies) {
    this.regexHelper = dependencies.regexHelper;
    this.logger = dependencies.logger;

    /**
     * Verarbeitet TextFrames direkt
     *
     * @param {Document} doc - Das Dokument
     * @param {number} parentPage - Elternseite
     * @param {number} targetPage - Zielseite
     * @param {Object} settings - Einstellungen
     * @returns {string} CSV-Zeilen
     */
    this.verarbeiteTextFrames = function(doc, parentPage, targetPage, settings) {
        if (!settings.export.textFrames) return "";

        var zeilen = [];
        var page = doc.pages[targetPage - 1];

        for (var i = 0; i < page.allPageItems.length; i++) {
            var item = page.allPageItems[i];
            if (item instanceof TextFrame) {
                var lines = item.lines;
                var result = this.verarbeiteZeilen(lines, page.name, settings.regex.textFrame);
                if (result) zeilen.push(result);
            }
        }

        return zeilen.join("");
    };

    /**
     * Verarbeitet Stories
     *
     * @param {Document} doc - Das Dokument
     * @param {number} parentPage - Elternseite
     * @param {number} targetPage - Zielseite
     * @param {Object} settings - Einstellungen
     * @returns {string} CSV-Zeilen
     */
    this.verarbeiteStories = function(doc, parentPage, targetPage, settings) {
        if (!settings.export.textInStories) return "";

        var zeilen = [];
        var targetPageObj = doc.pages[targetPage - 1];

        for (var i = 0; i < doc.stories.length; i++) {
            var story = doc.stories[i];

            for (var j = 0; j < story.textFrames.length; j++) {
                var textFrame = story.textFrames[j];

                if (textFrame.parentPage && textFrame.parentPage.name === targetPageObj.name) {
                    var lines = textFrame.lines;
                    var result = this.verarbeiteZeilen(lines, targetPageObj.name, settings.regex.textFrame);
                    if (result) zeilen.push(result);
                }
            }
        }

        return zeilen.join("");
    };

    /**
     * Verarbeitet einzelne Zeilen mit Regex
     *
     * @param {Lines} lines - Die Zeilen
     * @param {string} seitenName - Seitenname
     * @param {RegExp} regex - Regex für Artikelnummer
     * @returns {string} CSV-Zeilen
     */
    this.verarbeiteZeilen = function(lines, seitenName, regex) {
        var zeilen = [];

        for (var i = 0; i < lines.length; i++) {
            try {
                var line = lines[i];
                var inhalt = String(line.contents);
                var match = regex.exec(inhalt);

                if (!match) continue;

                var artikelNr = match[0];
                var chars = line.characters;
                var baseline = chars[0].baseline;
                var offset = chars[0].horizontalOffset;
                var endOffset = chars[chars.length - 1].horizontalOffset;

                var zeile = CSV_ZEILE_TEXT + "," + seitenName + "," + artikelNr + "," +
                           (baseline - line.ascent - line.descent) + "," + offset + "," +
                           baseline + "," + endOffset;

                zeilen.push(zeile);
            } catch (e) {
                // Fehler bei einzelner Zeile ignorieren
            }
        }

        return zeilen.length > 0 ? zeilen.join("\n") + "\n" : "";
    };
}
/**
 * TableProcessor - Tabellen-Extraktion
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {TextProcessor} dependencies.textProcessor
 * @param {RegexHelper} dependencies.regexHelper
 * @param {Logger} dependencies.logger
 */
function TableProcessor(dependencies) {
    this.textProcessor = dependencies.textProcessor;
    this.regexHelper = dependencies.regexHelper;
    this.logger = dependencies.logger;

    /**
     * Verarbeitet Tabellen in Stories
     *
     * @param {Document} doc - Das Dokument
     * @param {number} parentPage - Elternseite
     * @param {number} targetPage - Zielseite
     * @param {Object} settings - Einstellungen
     * @returns {string} CSV-Zeilen
     */
    this.verarbeiteTables = function(doc, parentPage, targetPage, settings) {
        if (!settings.export.tables) return "";

        var zeilen = [];
        var targetPageObj = doc.pages[targetPage - 1];

        for (var i = 0; i < doc.stories.length; i++) {
            var story = doc.stories[i];

            for (var j = 0; j < story.textFrames.length; j++) {
                var textFrame = story.textFrames[j];

                if (textFrame.parentPage && textFrame.parentPage.name === targetPageObj.name) {
                    for (var k = 0; k < textFrame.tables.length; k++) {
                        var table = textFrame.tables[k];
                        var result = this.verarbeiteTabelle(table, targetPageObj.name, settings.regex.table);
                        if (result) zeilen.push(result);
                    }
                }
            }
        }

        return zeilen.join("");
    };

    /**
     * Verarbeitet eine einzelne Tabelle
     *
     * @param {Table} table - Die Tabelle
     * @param {string} seitenName - Seitenname
     * @param {RegExp} regex - Regex für Artikelnummer
     * @returns {string} CSV-Zeilen
     */
    this.verarbeiteTabelle = function(table, seitenName, regex) {
        var zeilen = [];
        var cells = table.cells;

        for (var i = 0; i < cells.length; i++) {
            var cell = cells[i];
            var textStyleRanges = cell.textStyleRanges;

            if (textStyleRanges.length > 0) {
                var lines = textStyleRanges[0].lines;
                var result = this.textProcessor.verarbeiteZeilen(lines, seitenName, regex);
                if (result) zeilen.push(result);
            }
        }

        return zeilen.join("");
    };
}
/**
 * PageItemsProcessor - PageItems mit Labels
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {StringUtils} dependencies.stringUtils
 * @param {Logger} dependencies.logger
 */
function PageItemsProcessor(dependencies) {
    this.stringUtils = dependencies.stringUtils;
    this.logger = dependencies.logger;

    /**
     * Verarbeitet PageItems mit Labels
     *
     * @param {Page} seite - Die Seite
     * @param {Layer} layer - Der Layer
     * @param {number} parentPage - Elternseite
     * @returns {string} CSV-Zeilen
     */
    this.verarbeitePageItems = function(seite, layer, parentPage) {
        var zeilen = [];

        for (var i = 0; i < seite.allPageItems.length; i++) {
            var item = seite.allPageItems[i];

            if (item.label === "") continue;
            if (item.itemLayer !== layer) continue;

            var linkType = "unknown";
            var bounds = item.visibleBounds;

            if (item.allGraphics.length === 1) {
                var graphic = item.allGraphics[0];
                if (graphic && graphic.itemLink && graphic.itemLink.name) {
                    linkType = graphic.itemLink.name;
                    bounds = graphic.visibleBounds;
                }
            }

            var zeile = CSV_ZEILE_PAGEITEM + "," + seite.name + ",\"" +
                       File.encode(linkType) + "\",\"" + File.encode(item.label) + "\"," + bounds;

            zeile += this.sammleParentItems(item.parent);

            zeilen.push(zeile);
        }

        return zeilen.length > 0 ? zeilen.join("\n") + "\n" : "";
    };

    /**
     * Sammelt Parent-Item-Bounds
     *
     * @param {Object} parentObject - Parent-Objekt
     * @returns {string} Komma-separierte Bounds
     */
    this.sammleParentItems = function(parentObject) {
        var bounds = "";
        var parentType = String(parentObject);

        while (parentType.indexOf("Polygon") > -1 ||
               parentType.indexOf("Rectangle") > -1 ||
               parentType.indexOf("Oval") > -1 ||
               parentType.indexOf("group") > -1) {
            bounds += "," + parentObject.visibleBounds;
            parentObject = parentObject.parent;
            parentType = String(parentObject);
        }

        return bounds;
    };
}
/**
 * InddProcessor - Hauptverarbeitung eines INDD-Dokuments
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {GraphicsProcessor} dependencies.graphicsProcessor
 * @param {TextProcessor} dependencies.textProcessor
 * @param {TableProcessor} dependencies.tableProcessor
 * @param {PageItemsProcessor} dependencies.pageItemsProcessor
 * @param {DocumentService} dependencies.documentService
 * @param {Logger} dependencies.logger
 */
function InddProcessor(dependencies) {
    this.graphicsProcessor = dependencies.graphicsProcessor;
    this.textProcessor = dependencies.textProcessor;
    this.tableProcessor = dependencies.tableProcessor;
    this.pageItemsProcessor = dependencies.pageItemsProcessor;
    this.documentService = dependencies.documentService;
    this.logger = dependencies.logger;

    /**
     * Verarbeitet ein INDD-Dokument
     *
     * @param {Document} inddDoc - Das zu verarbeitende Dokument
     * @param {Object} settings - Export-Einstellungen
     * @param {Object} sourceItem - Das Quell-ExportItem
     * @param {Function} layerDialogCallback - Callback für Layer-Auswahl
     * @returns {Object} CsvDataModel
     */
    this.verarbeiteDokument = function(inddDoc, settings, sourceItem, layerDialogCallback) {
        this.logger.log("Verarbeite Dokument: " + inddDoc.name);

        var layer = this.ermittleLayer(inddDoc, settings, layerDialogCallback);
        if (!layer) {
            throw new Error("Layer konnte nicht ermittelt werden für: " + inddDoc.name);
        }

        var targetPage = this.ermittleZielseite(sourceItem);
        this.logger.log("Verwende Layer: " + layer.name + ", Zielseite: " + targetPage);

        var alleDaten = this.sammleAlleDaten(inddDoc, layer, targetPage, settings);

        return erstelleCsvDataModel(inddDoc, sourceItem, alleDaten, layer.name);
    };

    /**
     * Ermittelt den zu verwendenden Layer
     *
     * @param {Document} doc - Das Dokument
     * @param {Object} settings - Einstellungen
     * @param {Function} layerDialogCallback - Callback für Dialog
     * @returns {Layer} Der ermittelte Layer
     */
    this.ermittleLayer = function(doc, settings, layerDialogCallback) {
        if (settings.layer.auto && layerDialogCallback) {
            var layerNamen = this.documentService.holeAlleLayer(doc);
            var ausgewaehlterLayer = layerDialogCallback(layerNamen, doc.name);
            if (ausgewaehlterLayer) {
                return this.documentService.holeLayerNachName(doc, ausgewaehlterLayer);
            }
        }

        if (settings.layer.manual && settings.layer.name) {
            return this.documentService.holeLayerNachName(doc, settings.layer.name);
        }

        if (doc.layers.length > 0) {
            return doc.layers[0];
        }

        return null;
    };

    /**
     * Ermittelt die Zielseite
     *
     * @param {Object} sourceItem - Das Quell-Item
     * @returns {number} Seitennummer
     */
    this.ermittleZielseite = function(sourceItem) {
        try {
            if (sourceItem.link && sourceItem.link.parent && sourceItem.link.parent.pageNumber) {
                return sourceItem.link.parent.pageNumber;
            }
        } catch (e) {
            this.logger.logWarnung("Konnte Zielseite nicht ermitteln, verwende Seite 1");
        }
        return 1;
    };

    /**
     * Sammelt alle Daten aus dem Dokument
     *
     * @param {Document} doc - Das Dokument
     * @param {Layer} layer - Der Layer
     * @param {number} targetPage - Zielseite
     * @param {Object} settings - Einstellungen
     * @returns {string} Gesammelte CSV-Zeilen
     */
    this.sammleAlleDaten = function(doc, layer, targetPage, settings) {
        var daten = "";
        var page = doc.pages[targetPage - 1];

        if (settings.export.graphics) {
            daten += this.graphicsProcessor.verarbeiteGraphics(page, layer, targetPage);
        }

        if (settings.export.textFrames) {
            daten += this.textProcessor.verarbeiteTextFrames(doc, targetPage, targetPage, settings);
        }

        if (settings.export.textInStories) {
            daten += this.textProcessor.verarbeiteStories(doc, targetPage, targetPage, settings);
        }

        if (settings.export.tables) {
            daten += this.tableProcessor.verarbeiteTables(doc, targetPage, targetPage, settings);
        }

        if (settings.export.pageItems) {
            daten += this.pageItemsProcessor.verarbeitePageItems(page, layer, targetPage);
        }

        return daten;
    };
}
/**
 * DialogFactory - Zentrale Dialog-Erstellung
 *
 * Erstellt alle UI-Dialoge mit einheitlichem Styling.
 */
function DialogFactory() {
    /**
     * Erstellt den Hauptdialog
     *
     * @returns {number} User-Auswahl (0=Abbruch, 1-4=Aktionen)
     */
    this.erstelleHauptdialog = function() {
        var dialog = new Window("dialog", SKRIPT_NAME + " " + SKRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        var infoText = "Dieses Skript hilft bei folgenden Aktionen:\n" +
                       "- INDD-Verknüpfungen durch PDFs ersetzen\n" +
                       "- PDF-Verknüpfungen durch INDDs ersetzen\n" +
                       "- Verknüpfungen im Katalog validieren\n" +
                       "- CSV-Daten für Blätterkatalog exportieren\n\n" +
                       "Bitte wähle eine Aktion:";

        dialog.add("statictext", undefined, infoText, {multiline: true});

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "column";
        buttonGroup.alignChildren = ["fill", "center"];

        var inddZuPdfBtn = buttonGroup.add("button", undefined, "INDD -> PDF Katalog erstellen");
        var pdfZuInddBtn = buttonGroup.add("button", undefined, "PDF -> INDD Katalog erstellen");
        var validierenBtn = buttonGroup.add("button", undefined, "Katalog validieren");
        var csvExportBtn = buttonGroup.add("button", undefined, "CSV Daten Export");

        var abbrechenBtn = dialog.add("button", undefined, "Abbrechen", {name: "cancel"});

        inddZuPdfBtn.onClick = function() { dialog.close(1); };
        pdfZuInddBtn.onClick = function() { dialog.close(2); };
        validierenBtn.onClick = function() { dialog.close(3); };
        csvExportBtn.onClick = function() { dialog.close(4); };
        abbrechenBtn.onClick = function() { dialog.close(0); };

        return dialog.show();
    };

    /**
     * Erstellt Layer-Auswahl-Dialog
     *
     * @param {Array} layerNamen - Array von Layer-Namen
     * @param {string} dokumentName - Name des Dokuments
     * @returns {string|null} Ausgewählter Layer-Name
     */
    this.erstelleLayerDialog = function(layerNamen, dokumentName) {
        var dialog = new Window("dialog", "Layer auswählen");
        dialog.orientation = "column";

        dialog.add("statictext", undefined, "Dokument: " + dokumentName);
        dialog.add("statictext", undefined, "Wähle eine Ebene:");

        var dropdown = dialog.add("dropdownlist", undefined, layerNamen);
        dropdown.selection = 0;

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        var okButton = buttonGroup.add("button", undefined, "OK");
        var cancelButton = buttonGroup.add("button", undefined, "Abbrechen");

        var ausgewaehlterLayer = null;

        okButton.onClick = function() {
            ausgewaehlterLayer = layerNamen[dropdown.selection.index];
            dialog.close();
        };

        cancelButton.onClick = function() {
            dialog.close();
        };

        dialog.show();
        return ausgewaehlterLayer;
    };

    /**
     * Erstellt Fortschritts-Dialog
     *
     * @param {string} nachricht - Initiale Nachricht
     * @returns {Object} Dialog-Objekt mit Methoden
     */
    this.erstelleFortschrittsDialog = function(nachricht) {
        var w = new Window("palette", "Fortschritt", undefined, {closeButton: false});
        w.alignChildren = "fill";

        var t = w.add("statictext", undefined, nachricht);
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
        var abgebrochen = false;

        cancelButton.onClick = function() {
            abgebrochen = true;
            w.close();
        };

        function formatTime(seconds) {
            var h = Math.floor(seconds / 3600);
            var m = Math.floor((seconds % 3600) / 60);
            var s = Math.floor(seconds % 60);
            return (h < 10 ? "0" : "") + h + ":" + (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s;
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
            }
            w.layout.layout(true);
        }

        var dialogObj = {
            schliesse: function() { w.close(); },
            erhoehe: function(value) { b.value += value; updateTiming(); },
            setzeNachricht: function(msg) { t.text = msg; w.update(); },
            setzeMax: function(steps) { b.maxvalue = steps; updateTiming(); },
            istAbgebrochen: function() { return abgebrochen; }
        };

        w.show();
        return dialogObj;
    };

    /**
     * Erstellt einfachen Bestätigungs-Dialog
     *
     * @param {string} nachricht - Die Nachricht
     * @param {string} titel - Dialog-Titel
     * @returns {boolean} True wenn bestätigt
     */
    this.erstelleBestaetigungsDialog = function(nachricht, titel) {
        var result = confirm(nachricht, false, titel || SKRIPT_NAME);
        return result;
    };
}
/**
 * ValidationEngine - Link-Analyse und Mode-Detection
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {FileService} dependencies.fileService
 * @param {Logger} dependencies.logger
 */
function ValidationEngine(dependencies) {
    this.fileService = dependencies.fileService;
    this.logger = dependencies.logger;

    /**
     * Validiert ein Dokument und bestimmt den Export-Modus
     *
     * @param {Document} doc - Das zu validierende Dokument
     * @returns {Object} ValidationResult
     */
    this.validiereDokument = function(doc) {
        this.logger.log("Starte Validierung für: " + doc.name);

        var alleLinks = doc.links;
        var inddLinks = [];
        var pdfLinks = [];

        for (var i = 0; i < alleLinks.length; i++) {
            var linkModel = erstelleLinkModel(alleLinks[i]);

            if (linkModel.typ === "INDD") {
                inddLinks.push(linkModel);
            } else if (linkModel.typ === "PDF") {
                pdfLinks.push(linkModel);
            }
        }

        this.logger.log("Gefunden: " + inddLinks.length + " INDD-Links, " + pdfLinks.length + " PDF-Links");

        var modus = this.bestimmeModus(inddLinks, pdfLinks);
        var mappings = [];

        if (modus === EXPORT_MODE_CONVERT || modus === EXPORT_MODE_MIXED) {
            mappings = this.findeInddPaare(pdfLinks);
        }

        return {
            modus: modus,
            inddLinks: inddLinks,
            pdfLinks: pdfLinks,
            mappings: mappings,
            stats: {
                totalLinks: alleLinks.length,
                validIndd: inddLinks.length,
                validPdf: pdfLinks.length
            }
        };
    };

    /**
     * Bestimmt den Export-Modus
     *
     * @param {Array} inddLinks - INDD-Links
     * @param {Array} pdfLinks - PDF-Links
     * @returns {string} Modus
     */
    this.bestimmeModus = function(inddLinks, pdfLinks) {
        var hatIndd = inddLinks.length > 0;
        var hatPdf = pdfLinks.length > 0;

        if (hatIndd && !hatPdf) {
            this.logger.log("Modus: DIRECT (nur INDD-Links)");
            return EXPORT_MODE_DIRECT;
        }

        if (!hatIndd && hatPdf) {
            this.logger.log("Modus: CONVERT (nur PDF-Links)");
            return EXPORT_MODE_CONVERT;
        }

        if (hatIndd && hatPdf) {
            this.logger.log("Modus: MIXED (INDD + PDF-Links)");
            return EXPORT_MODE_MIXED;
        }

        this.logger.logFehler("Keine relevanten Links gefunden", "ValidationEngine");
        return EXPORT_MODE_ERROR;
    };

    /**
     * Findet INDD-Paare für PDF-Links
     *
     * @param {Array} pdfLinks - PDF-Links
     * @returns {Array} Mappings von PDF zu INDD
     */
    this.findeInddPaare = function(pdfLinks) {
        var mappings = [];

        for (var i = 0; i < pdfLinks.length; i++) {
            var pdfLink = pdfLinks[i];
            var pdfDatei = File(pdfLink.dateiPfad);
            var pdfOrdner = pdfDatei.path;

            var basisName = this.fileService.holeBasisname(pdfLink.dateiName);
            var inddName = basisName + ".indd";
            var inddPfadOriginal = pdfOrdner + "/" + inddName;

            var inddDatei = this.fileService.findeDateiMitFallback(inddPfadOriginal);

            mappings.push({
                pdfLink: pdfLink,
                inddDatei: inddDatei,
                gefunden: inddDatei !== null && inddDatei.exists
            });

            if (inddDatei) {
                this.logger.log("INDD-Paar gefunden: " + pdfLink.dateiName + " -> " + inddDatei.name);
            } else {
                this.logger.logWarnung("Kein INDD-Paar gefunden für: " + pdfLink.dateiName);
            }
        }

        return mappings;
    };
}
/**
 * ExportCoordinator - Export-Prozess Koordination
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {DocumentService} dependencies.documentService
 * @param {InddProcessor} dependencies.inddProcessor
 * @param {CsvService} dependencies.csvService
 * @param {DialogFactory} dependencies.dialogFactory
 * @param {Logger} dependencies.logger
 * @param {ErrorHandler} dependencies.errorHandler
 */
function ExportCoordinator(dependencies) {
    this.documentService = dependencies.documentService;
    this.inddProcessor = dependencies.inddProcessor;
    this.csvService = dependencies.csvService;
    this.dialogFactory = dependencies.dialogFactory;
    this.logger = dependencies.logger;
    this.errorHandler = dependencies.errorHandler;

    /**
     * Koordiniert den gesamten Export-Prozess
     *
     * @param {Object} exportPlan - Der Export-Plan
     * @param {Object} settings - Export-Einstellungen
     * @returns {Object} Ergebnis mit Daten und Fehlern
     */
    this.koordiniereExport = function(exportPlan, settings) {
        this.logger.log("Starte Export-Koordination: " + exportPlan.gesamtAnzahl + " Items");

        var progressDialog = this.dialogFactory.erstelleFortschrittsDialog("INDD Dateien werden verarbeitet...");
        progressDialog.setzeMax(exportPlan.items.length);

        var alleDaten = [];
        var fehlerListe = [];

        var alteUILevel = app.scriptPreferences.userInteractionLevel;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

        var layerDialogGezeigt = false;
        var ausgewaehlterLayer = null;

        for (var i = 0; i < exportPlan.items.length; i++) {
            if (progressDialog.istAbgebrochen()) {
                this.logger.log("Export wurde vom Benutzer abgebrochen");
                break;
            }

            var item = exportPlan.items[i];
            progressDialog.setzeNachricht("Verarbeite " + (i + 1) + "/" + exportPlan.items.length + ": " + item.inddDatei.name);

            try {
                var csvData = this.verarbeiteItem(item, settings, function(layerNamen, docName) {
                    if (!layerDialogGezeigt && settings.layer.auto) {
                        layerDialogGezeigt = true;
                        ausgewaehlterLayer = this.dialogFactory.erstelleLayerDialog(layerNamen, docName);
                        return ausgewaehlterLayer;
                    }
                    return ausgewaehlterLayer;
                }.bind(this));

                if (csvData) {
                    alleDaten.push(csvData);
                }
            } catch (fehler) {
                this.errorHandler.sammleFehler(fehler, item.inddDatei.name, fehlerListe);
            }

            if (i % PROGRESS_UPDATE_INTERVALL === 0 || i === exportPlan.items.length - 1) {
                progressDialog.erhoehe(PROGRESS_UPDATE_INTERVALL);
            }
        }

        progressDialog.schliesse();
        app.scriptPreferences.userInteractionLevel = alteUILevel;

        this.logger.log("Export abgeschlossen: " + alleDaten.length + " erfolgreich, " + fehlerListe.length + " Fehler");

        return {
            daten: alleDaten,
            fehler: fehlerListe,
            erfolgreich: alleDaten.length,
            fehlgeschlagen: fehlerListe.length
        };
    };

    /**
     * Verarbeitet ein einzelnes Export-Item
     *
     * @param {Object} item - Das ExportItem
     * @param {Object} settings - Einstellungen
     * @param {Function} layerDialogCallback - Callback für Layer-Dialog
     * @returns {Object} CsvDataModel
     */
    this.verarbeiteItem = function(item, settings, layerDialogCallback) {
        if (!item.inddDatei || !item.inddDatei.exists) {
            throw new Error("INDD-Datei existiert nicht");
        }

        this.logger.log("Öffne: " + item.inddDatei.name);
        var doc = this.documentService.oeffneDokument(item.inddDatei, false);

        try {
            var csvData = this.inddProcessor.verarbeiteDokument(doc, settings, item, layerDialogCallback);
            return csvData;
        } finally {
            this.documentService.schliesseDokument(doc, SaveOptions.NO);
        }
    };
}
/**
 * LinkManager - Hauptkoordinator
 *
 * Entry-Point für alle Workflows.
 *
 * @param {Object} dependencies - Alle Abhängigkeiten
 */
function LinkManager(dependencies) {
    this.validationEngine = dependencies.validationEngine;
    this.exportCoordinator = dependencies.exportCoordinator;
    this.dialogFactory = dependencies.dialogFactory;
    this.relinkWorkflow = dependencies.relinkWorkflow;
    this.validationWorkflow = dependencies.validationWorkflow;
    this.csvExportWorkflow = dependencies.csvExportWorkflow;
    this.logger = dependencies.logger;
    this.errorHandler = dependencies.errorHandler;

    this.aktivesDokument = null;
    this.basisDokPfad = null;
    this.basisDokNameOhneErw = null;

    /**
     * Initialisiert den LinkManager
     */
    this.initialisiere = function() {
        this.logger.log("LinkManager wird initialisiert");

        if (app.documents.length === 0) {
            alert("Bitte öffne ein Dokument, bevor Du dieses Skript ausführst.", SKRIPT_NAME);
            return false;
        }

        this.aktivesDokument = app.activeDocument;

        if (!this.aktivesDokument.saved || !this.aktivesDokument.filePath) {
            alert("Dein aktives Dokument muss zuerst gespeichert werden.\nBitte speichere Dein Dokument und führe das Skript erneut aus.", SKRIPT_NAME);
            return false;
        }

        this.basisDokPfad = this.aktivesDokument.filePath;
        this.basisDokNameOhneErw = this.aktivesDokument.name.replace(/\.indd$/i, "");

        if (LOGGING_AKTIVIERT) {
            var logOrdner = Folder(this.basisDokPfad);
            var logDatei = new File(logOrdner.fsName + "/" + this.basisDokNameOhneErw + "_log.txt");
            this.logger.initialisiere(logDatei, true);
        }

        this.logger.log("=== " + SKRIPT_NAME + " " + SKRIPT_VERSION + " ===");
        this.logger.log("Dokument: " + this.aktivesDokument.name);

        return true;
    };

    /**
     * Startet den Hauptdialog und führt Workflows aus
     */
    this.starte = function() {
        if (!this.initialisiere()) {
            return;
        }

        var auswahl = this.dialogFactory.erstelleHauptdialog();

        try {
            switch (auswahl) {
                case 1:
                    this.relinkWorkflow.fuehreAus(".indd", ".pdf", "INDD -> PDF");
                    break;
                case 2:
                    this.relinkWorkflow.fuehreAus(".pdf", ".indd", "PDF -> INDD");
                    break;
                case 3:
                    this.validationWorkflow.fuehreAus();
                    break;
                case 4:
                    this.csvExportWorkflow.fuehreAus();
                    break;
                default:
                    this.logger.log("Benutzer hat abgebrochen");
                    break;
            }
        } catch (fehler) {
            var fehlerMsg = this.errorHandler.formatiereFehler(fehler, "LinkManager");
            alert("Kritischer Fehler:\n" + fehlerMsg, SKRIPT_NAME + " - Fehler");
            this.logger.logFehler(fehlerMsg, "LinkManager.starte");
        } finally {
            this.raeumerAuf();
        }
    };

    /**
     * Räumt auf nach der Ausführung
     */
    this.raeumerAuf = function() {
        this.logger.log("=== Script beendet ===");
        this.logger.schliesse();
    };
}
/**
 * RelinkWorkflow - INDD <-> PDF Link-Austausch
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {FileService} dependencies.fileService
 * @param {DocumentService} dependencies.documentService
 * @param {DialogFactory} dependencies.dialogFactory
 * @param {ErrorHandler} dependencies.errorHandler
 * @param {Logger} dependencies.logger
 */
function RelinkWorkflow(dependencies) {
    this.fileService = dependencies.fileService;
    this.documentService = dependencies.documentService;
    this.dialogFactory = dependencies.dialogFactory;
    this.errorHandler = dependencies.errorHandler;
    this.logger = dependencies.logger;

    /**
     * Führt den Relink-Workflow aus
     *
     * @param {string} quellExt - Quell-Extension (z.B. ".indd")
     * @param {string} zielExt - Ziel-Extension (z.B. ".pdf")
     * @param {string} titel - Aktion-Titel
     */
    this.fuehreAus = function(quellExt, zielExt, titel) {
        this.logger.log("Starte " + titel);

        var doc = app.activeDocument;
        var links = this.documentService.holeLinks(doc, quellExt);

        if (links.length === 0) {
            alert("Keine " + quellExt.toUpperCase() + "-Verknüpfungen im Dokument gefunden.", SKRIPT_NAME);
            this.logger.log("Keine passenden Links gefunden");
            return;
        }

        this.logger.log(links.length + " " + quellExt + "-Links gefunden");

        var progressDialog = this.dialogFactory.erstelleFortschrittsDialog("Verknüpfungen werden verarbeitet...");
        progressDialog.setzeMax(links.length);

        var fehlerListe = [];
        var erfolge = 0;

        var alteUILevel = app.scriptPreferences.userInteractionLevel;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

        for (var i = 0; i < links.length; i++) {
            if (progressDialog.istAbgebrochen()) {
                break;
            }

            var link = links[i];
            progressDialog.setzeNachricht("Verarbeite " + (i + 1) + "/" + links.length + ": " + link.name);

            try {
                var basisName = this.fileService.holeBasisname(link.name);
                var zielName = basisName + zielExt;

                var linkDatei = File(link.filePath);
                var zielPfad = linkDatei.path + "/" + zielName;

                var zielDatei = this.fileService.findeDateiMitFallback(zielPfad);

                if (zielDatei && zielDatei.exists) {
                    link.relink(zielDatei);
                    this.documentService.stelleTrimBoxEin(link);
                    erfolge++;
                    this.logger.log("Erfolg: " + link.name + " -> " + zielDatei.name);
                } else {
                    throw new Error("Zieldatei nicht gefunden: " + zielName);
                }
            } catch (fehler) {
                this.errorHandler.sammleFehler(fehler, link.name, fehlerListe);
            }

            progressDialog.erhoehe(1);
        }

        progressDialog.schliesse();
        app.scriptPreferences.userInteractionLevel = alteUILevel;

        this.zeigeErgebnis(titel, erfolge, fehlerListe);
    };

    /**
     * Zeigt das Ergebnis an
     *
     * @param {string} titel - Aktion-Titel
     * @param {number} erfolge - Anzahl Erfolge
     * @param {Array} fehlerListe - Fehler-Liste
     */
    this.zeigeErgebnis = function(titel, erfolge, fehlerListe) {
        var nachricht = "Aktion: " + titel + "\n\n";
        nachricht += "Erfolgreich: " + erfolge + "\n";
        nachricht += "Fehler: " + fehlerListe.length + "\n";

        if (fehlerListe.length > 0) {
            nachricht += "\n" + this.errorHandler.erstelleFehlerReport(fehlerListe);
        }

        alert(nachricht, SKRIPT_NAME + " - Ergebnis");
        this.logger.log("Relink abgeschlossen: " + erfolge + " erfolgreich, " + fehlerListe.length + " Fehler");
    };
}
/**
 * ValidationWorkflow - Katalog-Validierung
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {DocumentService} dependencies.documentService
 * @param {CsvService} dependencies.csvService
 * @param {DateUtils} dependencies.dateUtils
 * @param {Logger} dependencies.logger
 */
function ValidationWorkflow(dependencies) {
    this.documentService = dependencies.documentService;
    this.csvService = dependencies.csvService;
    this.dateUtils = dependencies.dateUtils;
    this.logger = dependencies.logger;

    /**
     * Führt die Validierung aus
     */
    this.fuehreAus = function() {
        this.logger.log("Starte Katalog-Validierung");

        var doc = app.activeDocument;
        var alleLinks = doc.links;
        var reportDaten = [];
        var toleranzInMs = DATUM_TOLERANZ_MINUTEN * 60 * 1000;

        for (var i = 0; i < alleLinks.length; i++) {
            var link = alleLinks[i];
            var linkDatei = File(link.filePath);
            var bemerkung = "";
            var seitenInfo = "Montagefläche";

            if (link.parent.parentPage) {
                var seitenName = link.parent.parentPage.name;
                var seitenNummer = parseInt(seitenName, 10);
                if (!isNaN(seitenNummer)) {
                    seitenInfo = seitenName;
                }
            }

            var statusInfo = "OK";
            if (link.status === LinkStatus.LINK_MISSING) {
                statusInfo = "Fehlend";
            } else if (link.status === LinkStatus.LINK_OUT_OF_DATE) {
                statusInfo = "Modifiziert";
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

                    if (istIndd && paarDatum.getTime() < linkDatum.getTime() - toleranzInMs) {
                        bemerkung = "Das zugehörige PDF ist möglicherweise veraltet.";
                    } else if (!istIndd && linkDatum.getTime() < paarDatum.getTime() - toleranzInMs) {
                        bemerkung = "PDF ist möglicherweise veraltet (INDD ist neuer).";
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
            alert("Keine auswertbaren Verknüpfungen gefunden.", SKRIPT_NAME);
            return;
        }

        this.erstelleReport(reportDaten);
    };

    /**
     * Erstellt den CSV-Report
     *
     * @param {Array} reportDaten - Report-Daten
     */
    this.erstelleReport = function(reportDaten) {
        var csvInhalt = "Seite im Katalog;Link-Name;Link-Typ;Status/Details;Dateipfad\n";

        for (var i = 0; i < reportDaten.length; i++) {
            var zeile = reportDaten[i];
            csvInhalt += '"' + zeile.seite + '";' +
                         '"' + zeile.name + '";' +
                         '"' + zeile.typ + '";' +
                         '"' + zeile.status + '";' +
                         '"' + zeile.pfad + '"\n';
        }

        var doc = app.activeDocument;
        var zeitstempel = this.dateUtils.formatiereZeitstempel(new Date());
        var reportName = doc.name.replace(/\.indd$/i, "") + "_validierung_" + zeitstempel + ".csv";

        var reportDatei = new File(doc.filePath + "/" + reportName);
        reportDatei.encoding = "UTF-8";
        reportDatei.open("w");
        reportDatei.write(csvInhalt);
        reportDatei.close();

        this.logger.log("Validierungs-Report erstellt: " + reportDatei.fsName);
        alert("Validierung abgeschlossen!\n\nReport gespeichert unter:\n" + reportDatei.fsName, SKRIPT_NAME);
    };
}
/**
 * CsvExportWorkflow - CSV-Export mit Auto-Detection
 *
 * @param {Object} dependencies - Abhängigkeiten
 * @param {ValidationEngine} dependencies.validationEngine
 * @param {ExportCoordinator} dependencies.exportCoordinator
 * @param {CsvService} dependencies.csvService
 * @param {DialogFactory} dependencies.dialogFactory
 * @param {DateUtils} dependencies.dateUtils
 * @param {Logger} dependencies.logger
 */
function CsvExportWorkflow(dependencies) {
    this.validationEngine = dependencies.validationEngine;
    this.exportCoordinator = dependencies.exportCoordinator;
    this.csvService = dependencies.csvService;
    this.dialogFactory = dependencies.dialogFactory;
    this.dateUtils = dependencies.dateUtils;
    this.logger = dependencies.logger;

    /**
     * Führt den CSV-Export-Workflow aus
     */
    this.fuehreAus = function() {
        this.logger.log("Starte CSV-Export-Workflow");

        var doc = app.activeDocument;
        var validierung = this.validationEngine.validiereDokument(doc);

        if (validierung.modus === EXPORT_MODE_ERROR) {
            alert("Keine relevanten Links für Export gefunden.", SKRIPT_NAME);
            return;
        }

        this.logger.log("Erkannter Modus: " + validierung.modus);

        var exportPlan = this.erstelleExportPlan(validierung);

        if (!exportPlan || exportPlan.gesamtAnzahl === 0) {
            alert("Keine Dateien zum Exportieren gefunden.", SKRIPT_NAME);
            return;
        }

        var nachricht = "Export-Modus: " + validierung.modus + "\n\n";
        nachricht += "Zu verarbeitende Dateien: " + exportPlan.gueltigeAnzahl + "\n";
        nachricht += "Fehlende Dateien: " + exportPlan.fehlendeAnzahl + "\n\n";
        nachricht += "Möchtest Du den Export starten?";

        if (!this.dialogFactory.erstelleBestaetigungsDialog(nachricht, "CSV Export - Bestätigung")) {
            this.logger.log("Export vom Benutzer abgebrochen");
            return;
        }

        var settings = this.erstelleStandardSettings();

        var ergebnis = this.exportCoordinator.koordiniereExport(exportPlan, settings);

        if (ergebnis.daten.length > 0) {
            this.schreibeCsv(ergebnis.daten, doc);
        }

        this.zeigeErgebnis(ergebnis);
    };

    /**
     * Erstellt einen Export-Plan aus der Validierung
     *
     * @param {Object} validierung - Validierungsergebnis
     * @returns {Object} ExportPlan
     */
    this.erstelleExportPlan = function(validierung) {
        var items = [];

        if (validierung.modus === EXPORT_MODE_DIRECT || validierung.modus === EXPORT_MODE_MIXED) {
            for (var i = 0; i < validierung.inddLinks.length; i++) {
                var inddLink = validierung.inddLinks[i];
                if (inddLink.existiert) {
                    items.push(erstelleExportItem(
                        EXPORT_MODE_DIRECT,
                        File(inddLink.dateiPfad),
                        null,
                        inddLink.elternSeite,
                        inddLink.link
                    ));
                }
            }
        }

        if (validierung.modus === EXPORT_MODE_CONVERT || validierung.modus === EXPORT_MODE_MIXED) {
            for (var j = 0; j < validierung.mappings.length; j++) {
                var mapping = validierung.mappings[j];
                if (mapping.gefunden) {
                    items.push(erstelleExportItem(
                        EXPORT_MODE_CONVERT,
                        mapping.inddDatei,
                        mapping.pdfLink.dateiName,
                        mapping.pdfLink.elternSeite,
                        mapping.pdfLink.link
                    ));
                }
            }
        }

        return erstelleExportPlan(validierung.modus, items);
    };

    /**
     * Erstellt Standard-Einstellungen
     *
     * @returns {Object} Settings-Objekt
     */
    this.erstelleStandardSettings = function() {
        return {
            mode: EXPORT_MODE_DIRECT,
            layer: {
                auto: STANDARD_LAYER_AUTO,
                manual: STANDARD_LAYER_MANUAL,
                name: STANDARD_LAYER_NAME
            },
            export: {
                graphics: STANDARD_EXPORT_GRAPHICS,
                textFrames: STANDARD_EXPORT_TEXTFRAMES,
                textInStories: STANDARD_EXPORT_TEXT_IN_STORIES,
                tables: STANDARD_EXPORT_TABLES,
                pageItems: STANDARD_EXPORT_PAGEITEMS
            },
            regex: {
                table: REGEX_TABLE_IN_STORY,
                textFrame: REGEX_TEXTFRAME_IN_STORY
            },
            csv: {
                single: STANDARD_CSV_EINZELN,
                multiple: STANDARD_CSV_MEHRFACH
            },
            sourceInfo: STANDARD_SOURCE_INFO
        };
    };

    /**
     * Schreibt die CSV-Datei(en)
     *
     * @param {Array} daten - CSV-Daten
     * @param {Document} doc - Das Dokument
     */
    this.schreibeCsv = function(daten, doc) {
        var ordner = Folder(doc.filePath);
        var zeitstempel = this.dateUtils.formatiereZeitstempel(new Date());
        var dateiName = doc.name.replace(/\.indd$/i, "") + "_export_" + zeitstempel + ".csv";

        var csvDatei = this.csvService.schreibeEinzelneCsv(daten, ordner, dateiName);
        this.logger.log("CSV erstellt: " + csvDatei.fsName);
    };

    /**
     * Zeigt das Ergebnis an
     *
     * @param {Object} ergebnis - Export-Ergebnis
     */
    this.zeigeErgebnis = function(ergebnis) {
        var nachricht = "CSV-Export abgeschlossen!\n\n";
        nachricht += "Erfolgreich verarbeitet: " + ergebnis.erfolgreich + "\n";
        nachricht += "Fehler: " + ergebnis.fehlgeschlagen + "\n";

        if (ergebnis.fehler.length > 0) {
            nachricht += "\nFehlerhafte Dateien:\n";
            for (var i = 0; i < Math.min(ergebnis.fehler.length, 5); i++) {
                nachricht += "- " + ergebnis.fehler[i].betroffeneDatei + "\n";
            }
            if (ergebnis.fehler.length > 5) {
                nachricht += "... und " + (ergebnis.fehler.length - 5) + " weitere\n";
            }
        }

        alert(nachricht, SKRIPT_NAME + " - Export Ergebnis");
    };
}

// ============================================
// MAIN EXECUTION
// ============================================
(function() {
    /**
     * Dependency Container
     */
    function erstelleDependencyContainer() {
        var container = {};

        container.stringUtils = new StringUtils();
        container.dateUtils = new DateUtils();
        container.regexHelper = new RegexHelper();
        container.errorHandler = new ErrorHandler();
        container.logger = new Logger();

        container.fileService = new FileService({
            logger: container.logger
        });

        container.documentService = new DocumentService({
            logger: container.logger,
            errorHandler: container.errorHandler
        });

        container.csvService = new CsvService({
            dateUtils: container.dateUtils,
            logger: container.logger
        });

        container.graphicsProcessor = new GraphicsProcessor({
            stringUtils: container.stringUtils,
            logger: container.logger
        });

        container.textProcessor = new TextProcessor({
            regexHelper: container.regexHelper,
            logger: container.logger
        });

        container.tableProcessor = new TableProcessor({
            textProcessor: container.textProcessor,
            regexHelper: container.regexHelper,
            logger: container.logger
        });

        container.pageItemsProcessor = new PageItemsProcessor({
            stringUtils: container.stringUtils,
            logger: container.logger
        });

        container.inddProcessor = new InddProcessor({
            graphicsProcessor: container.graphicsProcessor,
            textProcessor: container.textProcessor,
            tableProcessor: container.tableProcessor,
            pageItemsProcessor: container.pageItemsProcessor,
            documentService: container.documentService,
            logger: container.logger
        });

        container.dialogFactory = new DialogFactory();

        container.validationEngine = new ValidationEngine({
            fileService: container.fileService,
            logger: container.logger
        });

        container.exportCoordinator = new ExportCoordinator({
            documentService: container.documentService,
            inddProcessor: container.inddProcessor,
            csvService: container.csvService,
            dialogFactory: container.dialogFactory,
            logger: container.logger,
            errorHandler: container.errorHandler
        });

        container.relinkWorkflow = new RelinkWorkflow({
            fileService: container.fileService,
            documentService: container.documentService,
            dialogFactory: container.dialogFactory,
            errorHandler: container.errorHandler,
            logger: container.logger
        });

        container.validationWorkflow = new ValidationWorkflow({
            documentService: container.documentService,
            csvService: container.csvService,
            dateUtils: container.dateUtils,
            logger: container.logger
        });

        container.csvExportWorkflow = new CsvExportWorkflow({
            validationEngine: container.validationEngine,
            exportCoordinator: container.exportCoordinator,
            csvService: container.csvService,
            dialogFactory: container.dialogFactory,
            dateUtils: container.dateUtils,
            logger: container.logger
        });

        container.linkManager = new LinkManager({
            validationEngine: container.validationEngine,
            exportCoordinator: container.exportCoordinator,
            dialogFactory: container.dialogFactory,
            relinkWorkflow: container.relinkWorkflow,
            validationWorkflow: container.validationWorkflow,
            csvExportWorkflow: container.csvExportWorkflow,
            logger: container.logger,
            errorHandler: container.errorHandler
        });

        return container;
    }

    try {
        var container = erstelleDependencyContainer();
        container.linkManager.starte();
    } catch (e) {
        alert("Kritischer Fehler:\n\n" + e.message + "\n\nZeile: " + e.line,
              SKRIPT_NAME + " - Fehler");
    }
})();
