/**
 * Artikel Liste Export Script
 * Version: 1.5
 * Datum: 2025-10-29
 *
 * Durchsucht verknüpfte InDesign- und PDF-Dokumente und extrahiert Grafikinformationen.
 * Unterstützt CSV- und JSON-Export mit Ebenenauswahl und Preflight-Prüfung.
 */

// JSON Polyfill für ExtendScript (Minified) - Public Domain
if(typeof JSON!=="object"){JSON={}}(function(){"use strict";var rx_one=/^[\],:{}\s]*$/;var rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g;var rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g;var rx_four=/(?:^|:|,)(?:\s*\[)+/g;var rx_escapable=/[\\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;var rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;function f(n){return n<10?"0"+n:n}function this_value(){return this.valueOf()}if(typeof Date.prototype.toJSON!=="function"){Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z":null};Boolean.prototype.toJSON=this_value;Number.prototype.toJSON=this_value;String.prototype.toJSON=this_value}var gap;var indent;var meta={"\b":"\\b","\t":"\\t","\n":"\\n","\f":"\\f","\r":"\\r",'"':'\\"',"\\":"\\\\"};var rep;function quote(string){rx_escapable.lastIndex=0;return rx_escapable.test(string)?'"'+string.replace(rx_escapable,function(a){var c=meta[a];return typeof c==="string"?c:"\\u"+("0000"+a.charCodeAt(0).toString(16)).slice(-4)})+'"':'"'+string+'"'}function str(key,holder){var i;var k;var v;var length;var mind=gap;var partial;var value=holder[key];if(value&&typeof value==="object"&&typeof value.toJSON==="function"){value=value.toJSON(key)}if(typeof rep==="function"){value=rep.call(holder,key,value)}switch(typeof value){case"string":return quote(value);case"number":return isFinite(value)?String(value):"null";case"boolean":case"null":return String(value);case"object":if(!value){return"null"}gap+=indent;partial=[];if(Object.prototype.toString.apply(value)==="[object Array]"){length=value.length;for(i=0;i<length;i+=1){partial[i]=str(i,value)||"null"}v=partial.length===0?"[]":gap?"[\n"+gap+partial.join(",\n"+gap)+"\n"+mind+"]":"["+partial.join(",")+"]";gap=mind;return v}if(rep&&typeof rep==="object"){length=rep.length;for(i=0;i<length;i+=1){if(typeof rep[i]==="string"){k=rep[i];v=str(k,value);if(v){partial.push(quote(k)+(gap?": ":":")+v)}}}}else{for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=str(k,value);if(v){partial.push(quote(k)+(gap?": ":":")+v)}}}}v=partial.length===0?"{}":gap?"{\n"+gap+partial.join(",\n"+gap)+"\n"+mind+"}":"{"+partial.join(",")+"}";gap=mind;return v}}if(typeof JSON.stringify!=="function"){JSON.stringify=function(value,replacer,space){var i;gap="";indent="";if(typeof space==="number"){for(i=0;i<space;i+=1){indent+=" "}}else if(typeof space==="string"){indent=space}rep=replacer;if(replacer&&typeof replacer!=="function"&&(typeof replacer!=="object"||typeof replacer.length!=="number")){throw new Error("JSON.stringify")}return str("",{"":value})}}if(typeof JSON.parse!=="function"){JSON.parse=function(text,reviver){var j;function walk(holder,key){var k;var v;var value=holder[key];if(value&&typeof value==="object"){for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=walk(value,k);if(v!==undefined){value[k]=v}else{delete value[k]}}}}return reviver.call(holder,key,value)}text=String(text);rx_dangerous.lastIndex=0;if(rx_dangerous.test(text)){text=text.replace(rx_dangerous,function(a){return"\\u"+("0000"+a.charCodeAt(0).toString(16)).slice(-4)})}if(rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,""))){j=eval("("+text+")");return typeof reviver==="function"?walk({"":j},""):j}throw new SyntaxError("JSON.parse")}}}());

/**
 * Hauptfunktion des Scripts
 */
var main = function() {
    var SKRIPT_VERSION = "1.5";
    var SKRIPT_DATUM = "2025-10-29";
    var SKRIPT_NAME = "Artikel Liste Export";

    var doc = app.activeDocument;
    var links = doc.links;
    var ausgabeText = [["Ebene", "Seite", "Datei", "Produktnummer", "Format"]];
    var filePath = null;
    var exportFormat = "csv"; // "csv" oder "json"
    var selectedLayer = null;

    $.writeln("=== " + SKRIPT_NAME + " v" + SKRIPT_VERSION + " ===");

    /**
     * Zeigt den Hauptdialog zur Auswahl von Ebene und Exportformat an.
     * @returns {Object|null} Benutzereingaben oder null bei Abbruch
     */
    function zeigeHauptdialog() {
        // Alle Ebenen im Dokument sammeln
        var layerNamen = [];
        for (var i = 0; i < doc.layers.length; i++) {
            layerNamen.push(doc.layers[i].name);
        }

        if (layerNamen.length === 0) {
            alert("Keine Ebenen im Dokument gefunden.", SKRIPT_NAME);
            return null;
        }

        var dialog = new Window("dialog", SKRIPT_NAME + " v" + SKRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 10;
        dialog.margins = 15;

        // Layer-Auswahl Panel
        var layerPanel = dialog.add("panel", undefined, "Ebene für den Export auswählen");
        layerPanel.alignChildren = "fill";
        layerPanel.margins = 10;

        var layerDropdown = layerPanel.add("dropdownlist", undefined, layerNamen);
        layerDropdown.preferredSize.width = 350;

        // Standard-Ebene auswählen (DA + Seiten)
        var standardAuswahlIndex = 0;
        for (var i = 0; i < layerNamen.length; i++) {
            if (layerNamen[i].indexOf("DA") > -1 && layerNamen[i].indexOf("Seiten") > -1) {
                standardAuswahlIndex = i;
                break;
            }
        }
        layerDropdown.selection = standardAuswahlIndex;

        // Ausgabeformat Panel
        var formatPanel = dialog.add("panel", undefined, "Ausgabeformat");
        formatPanel.alignChildren = "left";
        formatPanel.margins = 10;

        var csvRadio = formatPanel.add("radiobutton", undefined, "CSV");
        var jsonRadio = formatPanel.add("radiobutton", undefined, "JSON");
        csvRadio.value = true; // CSV ist Standard

        // Button-Gruppe
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["fill", "center"];
        buttonGroup.spacing = 10;

        var exportBtn = buttonGroup.add("button", undefined, "Export Data", { name: "ok" });
        var abbrechenBtn = buttonGroup.add("button", undefined, "Abbrechen", { name: "cancel" });

        if (dialog.show() === 1) {
            if (!layerDropdown.selection) {
                alert("Bitte wähle eine Ebene aus.", SKRIPT_NAME);
                return null;
            }

            return {
                layerName: layerDropdown.selection.text,
                format: csvRadio.value ? "csv" : "json"
            };
        }

        return null;
    }

    /**
     * Führt eine Preflight-Prüfung für alle Links im Dokument durch.
     * Prüft auf fehlende, veraltete und nicht erreichbare Verknüpfungen.
     * @returns {Object} Ergebnis mit { erfolg: boolean, fehler: Array }
     */
    function fuehrePreflight() {
        $.writeln("Starte Preflight-Prüfung...");
        var fehlerListe = [];

        for (var i = 0; i < links.length; i++) {
            var link = links[i];
            var linkName = link.name || "Unbenannter Link";

            if (link.status === LinkStatus.LINK_MISSING) {
                fehlerListe.push(linkName + ": Verknüpfung fehlt");
            } else if (link.status === LinkStatus.LINK_OUT_OF_DATE) {
                fehlerListe.push(linkName + ": Verknüpfung ist veraltet");
            } else if (link.status === LinkStatus.LINK_INACCESSIBLE) {
                fehlerListe.push(linkName + ": Verknüpfung nicht erreichbar");
            }
        }

        return {
            erfolg: fehlerListe.length === 0,
            fehler: fehlerListe
        };
    }

    /**
     * Zeigt eine Preflight-Fehlermeldung an.
     * @param {Array} fehlerListe Liste der gefundenen Fehler
     */
    function zeigePreflightFehler(fehlerListe) {
        var nachricht = "⚠ Preflight-Fehler gefunden!\n\n";
        nachricht += "Folgende Links sind problematisch:\n\n";

        var maxAnzeige = Math.min(fehlerListe.length, 10);
        for (var i = 0; i < maxAnzeige; i++) {
            nachricht += "• " + fehlerListe[i] + "\n";
        }

        if (fehlerListe.length > maxAnzeige) {
            nachricht += "\n... und " + (fehlerListe.length - maxAnzeige) + " weitere Fehler.\n";
        }

        nachricht += "\nBitte aktualisiere die Verknüpfungen und versuche es erneut.";

        alert(nachricht, SKRIPT_NAME + " - Preflight-Fehler");
    }

    /**
     * Prüft, ob auf der ausgewählten Ebene gesperrte Objekte vorhanden sind.
     * @param {string} layerName Name der zu prüfenden Ebene
     * @returns {Object} Ergebnis mit { erfolg: boolean, anzahl: number, liste: Array }
     */
    function pruefeGesperrteObjekte(layerName) {
        $.writeln("Prüfe auf gesperrte Objekte auf Ebene '" + layerName + "'...");
        var gesperrteObjekte = 0;
        var gesperrteListe = [];

        try {
            var layer = doc.layers.itemByName(layerName);
            if (!layer || !layer.isValid) {
                $.writeln("WARNUNG: Ebene '" + layerName + "' nicht gefunden.");
                return { erfolg: true, anzahl: 0, liste: [] };
            }

            // Durch alle Seiten iterieren
            for (var p = 0; p < doc.pages.length; p++) {
                var page = doc.pages[p];

                // Durch alle PageItems auf der Seite iterieren
                for (var i = 0; i < page.allPageItems.length; i++) {
                    var item = page.allPageItems[i];

                    // Prüfen, ob das Item auf der ausgewählten Ebene liegt
                    try {
                        if (item.itemLayer && item.itemLayer.name === layerName) {
                            // Prüfen, ob das Item gesperrt ist
                            if (item.locked === true) {
                                gesperrteObjekte++;

                                // Name des Objekts ermitteln
                                var objektName = item.name || item.label || item.constructor.name || "Unbenanntes Objekt";
                                var seitenName = page.name || "Unbekannt";

                                gesperrteListe.push({
                                    name: objektName,
                                    seite: seitenName,
                                    typ: item.constructor.name
                                });

                                $.writeln("Gesperrtes Objekt gefunden auf Seite " + seitenName + ": " + objektName + " (" + item.constructor.name + ")");
                            }
                        }
                    } catch (e) {
                        // Fehler beim Zugriff auf einzelnes Item ignorieren
                    }
                }
            }

        } catch (e) {
            $.writeln("FEHLER bei der Prüfung gesperrter Objekte: " + e.toString());
        }

        $.writeln("Gesperrte Objekte gefunden: " + gesperrteObjekte);
        return {
            erfolg: gesperrteObjekte === 0,
            anzahl: gesperrteObjekte,
            liste: gesperrteListe
        };
    }

    /**
     * Zeigt eine Fehlermeldung für gesperrte Objekte an.
     * @param {Object} ergebnis Ergebnis-Objekt mit anzahl und liste
     */
    function zeigeGesperrteObjekteFehler(ergebnis) {
        var nachricht = "⚠ Gesperrte Objekte gefunden!\n\n";
        nachricht += "Auf der gewählten Ebene befinden sich " + ergebnis.anzahl + " gesperrte(s) Objekt(e).\n\n";

        // Liste der gesperrten Objekte anzeigen
        if (ergebnis.liste && ergebnis.liste.length > 0) {
            var maxAnzeige = Math.min(ergebnis.liste.length, 15);
            for (var i = 0; i < maxAnzeige; i++) {
                var obj = ergebnis.liste[i];
                nachricht += "• Seite " + obj.seite + ": " + obj.name + "\n";
            }

            if (ergebnis.liste.length > maxAnzeige) {
                nachricht += "\n... und " + (ergebnis.liste.length - maxAnzeige) + " weitere.\n";
            }
            nachricht += "\n";
        }

        nachricht += "Das Script kann gesperrte Objekte nicht auslesen.\n";
        nachricht += "Bitte entsperre alle Objekte auf der Ebene und versuche es erneut.";

        alert(nachricht, SKRIPT_NAME + " - Gesperrte Objekte");
    }

    /**
     * Ermittelt die Elternseite eines Links.
     * @param {Link} link Der zu untersuchende Link
     * @returns {Page|null} Die Elternseite oder null
     */
    function getParentPage(link) {
        try {
            var parent = link.parent;
            while (parent != null) {
                if (parent.hasOwnProperty('parentPage') && parent.parentPage != null && parent.parentPage.isValid) {
                    return parent.parentPage;
                }
                if (parent.constructor.name === "Spread") {
                    return null;
                }
                parent = parent.parent;
            }
        } catch (e) {
            $.writeln("Fehler bei der Ermittlung der Elternseite für Link ID " + link.id + ": " + e.toString().substring(0,100));
        }
        return null;
    }

    /**
     * Durchsucht alle Links auf der ausgewählten Ebene.
     * @param {string} layerName Name der zu durchsuchenden Ebene
     */
    function suche(layerName) {
        var linkCount = 0;

        for (var i = 0; i < links.length; i++) {
            var link = links[i];

            // Variablen für die Ausgabezeile initialisieren
            var ebenenName = "Nicht ermittelbar";
            var seitenName = "Nicht ermittelbar";
            var dateiNameFuerReport = "Nicht ermittelbar";
            var produktNummerFuerReport = "N/A";
            var formatFuerReport = "N/A";

            // Grundlegende Link-Informationen ermitteln
            try {
                var parentPageObj = getParentPage(link);
                if (parentPageObj && parentPageObj.isValid && parentPageObj.name) {
                    seitenName = parentPageObj.name;
                } else if (parentPageObj === null && link.parent && link.parent.constructor.name === "Spread") {
                    seitenName = "Montagefläche";
                } else {
                    seitenName = "Seite nicht gefunden";
                }

                if (link.parent && link.parent.itemLayer && link.parent.itemLayer.isValid) {
                    ebenenName = link.parent.itemLayer.name;
                } else {
                    ebenenName = "Ebene nicht gefunden";
                }
            } catch (e_grundinfo) {
                $.writeln("DEBUG: Fehler bei Ermittlung von Ebene/Seite für Link ID " + link.id + ": " + e_grundinfo.toString().substring(0,100));
                ebenenName = "Fehler Ebene/Seite";
                seitenName = "";
            }

            // Prüfen, ob Link auf der ausgewählten Ebene liegt
            if (ebenenName !== layerName) {
                continue; // Überspringen, wenn nicht auf der gewählten Ebene
            }

            // Dateiname für die Ausgabe ermitteln
            if (link.filePath) {
                dateiNameFuerReport = decodeURI(new File(link.filePath).name).replace(/%20/g, ' ');
            } else if (link.name) {
                dateiNameFuerReport = decodeURI(link.name).replace(/%20/g, ' ');
            } else {
                dateiNameFuerReport = "Dateiname unbekannt";
            }

            var istValiderInDesignPfad = (link.filePath && link.filePath.match(/\.indd$/i));
            var istValiderPdfPfad = (link.filePath && link.filePath.match(/\.pdf$/i));

            if (istValiderInDesignPfad || istValiderPdfPfad) {
                try {
                    var linkedDocFile = new File(link.filePath);

                    if (linkedDocFile.exists) {
                        var produktNrArray = extractSixDigitBlocks(dateiNameFuerReport);

                        // Prüfe die Höhe des Rahmens
                        if (link.parent && link.parent.geometricBounds &&
                            typeof link.parent.geometricBounds[0] === 'number' &&
                            typeof link.parent.geometricBounds[2] === 'number') {
                            var linkFrameHeight = link.parent.geometricBounds[2] - link.parent.geometricBounds[0];
                            formatFuerReport = linkFrameHeight < 240 ? "0,5" : "1";
                        } else {
                            formatFuerReport = "Formatfehler";
                            $.writeln("DEBUG: Fehler beim Ermitteln der Rahmenhöhe für Link: " + dateiNameFuerReport + " auf Seite " + seitenName);
                        }

                        if (produktNrArray.length > 0) {
                            for (var e = 0; e < produktNrArray.length; e++) {
                                produktNummerFuerReport = produktNrArray[e];
                                ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                                linkCount++;
                            }
                        } else {
                            produktNummerFuerReport = "Keine Produktnr.";
                            ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                            linkCount++;
                        }

                    } else {
                        var dateiTyp = istValiderInDesignPfad ? "InDesign" : "PDF";
                        produktNummerFuerReport = "FEHLER: " + dateiTyp + "-Datei nicht gefunden";
                        formatFuerReport = "N/A";
                        $.writeln("FEHLER: Datei existiert nicht: " + linkedDocFile.fullName + " (Verknüpft auf Seite: " + seitenName + ")");
                        ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                        linkCount++;
                    }
                } catch (e_verarbeitung) {
                    produktNummerFuerReport = "FEHLER: Verarbeitung";
                    formatFuerReport = "N/A";
                    $.writeln("FEHLER: Bei der Verarbeitung der Datei: " + dateiNameFuerReport + " (Verknüpft auf Seite: " + seitenName + ") - Details: " + e_verarbeitung.toString().substring(0,100));
                    ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                    linkCount++;
                }
            } else {
                // Kein gültiger InDesign- oder PDF-Pfad
                var sollteInDesignSeinAnhandName = (link.name && link.name.match(/\.indd$/i));
                var istBekanntFehlendOderNichtZugreifbarAlsINDD =
                    (link.status === LinkStatus.LINK_MISSING || link.status === LinkStatus.LINK_INACCESSIBLE) &&
                    (typeof link.linkType === 'string' && link.linkType.match(/indesign/i));

                var solltePdfSeinAnhandName = (link.name && link.name.match(/\.pdf$/i));
                var istBekanntFehlendOderNichtZugreifbarAlsPDF =
                    (link.status === LinkStatus.LINK_MISSING || link.status === LinkStatus.LINK_INACCESSIBLE) &&
                    (typeof link.linkType === 'string' && link.linkType.match(/pdf/i));

                if (sollteInDesignSeinAnhandName || istBekanntFehlendOderNichtZugreifbarAlsINDD ||
                    solltePdfSeinAnhandName || istBekanntFehlendOderNichtZugreifbarAlsPDF) {

                    var problemTyp = "";
                    if (sollteInDesignSeinAnhandName || istBekanntFehlendOderNichtZugreifbarAlsINDD) {
                        problemTyp = "InDesign";
                    } else {
                        problemTyp = "PDF";
                    }

                    if (link.status === LinkStatus.LINK_MISSING) {
                        produktNummerFuerReport = "FEHLER: Link fehlt (" + problemTyp + ")";
                    } else if (link.status === LinkStatus.LINK_INACCESSIBLE) {
                        produktNummerFuerReport = "FEHLER: Link nicht zugreifbar (" + problemTyp + ")";
                    } else if ((sollteInDesignSeinAnhandName && (!link.filePath || !link.filePath.match(/\.indd$/i))) ||
                               (solltePdfSeinAnhandName && (!link.filePath || !link.filePath.match(/\.pdf$/i))) ) {
                        produktNummerFuerReport = "FEHLER: Pfad falsch/fehlt (" + problemTyp + "-Name)";
                    } else {
                        produktNummerFuerReport = "FEHLER: Unbek. " + problemTyp + "-Linkproblem";
                    }
                    formatFuerReport = "N/A";
                    $.writeln("FEHLER: Problematischer " + problemTyp + "-Link erfasst: " + dateiNameFuerReport + " (Verknüpft auf Seite: " + seitenName + ", Grund: " + produktNummerFuerReport +")");
                    ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                    linkCount++;
                }
            }
        }

        $.writeln("Verarbeitung abgeschlossen. " + linkCount + " Links auf Ebene '" + layerName + "' gefunden.");
        return linkCount;
    }

    /**
     * Extrahiert 6-stellige Zahlenblöcke aus einem String.
     * @param {string} inputString Der zu durchsuchende String
     * @returns {Array} Array mit gefundenen 6-stelligen Zahlen
     */
    function extractSixDigitBlocks(inputString) {
        var matches = [];
        if (typeof inputString !== 'string') {
            return matches;
        }
        var parts = inputString.split(/[\s\+\-_]/);
        for (var i = 0; i < parts.length; i++) {
            if (/^\d{6}$/.test(parts[i])) {
                matches.push(parts[i]);
            }
        }
        return matches;
    }

    /**
     * Generiert einen Zeitstempel im Format YYMMDD_HHMM.
     * @returns {string} Formatierter Zeitstempel
     */
    function generiereZeitstempel() {
        var jetzt = new Date();
        var jahr = ("0" + (jetzt.getFullYear() % 100)).slice(-2);
        var monat = ("0" + (jetzt.getMonth() + 1)).slice(-2);
        var tag = ("0" + jetzt.getDate()).slice(-2);
        var stunde = ("0" + jetzt.getHours()).slice(-2);
        var minute = ("0" + jetzt.getMinutes()).slice(-2);
        return jahr + monat + tag + "_" + stunde + minute;
    }

    /**
     * Generiert einen ISO-formatierten Datums-String (ExtendScript-kompatibel).
     * @param {Date} datum Das zu formatierende Datum
     * @returns {string} ISO-formatierter String (YYYY-MM-DDTHH:MM:SSZ)
     */
    function generiereISODatum(datum) {
        var jahr = datum.getFullYear();
        var monat = ("0" + (datum.getMonth() + 1)).slice(-2);
        var tag = ("0" + datum.getDate()).slice(-2);
        var stunde = ("0" + datum.getHours()).slice(-2);
        var minute = ("0" + datum.getMinutes()).slice(-2);
        var sekunde = ("0" + datum.getSeconds()).slice(-2);
        return jahr + "-" + monat + "-" + tag + "T" + stunde + ":" + minute + ":" + sekunde + "Z";
    }

    /**
     * Schreibt die Daten als CSV-Datei.
     * @param {File} file Die Zieldatei
     */
    function writeCSV(file) {
        if (!file) {
            $.writeln("FEHLER: Kein Speicherpfad für CSV vorhanden.");
            alert("Interner Fehler: Kein Speicherpfad für CSV vorhanden.", SKRIPT_NAME);
            return false;
        }

        file.encoding = "UTF-8";
        file.open("w");

        for (var i = 0; i < ausgabeText.length; i++) {
            var rowArray = ausgabeText[i];
            for (var j = 0; j < rowArray.length; j++) {
                if (typeof rowArray[j] === 'undefined' || rowArray[j] === null) {
                    rowArray[j] = "";
                }
            }
            var row = rowArray.join(";");
            file.writeln(row);
        }

        file.close();
        $.writeln("CSV-Datei erfolgreich gespeichert: " + file.fsName);
        return true;
    }

    /**
     * Schreibt die Daten als JSON-Datei mit Metadaten.
     * Items werden nach Seitennummer sortiert.
     * @param {File} file Die Zieldatei
     */
    function writeJSON(file) {
        if (!file) {
            $.writeln("FEHLER: Kein Speicherpfad für JSON vorhanden.");
            alert("Interner Fehler: Kein Speicherpfad für JSON vorhanden.", SKRIPT_NAME);
            return false;
        }

        try {
            // Items vorbereiten (Header überspringen)
            var items = [];
            for (var i = 1; i < ausgabeText.length; i++) {
                var row = ausgabeText[i];
                items.push({
                    "Ebene": row[0] || "",
                    "Seite": row[1] || "",
                    "Datei": row[2] || "",
                    "Produktnummer": row[3] || "",
                    "Format": row[4] || ""
                });
            }

            // Nach Seitennummer sortieren
            items.sort(function(a, b) {
                var seiteA = parseInt(a.Seite, 10);
                var seiteB = parseInt(b.Seite, 10);

                // NaN-Werte ans Ende
                if (isNaN(seiteA) && isNaN(seiteB)) return 0;
                if (isNaN(seiteA)) return 1;
                if (isNaN(seiteB)) return -1;

                return seiteA - seiteB;
            });

            var headerZeile = ausgabeText.length > 0 ? ausgabeText[0].slice(0) : ["Ebene", "Seite", "Datei", "Produktnummer", "Format"];
            var erforderlicheSpalten = ["Produktnummer", "Seite", "Format", "Datei"];
            var produktnummernZaehler = {};
            var seitenZaehler = {};
            var produktnummernAnzahl = 0;
            var seitenAnzahl = 0;

            for (var j = 0; j < items.length; j++) {
                var produktKey = items[j].Produktnummer || "";
                if (produktKey !== "" && !produktnummernZaehler.hasOwnProperty(produktKey)) {
                    produktnummernZaehler[produktKey] = true;
                    produktnummernAnzahl++;
                }
                var seitenKey = items[j].Seite || "";
                if (seitenKey !== "" && !seitenZaehler.hasOwnProperty(seitenKey)) {
                    seitenZaehler[seitenKey] = true;
                    seitenAnzahl++;
                }
            }

            var dokumentPfad = "";
            try {
                if (doc.saved && doc.fullName) {
                    dokumentPfad = doc.fullName.fsName;
                }
            } catch (pfadFehler) {
                dokumentPfad = "";
            }

            // JSON-Objekt mit Metadaten erstellen
            var jsonObjekt = {
                "version": SKRIPT_VERSION,
                "exported": generiereISODatum(new Date()),
                "document": doc.name,
                "documentPath": dokumentPfad,
                "layer": selectedLayer,
                "itemCount": items.length,
                "schema": {
                    "headers": headerZeile,
                    "required": erforderlicheSpalten,
                    "delimiter": ";"
                },
                "statistics": {
                    "uniqueProduktnummern": produktnummernAnzahl,
                    "uniqueSeiten": seitenAnzahl
                },
                "items": items
            };

            // JSON-String erzeugen
            var jsonString = JSON.stringify(jsonObjekt, null, 2);

            // Datei schreiben
            file.encoding = "UTF-8";
            file.open("w");
            file.write(jsonString);
            file.close();

            $.writeln("JSON-Datei erfolgreich gespeichert: " + file.fsName);
            return true;

        } catch (e) {
            $.writeln("FEHLER: JSON-Serialisierung fehlgeschlagen: " + e.toString());
            alert("Fehler beim Erstellen der JSON-Datei:\n" + e.message, SKRIPT_NAME + " - JSON-Fehler");
            return false;
        }
    }

    // === HAUPTABLAUF ===

    // 1. Hauptdialog anzeigen
    var benutzerEingabe = zeigeHauptdialog();
    if (!benutzerEingabe) {
        $.writeln("Vorgang abgebrochen.");
        return;
    }

    selectedLayer = benutzerEingabe.layerName;
    exportFormat = benutzerEingabe.format;

    $.writeln("Ausgewählte Ebene: " + selectedLayer);
    $.writeln("Ausgewähltes Format: " + exportFormat.toUpperCase());

    // 2. Preflight-Prüfung durchführen
    var preflightErgebnis = fuehrePreflight();
    if (!preflightErgebnis.erfolg) {
        zeigePreflightFehler(preflightErgebnis.fehler);
        $.writeln("Export abgebrochen wegen Preflight-Fehlern.");
        return;
    }

    $.writeln("Preflight erfolgreich. Alle Links sind in Ordnung.");

    // 3. Prüfen, ob gesperrte Objekte auf der Ebene vorhanden sind
    var gesperrteObjekteErgebnis = pruefeGesperrteObjekte(selectedLayer);
    if (!gesperrteObjekteErgebnis.erfolg) {
        zeigeGesperrteObjekteFehler(gesperrteObjekteErgebnis);
        $.writeln("Export abgebrochen wegen gesperrter Objekte.");
        return;
    }

    $.writeln("Keine gesperrten Objekte gefunden. Fortfahren mit Export.");

    // 4. Links auf der ausgewählten Ebene suchen
    var gefundeneLinks = suche(selectedLayer);

    // Prüfen, ob Ergebnisse gefunden wurden
    if (gefundeneLinks === 0) {
        var nachricht = "ℹ Keine Ergebnisse\n\n";
        nachricht += "Auf der gewählten Ebene '" + selectedLayer + "' wurden keine\n";
        nachricht += "InDesign- oder PDF-Verknüpfungen gefunden.";
        alert(nachricht, SKRIPT_NAME);
        $.writeln("Keine Ergebnisse gefunden.");
        return;
    }

    // 5. Speicherort wählen mit automatischem Dateinamen
    var zeitstempel = generiereZeitstempel();
    var dateiEndung = exportFormat === "json" ? ".json" : ".csv";
    var standardDateiName = "artikel_liste_" + zeitstempel + dateiEndung;

    // Preset mit vorgeschlagenem Dateinamen (nur Dateiname, kein vollständiger Pfad)
    filePath = File.saveDialog("Speichern unter", standardDateiName);

    if (filePath) {
        $.writeln("Speicherort gewählt: " + filePath.fsName);

        // 6. Datei schreiben
        var erfolg = false;
        if (exportFormat === "json") {
            erfolg = writeJSON(filePath);
        } else {
            erfolg = writeCSV(filePath);
        }

        // 7. Erfolgsmeldung
        if (erfolg) {
            var erfolgNachricht = exportFormat.toUpperCase() + "-Datei wurde erfolgreich gespeichert!\n\n";
            erfolgNachricht += "Datei: " + filePath.fsName + "\n";
            erfolgNachricht += "Einträge: " + (ausgabeText.length - 1);
            alert(erfolgNachricht, SKRIPT_NAME + " - Export erfolgreich");
        }
    } else {
        alert("Speichervorgang abgebrochen.", SKRIPT_NAME);
        $.writeln("Speichervorgang abgebrochen.");
    }
};

// Führe das Hauptskript aus
main();
