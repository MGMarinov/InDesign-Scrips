// Hauptskript zum Durchsuchen verknüpfter InDesign- und PDF-Dokumente und Extrahieren von Grafikinformationen
var main = function() {
    //allgemeine Variablen
    var doc = app.activeDocument; // Aktives Dokument
    var links = doc.links; // Alle Verknüpfungen im Dokument

    //Array enthält Infos zu Seite und INDD/PDF-Datei
    var ausgabeText = [
        ["Ebene", "Seite", "Datei", "Produktnummer", "Format"]];
    $.writeln("start");

    //Speicherort wählen
    var filePath = File.saveDialog("Speichern unter", "*.csv");

    if (filePath) { // Nur fortfahren, wenn ein Speicherort gewählt wurde
        $.writeln("Speicherort gewählt: " + filePath.fsName); // fsName für plattformkorrekten Pfad
        suche();
        writeCSV(); // CSV-Schreiben nach der Suche aufrufen
    } else {
        alert("Speichervorgang abgebrochen. Keine Datei wurde gespeichert.");
        $.writeln("Speichervorgang abgebrochen.");
    }

    function suche() {
        for (var i = 0; i < links.length; i++) {
            var link = links[i];

            // Variablen für die CSV-Zeile initialisieren
            var ebenenName = "Nicht ermittelbar";
            var seitenName = "Nicht ermittelbar";
            var dateiNameFuerReport = "Nicht ermittelbar";
            var produktNummerFuerReport = "N/A"; // Standardwert, wird bei Erfolg oder spezifischem Fehler überschrieben
            var formatFuerReport = "N/A"; // Standardwert

            // Grundlegende Link-Informationen (Ebene, Seite) versuchen zu ermitteln
            try {
                var parentPageObj = getParentPage(link); // getParentPage hat eigene Fehlerbehandlung
                if (parentPageObj && parentPageObj.isValid && parentPageObj.name) { // Zusätzliche Prüfung auf isValid
                    seitenName = parentPageObj.name;
                } else if (parentPageObj === null && link.parent && link.parent.constructor.name === "Spread") {
                    // Element ist direkt auf der Druckbogenmontagefläche (Spread), nicht auf einer Seite platziert
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
                // Fehler bei der Ermittlung der grundlegenden Informationen protokollieren
                $.writeln("DEBUG: Fehler bei Ermittlung von Ebene/Seite für Link ID " + link.id + " (" + (link.name || (link.filePath || "Unbenannter Link")) + "): " + e_grundinfo.toString().substring(0,100));
                ebenenName = "Fehler Ebene/Seite";
                seitenName = "";
            }

            // Dateiname für die Ausgabe ermitteln
            if (link.filePath) {
                // decodeURI, um URL-kodierte Zeichen (z.B. %20 für Leerzeichen) korrekt darzustellen
                dateiNameFuerReport = decodeURI(new File(link.filePath).name).replace(/%20/g, ' ');
            } else if (link.name) {
                // Fallback auf link.name, wenn filePath nicht vorhanden ist (typisch für fehlende Verknüpfungen)
                dateiNameFuerReport = decodeURI(link.name).replace(/%20/g, ' ');
            } else {
                dateiNameFuerReport = "Dateiname unbekannt";
            }

            var istValiderInDesignPfad = (link.filePath && link.filePath.match(/\.indd$/i));
            var istValiderPdfPfad = (link.filePath && link.filePath.match(/\.pdf$/i));

            if (istValiderInDesignPfad || istValiderPdfPfad) {
                // Verknüpfung ist ein InDesign- oder PDF-Dokument laut Pfad
                try {
                    var linkedDocFile = new File(link.filePath);

                    if (linkedDocFile.exists) {
                        // Erfolgreicher Fall: Datei existiert
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
                            }
                        } else {
                            // Keine Produktnummer gefunden, aber Link und Datei sind OK
                            produktNummerFuerReport = "Keine Produktnr.";
                            ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                        }

                    } else {
                        // Datei existiert nicht am Pfad
                        var dateiTyp = istValiderInDesignPfad ? "InDesign" : "PDF";
                        produktNummerFuerReport = "FEHLER: " + dateiTyp + "-Datei nicht gefunden";
                        formatFuerReport = "N/A";
                        $.writeln("FEHLER: Datei existiert nicht: " + linkedDocFile.fullName + " (Verknüpft auf Seite: " + seitenName + ")");
                        ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                    }
                } catch (e_verarbeitung) {
                    // Fehler bei der Verarbeitung der existierenden Datei oder anderer Fehler im try-Block
                    produktNummerFuerReport = "FEHLER: Verarbeitung";
                    formatFuerReport = "N/A";
                    $.writeln("FEHLER: Bei der Verarbeitung der Datei: " + dateiNameFuerReport + " (Verknüpft auf Seite: " + seitenName + ") - Details: " + e_verarbeitung.toString().substring(0,100));
                    ausgabeText.push([ebenenName, seitenName, dateiNameFuerReport, produktNummerFuerReport, formatFuerReport]);
                }
            } else {
                // Kein gültiger InDesign- oder PDF-Pfad (link.filePath ist null, leer oder zeigt nicht auf .indd/.pdf)
                // Prüfen, ob es sich um einen fehlerhaften Link handelt, der eigentlich InDesign oder PDF sein sollte
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
                    } else { // Muss PDF sein aufgrund der äußeren if-Bedingung
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
                } else {
                    // Reguläre nicht-InDesign/nicht-PDF Verknüpfung oder anderer nicht relevanter Fall, wird ignoriert.
                }
            }
        }
    } // Ende suche()

    // Funktion, um die Elternseite eines Links zu ermitteln
    function getParentPage(link) {
        try {
            var parent = link.parent;
            while (parent != null) {
                if (parent.hasOwnProperty('parentPage') && parent.parentPage != null && parent.parentPage.isValid) { // isValid Prüfung hinzugefügt
                    return parent.parentPage;
                }
                // Prüfen, ob das Elternelement direkt ein Spread ist (z.B. Element auf Montagefläche)
                if (parent.constructor.name === "Spread") {
                    return null; // Signalisiert, dass es auf der Montagefläche ist oder keine Seite hat
                }
                parent = parent.parent;
            }
        } catch (e) {
            // Gekürzte Fehlermeldung zur besseren Lesbarkeit in der Konsole
            $.writeln("Fehler bei der Ermittlung der Elternseite für Link ID " + link.id + ": " + e.toString().substring(0,100));
        }
        return null;
    }

    // Funktion zum Schreiben der CSV-Datei
    function writeCSV() {
        if (!filePath) {
            $.writeln("FEHLER: Kein Speicherpfad für CSV vorhanden beim Versuch zu schreiben.");
            alert("Interner Fehler: Kein Speicherpfad für CSV vorhanden.");
            return;
        }
        var file = new File(filePath);
        file.encoding = "UTF-8"; // UTF-8 Encoding setzen

        file.open("w");
        for (var i = 0; i < ausgabeText.length; i++) {
            var rowArray = ausgabeText[i];
            for (var j = 0; j < rowArray.length; j++) {
                if (typeof rowArray[j] === 'undefined' || rowArray[j] === null) {
                    rowArray[j] = ""; // Ersetze undefined/null durch leeren String für saubere CSV
                }
            }
            var row = rowArray.join(";"); // Zeile mit Semikolon trennen
            file.writeln(row); // Zeile in die Datei schreiben
        }
        file.close();
        alert("CSV-Datei wurde erfolgreich unter '" + filePath.fsName + "' gespeichert.");
    }

    function extractSixDigitBlocks(inputString) {
        var matches = [];
        // Sicherstellen, dass inputString ein String ist, um Fehler zu vermeiden
        if (typeof inputString !== 'string') {
            return matches;
        }
        // Splitte den Eingabestring nach Leerzeichen, Plus, Minus und Unterstrich
        var parts = inputString.split(/[\s\+\-_]/);
        for (var i = 0; i < parts.length; i++) {
            // Überprüfe, ob der Teil genau 6-stellig ist und nur aus Ziffern besteht
            if (/^\d{6}$/.test(parts[i])) {
                matches.push(parts[i]);
            }
        }
        return matches;
    }

}; // Ende main()

// Führe das Hauptskript aus
main();