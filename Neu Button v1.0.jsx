/*******************************************************************************
 *
 * Neu Button Skript
 *
 * Beschreibung:
 * Dieses Skript automatisiert die Erstellung von "Neu Button"-Textrahmen auf
 * allen Seiten eines InDesign-Dokuments. Es erkennt automatisch die Seitenhöhe
 * (ganze oder halbe Seite) und passt die vertikale Position der Objekte an.
 * Das Skript unterstützt zweisprachige Inhalte (Deutsch/Französisch) und wählt
 * je nach Seitenposition (links/rechts) den passenden Absatzstil aus. Die
 * Objekte werden auf einer dedizierten "Neu"-Ebene platziert. Eine zweite
 * Phase setzt die Stile zurück, um Formatierungsüberschreibungen zu bereinigen.
 *
 * Verwendung:
 * 1. Öffne das Zieldokument in InDesign.
 * 2. Stelle sicher, dass die Formatquelldatei ("format.indd") unter dem
 * im Skript definierten Pfad erreichbar ist.
 * 3. Für eine korrekte Ebenenpositionierung sollten im Dokument "Bilder"-
 * und/oder "Standard"-Ebenen vorhanden sein.
 * 4. Führe das Skript über das InDesign-Skripten-Bedienfeld aus.
 *
 * Version: 1.0
 *
 ******************************************************************************/

#target indesign

app.doScript(hauptSkript, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Neu Button auf allen Seiten erstellen und Stile zurücksetzen");

/**
 * Sucht einen Absatzstil im Dokument, optional innerhalb einer Stilgruppe.
 * @param {Document} dok Das zu durchsuchende InDesign-Dokument.
 * @param {string} gruppenName Name der Stilgruppe. Leer lassen für die Suche im Stammverzeichnis.
 * @param {string} stilName Name des zu suchenden Absatzstils.
 * @returns {ParagraphStyle|null} Das gefundene Stil-Objekt oder null.
 */
function findeAbsatzStil(dok, gruppenName, stilName) {
    if (gruppenName !== "") {
        var stilGruppe = dok.paragraphStyleGroups.itemByName(gruppenName);
        if (stilGruppe.isValid) {
            var stil = stilGruppe.paragraphStyles.itemByName(stilName);
            if (stil.isValid) return stil;
        }
    } else {
        var stil = dok.paragraphStyles.itemByName(stilName);
        if (stil.isValid) return stil;
    }
    return null;
}

/**
 * Sucht einen Objektstil im Dokument, optional innerhalb einer Stilgruppe.
 * @param {Document} dok Das zu durchsuchende InDesign-Dokument.
 * @param {string} gruppenName Name der Stilgruppe. Leer lassen für die Suche im Stammverzeichnis.
 * @param {string} stilName Name des zu suchenden Objektstils.
 * @returns {ObjectStyle|null} Das gefundene Stil-Objekt oder null.
 */
function findeObjektStil(dok, gruppenName, stilName) {
    if (gruppenName !== "") {
        var stilGruppe = dok.objectStyleGroups.itemByName(gruppenName);
        if (stilGruppe.isValid) {
            var stil = stilGruppe.objectStyles.itemByName(stilName);
            if (stil.isValid) return stil;
        }
    } else {
        var stil = dok.objectStyles.itemByName(stilName);
        if (stil.isValid) return stil;
    }
    return null;
}

/**
 * Konvertiert einen Millimeterwert in die aktuelle Maßeinheit des Skripts.
 * @param {number} mm Der Wert in Millimetern.
 * @returns {number} Der numerische Wert.
 */
function mmZuDokumentEinheit(mm) {
    var alteMasseinheit = app.scriptPreferences.measurementUnit;
    app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
    var wert = mm;
    app.scriptPreferences.measurementUnit = alteMasseinheit;
    return wert;
}

/**
 * Stellt die "Neu"-Ebene sicher und positioniert sie korrekt.
 * Verwendet eine "create then move" Methode für eine robuste Positionierung.
 * @param {Document} dok Das aktive InDesign-Dokument.
 * @returns {Layer} Das Ebenen-Objekt der "Neu"-Ebene.
 */
function ebenenVorbereiten(dok) {
    dok.selection = null;

    var ebeneName = "Neu";
    var ankerEbeneName = "Bilder";
    var fallbackEbeneName = "Standard";
    var neueEbene = dok.layers.itemByName(ebeneName);
    var ebeneExistiert = neueEbene.isValid;

    if (!ebeneExistiert) {
        neueEbene = dok.layers.add({name: ebeneName});
    }

    var ankerEbene = dok.layers.itemByName(ankerEbeneName);
    var fallbackEbene = dok.layers.itemByName(fallbackEbeneName);

    if (ankerEbene.isValid) {
        neueEbene.move(LocationOptions.AFTER, ankerEbene);
    } else if (fallbackEbene.isValid) {
        neueEbene.move(LocationOptions.BEFORE, fallbackEbene);
    }

    if (!ebeneExistiert) {
        neueEbene.layerColor = UIColors.RED;
    }
    
    return neueEbene;
}

/**
 * Die Hauptfunktion des Skripts. Führt alle Schritte zur Erstellung und
 * Formatierung der Objekte aus.
 */
function hauptSkript() {
    
    var formatquelldateiPfad = "x:\\_Grafik 2024\\Temp\\Neu Button\\Test Seite\\format.indd";
    
    // Objekt- und Gruppenstile
    var OBJEKT_GRUPPENNAME = ""; 
    var OBJEKT_STILNAME = "Neu Button Halbe Seite";
    var ABSATZ_GRUPPENNAME = "Formate NEU";

    // Deutsche Inhalte
    var ABSATZ_STILNAME_DE_LINKS = "Neu Button links";
    var ABSATZ_STILNAME_DE_RECHTS = "Neu Button rechts";
    var TEXTINHALT_DE = "Neu\tNeu\tNeu\tNeu\tNeu\tNeu\tNeu\tNeu\tNeu\tNeu";

    // Französische Inhalte
    var ABSATZ_STILNAME_FRZ_LINKS = "Nue Button links FRZ";
    var ABSATZ_STILNAME_FRZ_RECHTS = "Neu Button rechts FRZ";
    var TEXTINHALT_FRZ = "NOUVEAUTÉ\tNOUVEAUTÉ\tNOUVEAUTÉ\tNOUVEAUTÉ\tNOUVEAUTÉ\tNOUVEAUTÉ";
    
    // Allgemeine Einstellungen
    var rahmenBreite = 130;
    var rahmenHoehe = 9;
    var DREHWINKEL = 90;
    var BLEED_MM = 5;
    
    if (app.documents.length === 0) {
        alert("Fehler: Bitte öffne ein Dokument.");
        return;
    }
    var dok = app.activeDocument;

    var seitenHoehe = dok.documentPreferences.pageHeight;
    var seitenTyp = (seitenHoehe > 200) ? "GANZE_SEITE" : "HALBE_SEITE"; 

    var zielEbene = ebenenVorbereiten(dok);

    var formatquelldatei = new File(formatquelldateiPfad);
    if (!formatquelldatei.exists) {
        alert("Fehler: Die Formatvorlagendatei wurde nicht gefunden:\n" + formatquelldateiPfad);
        return;
    }

    try {
        dok.importStyles(ImportFormat.OBJECT_STYLES_FORMAT, formatquelldatei, GlobalClashResolutionStrategy.LOAD_ALL_WITH_OVERWRITE);
        dok.importStyles(ImportFormat.PARAGRAPH_STYLES_FORMAT, formatquelldatei, GlobalClashResolutionStrategy.LOAD_ALL_WITH_OVERWRITE);
    } catch (e) {
        alert("Fehler beim Importieren der Formate: " + e.message);
        return;
    }
    
    var objektStil = findeObjektStil(dok, OBJEKT_GRUPPENNAME, OBJEKT_STILNAME);
    if (objektStil === null) {
        alert("Fehler: Das Objektformat '" + OBJEKT_STILNAME + "' konnte nicht gefunden werden.");
        return;
    }

    var breiteInPunkten = mmZuDokumentEinheit(rahmenBreite);
    var hoeheInPunkten = mmZuDokumentEinheit(rahmenHoehe);
    var bleedInPunkten = mmZuDokumentEinheit(BLEED_MM);

    // ===================================================================
    // PHASE 1: OBJEKTE ERSTELLEN MIT TEXTRAHMENOPTIONEN
    // ===================================================================
    for (var i = 0; i < dok.pages.length; i++) {
        var seite = dok.pages.item(i);
        var seitenSeite = seite.side;
        var seitenBounds = seite.bounds;
        
        var aktuellerAbsatzStilName;
        var aktuellerTextInhalt;

        if (i < 2) {
            aktuellerTextInhalt = TEXTINHALT_DE;
            aktuellerAbsatzStilName = (seitenSeite === PageSideOptions.LEFT_HAND) ? ABSATZ_STILNAME_DE_LINKS : ABSATZ_STILNAME_DE_RECHTS;
        } else {
            aktuellerTextInhalt = TEXTINHALT_FRZ;
            aktuellerAbsatzStilName = (seitenSeite === PageSideOptions.LEFT_HAND) ? ABSATZ_STILNAME_FRZ_LINKS : ABSATZ_STILNAME_FRZ_RECHTS;
        }

        var absatzStil = findeAbsatzStil(dok, ABSATZ_GRUPPENNAME, aktuellerAbsatzStilName);
        if (absatzStil === null) { continue; }
        
        var textRahmen = seite.textFrames.add(zielEbene, { 
            geometricBounds: [0, 0, hoeheInPunkten, breiteInPunkten] 
        });
        
        var prefs = textRahmen.textFramePreferences;
        prefs.textColumnCount = 1;

        if (i === 0) { // Seite 1
            prefs.insetSpacing = [mmZuDokumentEinheit(0), mmZuDokumentEinheit(2), mmZuDokumentEinheit(0), mmZuDokumentEinheit(0)];
            prefs.verticalJustification = VerticalJustification.BOTTOM_ALIGN;
        } else if (i === 1) { // Seite 2
            prefs.insetSpacing = [mmZuDokumentEinheit(0.5), mmZuDokumentEinheit(2.3), mmZuDokumentEinheit(0), mmZuDokumentEinheit(0)];
            prefs.verticalJustification = VerticalJustification.TOP_ALIGN;
        } else if (i === 2) { // Seite 3
            prefs.insetSpacing = [mmZuDokumentEinheit(0), mmZuDokumentEinheit(2.3), mmZuDokumentEinheit(0.5), mmZuDokumentEinheit(0)];
            prefs.verticalJustification = VerticalJustification.BOTTOM_ALIGN;
        } else { // Seite 4 und alle folgenden
            prefs.insetSpacing = [mmZuDokumentEinheit(0.5), mmZuDokumentEinheit(2.3), mmZuDokumentEinheit(0), mmZuDokumentEinheit(0)];
            prefs.verticalJustification = VerticalJustification.TOP_ALIGN;
        }

        textRahmen.contents = aktuellerTextInhalt;
        textRahmen.parentStory.appliedParagraphStyle = absatzStil;
        textRahmen.rotationAngle = DREHWINKEL;
        textRahmen.applyObjectStyle(objektStil, true);
        textRahmen.parentStory.appliedCharacterStyle = dok.characterStyles.itemByName("[Ohne]");
        
        var y1, x1, y2, x2;
        
        if (seitenTyp === "GANZE_SEITE") {
            y1 = mmZuDokumentEinheit(117);
            y2 = y1 + breiteInPunkten;
        } else { // HALBE_SEITE
            y1 = -bleedInPunkten;
            y2 = y1 + breiteInPunkten;
        }
        
        if (seitenSeite === PageSideOptions.LEFT_HAND) {
            x1 = seitenBounds[1] - bleedInPunkten;
            x2 = x1 + hoeheInPunkten;
        } else {
            x2 = seitenBounds[3] + bleedInPunkten;
            x1 = x2 - hoeheInPunkten;
        }
        textRahmen.geometricBounds = [y1, x1, y2, x2];
    }

    // ===================================================================
    // PHASE 2: STILE ZURÜCKSETZEN UND ERNEUT ANWENDEN
    // ===================================================================
    var alleStile = [
        ABSATZ_STILNAME_DE_LINKS, ABSATZ_STILNAME_DE_RECHTS,
        ABSATZ_STILNAME_FRZ_LINKS, ABSATZ_STILNAME_FRZ_RECHTS
    ];

    var stilGruppe = dok.paragraphStyleGroups.itemByName(ABSATZ_GRUPPENNAME);
    if (stilGruppe.isValid) {
        for (var k = 0; k < alleStile.length; k++) {
            var stilZumLoeschen = stilGruppe.paragraphStyles.itemByName(alleStile[k]);
            if (stilZumLoeschen.isValid) { stilZumLoeschen.remove(); }
        }
    }

    try {
        dok.importStyles(ImportFormat.PARAGRAPH_STYLES_FORMAT, formatquelldatei, GlobalClashResolutionStrategy.LOAD_ALL_WITH_OVERWRITE);
    } catch (e) {
        alert("Fehler beim erneuten Importieren der Formate: " + e.message);
        return;
    }

    for (var j = 0; j < dok.pages.length; j++) {
        var seite = dok.pages.item(j);
        if (seite.textFrames.length === 0) continue;
        var textRahmen = seite.textFrames.item(0);
        var seitenSeite = seite.side;

        var stilName;
        if (j < 2) {
             stilName = (seitenSeite === PageSideOptions.LEFT_HAND) ? ABSATZ_STILNAME_DE_LINKS : ABSATZ_STILNAME_DE_RECHTS;
        } else {
             stilName = (seitenSeite === PageSideOptions.LEFT_HAND) ? ABSATZ_STILNAME_FRZ_LINKS : ABSATZ_STILNAME_FRZ_RECHTS;
        }
        
        var neuerAbsatzStil = findeAbsatzStil(dok, ABSATZ_GRUPPENNAME, stilName);
        if (neuerAbsatzStil !== null) {
            textRahmen.parentStory.appliedParagraphStyle = neuerAbsatzStil;
        }
    }
}