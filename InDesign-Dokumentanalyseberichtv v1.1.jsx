/*******************************************************************************
 *
 * InDesign-Dokumentanalyseberichtv
 *
 * Beschreibung:
 * Dieses Skript erstellt einen detaillierten Bericht über den Inhalt eines
 * InDesign-Dokuments im JSON-Format. Es analysiert Dokumenteigenschaften
 * (Seitengröße, Anschnitt, Farbfelder) sowie Objekte, deren Ebenen,
 * Textinhalte und Stile auf jeder Seite. Nach der Erstellung wird der
 * Speicherort des Berichts automatisch geöffnet.
 *
 * Version: 1.1
 *
 ******************************************************************************/

#target indesign

// ===================================================================
// JSON Polyfill (json2.js) von Douglas Crockford
// Stellt JSON-Funktionalität für ältere JavaScript-Engines bereit.
// ===================================================================
if(typeof JSON!=='object'){JSON={};}
(function(){'use strict';function f(n){return n<10?'0'+n:n;}
if(typeof Date.prototype.toJSON!=='function'){Date.prototype.toJSON=function(key){return isFinite(this.valueOf())?this.getUTCFullYear()+'-'+
f(this.getUTCMonth()+1)+'-'+
f(this.getUTCDate())+'T'+
f(this.getUTCHours())+':'+
f(this.getUTCMinutes())+':'+
f(this.getUTCSeconds())+'Z':null;};String.prototype.toJSON=Number.prototype.toJSON=Boolean.prototype.toJSON=function(key){return this.valueOf();};}
var cx=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,escapable=/[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,gap,indent,meta={'\b':'\\b','\t':'\\t','\n':'\\n','\f':'\\f','\r':'\\r','"':'\\"','\\':'\\\\'},rep;function quote(string){escapable.lastIndex=0;return escapable.test(string)?'"'+string.replace(escapable,function(a){var c=meta[a];return typeof c==='string'?c:'\\u'+('0000'+a.charCodeAt(0).toString(16)).slice(-4);})+'"':'"'+string+'"';}
function str(key,holder){var i,k,v,length,mind=gap,partial,value=holder[key];if(value&&typeof value==='object'&&typeof value.toJSON==='function'){value=value.toJSON(key);}
if(typeof rep==='function'){value=rep.call(holder,key,value);}
switch(typeof value){case'string':return quote(value);case'number':return isFinite(value)?String(value):'null';case'boolean':case'null':return String(value);case'object':if(!value){return'null';}
gap+=indent;partial=[];if(Object.prototype.toString.apply(value)==='[object Array]'){length=value.length;for(i=0;i<length;i+=1){partial[i]=str(i,value)||'null';}
v=partial.length===0?'[]':gap?'[\n'+gap+partial.join(',\n'+gap)+'\n'+mind+']':'['+partial.join(',')+']';gap=mind;return v;}
if(rep&&typeof rep==='object'){length=rep.length;for(i=0;i<length;i+=1){if(typeof rep[i]==='string'){k=rep[i];v=str(k,value);if(v){partial.push(quote(k)+(gap?': ':':')+v);}}}}else{for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=str(k,value);if(v){partial.push(quote(k)+(gap?': ':':')+v);}}}}
v=partial.length===0?'{}':gap?'{\n'+gap+partial.join(',\n'+gap)+'\n'+mind+'}':'{'+partial.join(',')+'}';gap=mind;return v;}}
if(typeof JSON.stringify!=='function'){JSON.stringify=function(value,replacer,space){var i;gap='';indent='';if(typeof space==='number'){for(i=0;i<space;i+=1){indent+=' ';}}else if(typeof space==='string'){indent=space;}
rep=replacer;if(replacer&&typeof replacer!=='function'&&(typeof replacer!=='object'||typeof replacer.length!=='number')){throw new Error('JSON.stringify');}
return str('',{'':value});};}
if(typeof JSON.parse!=='function'){JSON.parse=function(text,reviver){var j;function walk(holder,key){var k,v,value=holder[key];if(value&&typeof value==='object'){for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=walk(value,k);if(v!==undefined){value[k]=v;}else{delete value[k];}}}}
return reviver.call(holder,key,value);}
text=String(text);cx.lastIndex=0;if(cx.test(text)){text=text.replace(cx,function(a){return'\\u'+('0000'+a.charCodeAt(0).toString(16)).slice(-4);});}
if(/^[\],:{}\s]*$/.test(text.replace(/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g,'@').replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,']').replace(/(?:^|:|,)(?:\s*\[)+/g,''))){j=eval('('+text+')');return typeof reviver==='function'?walk({'':j},''):j;}
throw new SyntaxError('JSON.parse');};}}());
// ===================================================================
// Ende des JSON Polyfills
// ===================================================================

try {
    hauptfunktion();
} catch (e) {
    alert("Ein Fehler ist während der Skriptausführung aufgetreten:\n" + e.toString() + "\nZeile: " + e.line);
}

/**
 * Erstellt einen Zeitstempel-String für den Dateinamen.
 * @returns {string} Ein formatierter Zeitstempel (JJJJMMTT_HHMM).
 */
function erstelleZeitstempel() {
    var jetzt = new Date();
    var jahr = jetzt.getFullYear();
    var monat = ("0" + (jetzt.getMonth() + 1)).slice(-2);
    var tag = ("0" + jetzt.getDate()).slice(-2);
    var stunde = ("0" + jetzt.getHours()).slice(-2);
    var minute = ("0" + jetzt.getMinutes()).slice(-2);
    return jahr + monat + tag + "_" + stunde + minute;
}

/**
 * Die Hauptfunktion des Skripts, die den Analyseprozess steuert und
 * den JSON-Bericht erstellt.
 */
function hauptfunktion() {
    if (app.documents.length == 0) {
        alert("Kein Dokument geöffnet!\nBitte öffne eine Datei, bevor du das Skript ausführst.");
        return;
    }
    var aktivesDokument = app.activeDocument;

    if (!aktivesDokument.saved) {
        alert("Das Dokument ist nicht gespeichert!\nBitte speichere die Datei, bevor du den Bericht erstellst.");
        return;
    }

    var dokumentPfad = aktivesDokument.filePath;
    var berichtDateiName = "InDesign-Dokumentanalyseberichtv_" + erstelleZeitstempel() + ".json";
    var berichtDateiPfad = dokumentPfad + "/" + berichtDateiName;

    var reportObjekt = {};
    reportObjekt.berichtVersion = "1.1";
    reportObjekt.dokumentenname = aktivesDokument.name;
    reportObjekt.erstellungsdatum = new Date().toLocaleString();
    
    var docPrefs = aktivesDokument.documentPreferences;
    reportObjekt.dokumenteigenschaften = {
        seitengroesse: {
            breite: docPrefs.pageWidth.toFixed(2) + " pt",
            hoehe: docPrefs.pageHeight.toFixed(2) + " pt"
        },
        anschnitt: {
            einheitlich: docPrefs.documentBleedUniformSize,
            oben: docPrefs.documentBleedTopOffset.toFixed(2) + " pt",
            unten: docPrefs.documentBleedBottomOffset.toFixed(2) + " pt",
            innen: docPrefs.documentBleedInsideOrLeftOffset.toFixed(2) + " pt",
            aussen: docPrefs.documentBleedOutsideOrRightOffset.toFixed(2) + " pt"
        }
    };
    
    reportObjekt.farbfelder = [];
    var alleSwatches = aktivesDokument.swatches;
    for (var s = 0; s < alleSwatches.length; s++) {
        var swatch = alleSwatches[s];
        var swatchInfo = { name: swatch.name };
        try {
            swatchInfo.farbraum = swatch.space.toString();
            swatchInfo.farbwerte = swatch.colorValue;
        } catch(e) {
            swatchInfo.typ = "Spezialfarbe / Ohne Farbe";
        }
        reportObjekt.farbfelder.push(swatchInfo);
    }
    
    reportObjekt.seiten = [];
    for (var i = 0; i < aktivesDokument.pages.length; i++) {
        var seite = aktivesDokument.pages[i];
        var seitenDaten = { seitenname: seite.name, musterseiten_objekte: [], seitenspezifische_objekte: [] };

        var musterSeite = seite.appliedMaster;
        if (musterSeite !== null) {
            var musterObjekte = musterSeite.pageItems.everyItem().getElements();
            for (var j = 0; j < musterObjekte.length; j++) {
                seitenDaten.musterseiten_objekte.push(verarbeiteObjekt(musterObjekte[j]));
            }
        }

        var seitenObjekte = seite.pageItems.everyItem().getElements();
        for (var k = 0; k < seitenObjekte.length; k++) {
            seitenDaten.seitenspezifische_objekte.push(verarbeiteObjekt(seitenObjekte[k]));
        }
        reportObjekt.seiten.push(seitenDaten);
    }

    var berichtDatei = new File(berichtDateiPfad);
    berichtDatei.encoding = "UTF-8";
    berichtDatei.open("w");
    berichtDatei.write(JSON.stringify(reportObjekt, null, "  "));
    berichtDatei.close();
    
    berichtDatei.parent.execute();

    alert("Analyse abgeschlossen!\n\nDer JSON-Bericht wurde gespeichert und der Ordner geöffnet.");
}

/**
 * Verarbeitet ein einzelnes Seitenobjekt und gibt dessen Informationen als Objekt zurück.
 * @param {PageItem} objekt Das zu analysierende InDesign-Objekt.
 * @returns {Object} Ein Objekt mit den analysierten Eigenschaften.
 */
function verarbeiteObjekt(objekt) {
    var objektInfo = {};
    try {
        objektInfo.objekttyp = objekt.constructor.name;
        objektInfo.ebene = objekt.itemLayer.name; 
        
        var grenzen = objekt.visibleBounds;
        objektInfo.position_mm = { y: grenzen[0].toFixed(2), x: grenzen[1].toFixed(2) };

        if (objekt.overridden) {
            objektInfo.status = "Überschriebenes Musterobjekt";
        }
        if (objekt.rotationAngle !== 0) {
            objektInfo.rotation = objekt.rotationAngle.toFixed(2) + "°";
        }
        if (objekt.appliedObjectStyle) {
            objektInfo.objektformat = objekt.appliedObjectStyle.name;
        }

        switch (objektInfo.objekttyp) {
            case "TextFrame":
                objektInfo.details = analysiereTextrahmen(objekt);
                break;
            case "Rectangle": case "Oval": case "Polygon":
                objektInfo.details = analysiereGrafikrahmen(objekt);
                break;
            case "Group":
                objektInfo.gruppeninhalt = [];
                var gruppenObjekte = objekt.pageItems.everyItem().getElements();
                for (var k = 0; k < gruppenObjekte.length; k++) {
                    objektInfo.gruppeninhalt.push(verarbeiteObjekt(gruppenObjekte[k]));
                }
                break;
            case "Image": case "PDF": case "EPS":
                objektInfo.details = analysierePlatziertesObjekt(objekt);
                break;
            default:
                objektInfo.details = analysiereGrafikrahmen(objekt);
                break;
        }
    } catch (e) {
        objektInfo.fehler = "Fehler bei der Verarbeitung: " + e.message;
    }
    return objektInfo;
}

/**
 * Analysiert die Eigenschaften eines Textrahmens.
 * @param {TextFrame} rahmen Der zu analysierende Textrahmen.
 * @returns {Object} Ein Objekt mit den analysierten Texteigenschaften.
 */
function analysiereTextrahmen(rahmen) {
    var textInfo = {};
    try {
        if (rahmen.parentStory && rahmen.parentStory.contents.length > 0) {
            var story = rahmen.parentStory;
            textInfo.inhalt_vorschau = story.contents.substring(0, 75).replace(/[\r\n]+/g, ' ');
            textInfo.absatzformate = findeEindeutigeStile(story.paragraphs, "appliedParagraphStyle");
            textInfo.zeichenformate = findeEindeutigeStile(story.characters, "appliedCharacterStyle");
        } else {
            textInfo.inhalt = "leer";
        }
        
        var prefs = rahmen.textFramePreferences;
        var insets = prefs.insetSpacing;
        textInfo.textrahmenoptionen = {
            spaltenanzahl: prefs.textColumnCount,
            abstand_pt: { oben: insets[0].toFixed(2), unten: insets[2].toFixed(2), links: insets[1].toFixed(2), rechts: insets[3].toFixed(2) },
            vertikale_ausrichtung: konvertiereVJust(prefs.verticalJustification)
        };

    } catch (e) {
        textInfo.fehler = "Fehler bei der Textrahmenanalyse: " + e.message;
    }
    
    var grafikInfo = analysiereGrafikrahmen(rahmen);
    for (var key in grafikInfo) {
        textInfo[key] = grafikInfo[key];
    }
    return textInfo;
}

/**
 * Analysiert die Füllung, Kontur und den grafischen Inhalt eines Rahmens.
 * @param {PageItem} objekt Der zu analysierende Grafikrahmen oder ein anderes Objekt.
 * @returns {Object} Ein Objekt mit den grafischen Eigenschaften.
 */
function analysiereGrafikrahmen(objekt) {
    var grafikInfo = {};
    try {
        grafikInfo.fuellung = objekt.fillColor.name;
        grafikInfo.kontur = { farbe: objekt.strokeColor.name, staerke: objekt.strokeWeight + " pt" };
    } catch (e) {}
    if (objekt.graphics.length > 0) {
        grafikInfo.grafikinhalt = analysierePlatziertesObjekt(objekt.graphics[0]);
    }
    return grafikInfo;
}

/**
 * Analysiert ein platziertes Objekt wie ein Bild oder PDF.
 * @param {Graphic} objekt Das platzierte Grafikobjekt.
 * @returns {Object} Ein Objekt mit den Verknüpfungsinformationen.
 */
function analysierePlatziertesObjekt(objekt) {
    var platziertInfo = {};
    if (objekt.itemLink && objekt.itemLink.isValid) {
        platziertInfo.verknuepfte_datei = decodeURI(objekt.itemLink.name);
        platziertInfo.status = objekt.itemLink.status.toString();
    } else {
        platziertInfo.status = "Eingebettetes oder fehlendes Objekt.";
    }
    return platziertInfo;
}

/**
 * Konvertiert den Enum-Wert der vertikalen Ausrichtung in einen lesbaren String.
 * @param {number} enumWert Der Enum-Wert von VerticalJustification.
 * @returns {string} Der lesbare Name der Ausrichtung.
 */
function konvertiereVJust(enumWert) {
    switch (enumWert) {
        case VerticalJustification.TOP_ALIGN: return "Oben";
        case VerticalJustification.CENTER_ALIGN: return "Mitte";
        case VerticalJustification.BOTTOM_ALIGN: return "Unten";
        case VerticalJustification.JUSTIFY_ALIGN: return "Blocksatz";
        default: return "Unbekannt";
    }
}

/**
 * Findet eindeutige Stilnamen in einer Sammlung von InDesign-Objekten.
 * @param {Array} sammlung Die Sammlung von Objekten (z.B. paragraphs).
 * @param {string} stilEigenschaft Der Name der Stileigenschaft (z.B. "appliedParagraphStyle").
 * @returns {Array<string>} Ein Array mit den eindeutigen Stilnamen.
 */
function findeEindeutigeStile(sammlung, stilEigenschaft) {
    var stilNamenObjekt = {};
    var stilNamenArray = [];
    for (var i = 0; i < sammlung.length; i++) {
        try {
            var stilName = sammlung[i][stilEigenschaft].name;
            stilNamenObjekt[stilName] = true;
        } catch (e) {}
    }
    for (var key in stilNamenObjekt) {
        if (stilNamenObjekt.hasOwnProperty(key)) {
            stilNamenArray.push(key);
        }
    }
    return stilNamenArray;
}