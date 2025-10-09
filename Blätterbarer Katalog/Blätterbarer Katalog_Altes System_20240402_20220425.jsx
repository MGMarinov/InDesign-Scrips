/********************************************************************************

 Blätterkatalog Verlinkung Indesign

 (c) COMINTO GmbH 2022

 COMINTO GmbH
 Klosterstraße 49
 D-40211 Düsseldorf
 +49 211 - 6000 16 81
 +49 211 - 6000 16 89

 support@blaetterkatalog.de

 ********************************************************************************/

/* CUSTOMER SPECIFIC */
var CUSTOMER = "personalshop";
var DEFAULT_DIR = "/personalshop";
var EXTENSION = "*.indd"; // choose between *.indd and *.idml
var KEEP_CSV_NAMES = false;
var EXPORT_GRAPHICS = true; // export graphics to the csv
var EXPORT_TABLES = false; // export tables to the csv
var EXPORT_TEXT_FRAMES = true; // export textframes to the csv
var EXPORT_PAGE_ITEMS = false; // export page items to the csv
var REGEX_TABLE_IN_STORY = /\d{6}/; //regex for the article number in tables
var REGEX_TEXTFRAME_IN_STORY = /(\d{2}[.])?\d{3}[.]\d{3}/; // regex for the article number in textframes

/* TEST MODE
 0 : processes all .indd files in the directory (default)
 1 : current document - only the first 10 graphics per document
 2 : current document - complete
 */
var TEST_MODE = 0;

var ID_VERSION = 5.0;
app.scriptPreferences.version = 4.0;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;

var pages = [];
switch (TEST_MODE) {
    case 2:
    case 1:
        handleTestModeCurrent();
        break;
    case 0:
    default:
        handleTestModeAll();
}
function handleTestModeCurrent() {
    if (TEST_MODE > 0 && app.documents.length === 1) {
        alert("Test Modus mit dem aktiven Indesign Dokument");
        var targetFolder = Folder.selectDialog("Indesign Blätterkatalog Verlinkung\n\nBitte Export Verzeichnis der Verlinkungsdateien angeben: ", DEFAULT_DIR);
        if (targetFolder !== null) {
            exportBK(app.activeDocument, targetFolder);
            alert("Testverarbeitung beendet. Export im Verzeichnis " + targetFolder);
        } else {
            alert("Die Verarbeitung wurde abgebrochen.");
        }
    }
}

function handleTestModeAll() {
    var sourceFolder = Folder.selectDialog("Indesign Blätterkatalog Verlinkung:\n\nBitte Verzeichnis der Indesign Dateien auswählen:\nHinweis: Es werden alle Dateien in diesem Verzeichnis ohne Prüfung der Dateiendung verarbeitet. ", DEFAULT_DIR);

    if (sourceFolder !== null) {
        var targetFolder = Folder.selectDialog("Indesign Blätterkatalog Verlinkung\n\nBitte Export Verzeichnis der Verlinkungsdateien angeben: ", DEFAULT_DIR);
        if (targetFolder !== null) {
            exportPages(sourceFolder, targetFolder);
            return;
        }
    }
    alert("Die Verarbeitung wurde abgebrochen.");
}

function exportPages(sourceFolder, targetFolder) {
    var files = sourceFolder.getFiles(EXTENSION);
    if (files.length > 0) {
        alert("Im Ordner " + sourceFolder + " werden " + files.length + " Dateien verarbeitet");
    } else {
        alert("Keine " + EXTENSION + " Dateien im gewählten Ordner " + sourceFolder);
        return;
    }
    for (var i = 0; i < files.length; i++) {
        var filename = "" + files[i];
        if (filenameIsValid(filename)) {
            try {
				pages = [];
                var doc = app.open(files[i], false);
                exportBK(doc, targetFolder, filename);
                doc.close(SaveOptions.no);
            } catch (exception) {
                alert("Fehler: Datei " + files[i] + " konnte nicht geöffnet werden:\n" + exception);
            }
        }
    }
    alert("Export Blätterkatalog fertig.\nEs wurden " + files.length + " Dateien verarbeitet");
}

function exportBK(doc, targetFolder, filename) {
    var temp = "";
    try {
        temp = createHeader(doc);
        temp = processPages(doc, temp);
        temp = processStories(doc, temp);
    } catch (exception) {
        alert("Fehler bei Export aus Indesign " + filename + "\nFehler: " + exception + "\nBitte Rücksprache mit support@cominto.de");
    }
    writeToCSV(doc, temp, targetFolder);
}

function createHeader(doc) {
    return doc.zeroPoint + ","
        + doc.documentPreferences.pageWidth + ","
        + doc.documentPreferences.pageHeight + ","
        + ID_VERSION + "\n";
}

function processPages(doc, temp) {
    var testCounter = 0;
    for (var currentPageId = 0; currentPageId < doc.pages.length; currentPageId++) {
        var currentPage = doc.pages[currentPageId];
        if (EXPORT_GRAPHICS) {
            temp = processGraphics(currentPage, temp, testCounter);
        }
        if (EXPORT_PAGE_ITEMS) {
            temp = processPageItems(currentPage, temp, testCounter);
        }
    }
    return temp;
}

function processStories(doc, temp) {
    for (var currentStoryId = 0; currentStoryId < doc.stories.length; currentStoryId++) {
        var currentStory = doc.stories.item(currentStoryId);
        if (EXPORT_TABLES) {
            temp = processTablesInStory(currentStory, temp);
        }
        if (EXPORT_TEXT_FRAMES) {
            temp = processTextFramesInStory(currentStory, temp);
        }
    }
    return temp;
}

function processGraphics(currentPage, temp, testCounter) {
    for (var graphicCounter = 0; graphicCounter < currentPage.allGraphics.length; graphicCounter++) {
        var graphic = currentPage.allGraphics[graphicCounter];
        if (graphic !== null && graphic.itemLink !== null && graphic.itemLink.name !== null && graphic.visibleBounds !== null) {
            var parentObject = graphic.parent;
            var graphicFileName = File.encode(graphic.itemLink.name);
            var page = getPageForItem(parentObject);
            graphicFileName = graphicFileName.replace(",", "_");
            if (!isNaN(currentPage.name) && currentPage.name > 0) {
                pages.push(currentPage.name);
                temp = temp + "G," + currentPage.name + "," + page + "," + graphicFileName + "," + graphic.visibleBounds;
                parentObject = graphic.parent;
                temp = concatTempWithShapes(parentObject, temp);
                temp = temp + "\n";
                if (TEST_MODE === 1) {
                    testCounter++;
                    if (testCounter > 10) break;
                }
            }
        }
    }
    return temp;
}

function processPageItems(currentPage, temp, testCounter) {
    for (var pageItemCounter = 0; pageItemCounter < currentPage.allPageItems.length; pageItemCounter++) {
        var currentPageItem = currentPage.allPageItems[pageItemCounter];
        if (currentPageItem.label !== "") {
            var visibleBounds = currentPageItem.visibleBounds;
            var linkType = "unknown";
            if (currentPageItem.allGraphics.length === 1) {
                var graphic = currentPageItem.allGraphics[0];
                if (graphic !== null && graphic.itemLink !== null && graphic.itemLink.name !== null && graphic.visibleBounds !== null) {
                    linkType = graphic.itemLink.name;
                    visibleBounds = graphic.visibleBounds;
                }
            }
            if (!isNaN(currentPage.name) && currentPage.name > 0) {
                pages.push(currentPage.name);
                temp = temp + "W," + currentPage.name + ",\"" + File.encode(linkType) + "\",\"" + File.encode(currentPageItem.label) + "\"," + visibleBounds;
                var parentObject = currentPageItem.parent;
                temp = concatTempWithPageItems(parentObject, temp);
                temp = temp + "\n";
                if (TEST_MODE === 1) {
                    testCounter++;
                    if (testCounter > 10) break;
                }
            }
        }
    }
    return temp;
}

function processTablesInStory(currentStory, temp) {
    for (var tableCounter = 0; tableCounter < currentStory.tables.length; tableCounter++) {
        var table = currentStory.tables.item(tableCounter);
        var parentObject = table.parent;
        var pageName = getPageName(parentObject);
        temp = processTableCells(table, temp, pageName);
    }
    return temp;
}

function processTextFramesInStory(currentStory, temp) {
    for (var textFrameCounter = 0; textFrameCounter < currentStory.textFrames.length; textFrameCounter++) {
        var textFrame = currentStory.textFrames.item(textFrameCounter);
        var parentObject = textFrame.parent;
        var pageName = getPageName(parentObject);
        var lines = currentStory.textFrames.item(textFrameCounter).lines;
        temp = processLines(lines, temp, pageName, REGEX_TEXTFRAME_IN_STORY, "T");
    }
    return temp;
}

function processTableCells(table, temp, pageName) {
    var cells = table.cells;
    for (var cellCounter = 0; cellCounter < cells.length; cellCounter++) {
        var cell = cells.item(cellCounter);
        var textStyleRanges = cell.textStyleRanges;
        if (textStyleRanges.length > 0) {
			var textStyleRange = textStyleRanges.item(0);
			var lines = textStyleRange.lines;
			temp = processLines(lines, temp, pageName, REGEX_TABLE_IN_STORY, "T");
        }
    }
    return temp;
}

function processLines(lines, temp, pageName, regex, letter) {
    for (var lineCounter = 0; lineCounter < lines.length; lineCounter++) {
        try {
            var line = lines.item(lineCounter);
            var lineContents = "" + line.contents;
            var characters = line.characters;
            var horizontalOffset = characters[0].horizontalOffset;
            var result = regex.exec(lineContents);
            if (result === undefined) {
                continue;
            }
            var artNr = line.contents;
            if (result !== undefined) {
                artNr = result[0];
                if (!isNaN(pageName) && pageName > 0) {
                    pages.push(pageName);
                    temp = temp + letter + "," + pageName + "," + artNr + "," + (characters[0].baseline - line.ascent - line.descent) + "," + horizontalOffset + "," + (characters[0].baseline) + "," + (characters[characters.length - 1].horizontalOffset) + "\n";
                }
            }
        } catch (exception) {
            // continue with loop
        }
    }
    return temp;
}

function getPageForItem(parentObject) {
    var parentType = "" + parentObject;
    var page = 0;
    var counter = 0;
    while (parentType.indexOf("Page") < 0 && counter < 10) {
        if (parentType.indexOf("Rectangle") > 0) {
            page = parentObject.index;
        }
        parentObject = parentObject.parent;
        parentType = "" + parentObject;
        counter++;
    }
    return page;
}

function getPageName(parentObject) {
    var parentType = "" + parentObject;
    var counter = 0;
    while (parentType.indexOf("Page") < 0 && counter < 10) {
        parentObject = parentObject.parent;
        parentType = "" + parentObject;
        counter++;
    }
    var pageName = "";
    if (parentType.indexOf("Page") > 0) {
        pageName = parentObject.name;
    }
    return pageName;
}

function concatTempWithShapes(parentObject, temp) {
    var parentType = "" + parentObject;
    while (parentType.indexOf("Polygon") > 0
    || parentType.indexOf("Rectangle") > 0
    || parentType.indexOf("Oval") > 0) {
        temp = temp + "," + parentObject.visibleBounds;
        parentObject = parentObject.parent;
        parentType = "" + parentObject;
    }
    return temp;
}

function concatTempWithPageItems(parentObject, temp) {
    var parentType = "" + parentObject;
    while (parentType.indexOf("Polygon") > 0
    || parentType.indexOf("Rectangle") > 0
    || parentType.indexOf("Oval") > 0
    || parentType.indexOf("group") > 0) {
        temp = temp + "," + parentObject.visibleBounds;
        parentObject = parentObject.parent;
        parentType = "" + parentObject;
    }
    return temp;
}

function writeToCSV(doc, temp, targetFolder) {
    pages.sort(function (a, b) {
        return a - b;
    });
  	var filename;
	if (KEEP_CSV_NAMES) {
		filename = doc.name + ".csv";
	} else {
		filename = CUSTOMER + "_" + pages[0] + "-" + pages[pages.length - 1] + ".csv";
	}
    var folder = targetFolder + "/" + filename;
    var file = new File(folder);
    file.open("w");

    var success = file.write(temp);
    if (!success) alert("ERROR writing file");
    file.close();
}

function filenameIsValid(filename) {
    return filename != undefined && filename.indexOf("ridge") === -1 && filename.indexOf("cache") === -1 && filename.indexOf("DS_Store") === -1;
}