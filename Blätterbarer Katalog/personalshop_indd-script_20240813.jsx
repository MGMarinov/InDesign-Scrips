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
var EXPORT_TEXT_FRAMES_IN_STORIES = true; // export textframes from stories to the csv
var EXPORT_PAGE_ITEMS = false; // export page items to the csv
var REGEX_TABLE_IN_STORY = /\d{6}/; //regex for the article number in tables
var REGEX_TEXTFRAME_IN_STORY = /(\d{2}[.])?\d{3}[.]\d{3}/; // regex for the article number in textframes
var SELECTED_LAYER = "";

/* TEST MODE
 0 : processes all .indd files in the directory (default)
 1 : current document - only the first 10 graphics per document
 2 : current document - complete
 */
var TEST_MODE = 0;

var ID_VERSION = 5.0;
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
        var element = files[i];
        var filename = "" + element;
        if (filenameIsValid(filename)) {
            try {
                pages = [];
                var doc = app.open(element, false);
                exportBK(doc, targetFolder, filename);
                doc.close(SaveOptions.no);
            } catch (exception) {
                alert("Fehler: Datei " + element + " konnte nicht geöffnet werden:\n" + exception);
            }
        }
    }
    alert("Export Blätterkatalog fertig.\nEs wurden " + files.length + " Dateien verarbeitet");
}

function determineLayer(doc) {
    var layerNames = getAllLayerNames(doc);
    var selectedLayerIndex = showLayerSelectionDialog(layerNames);
    if (selectedLayerIndex !== null && selectedLayerIndex >= 0) {
        SELECTED_LAYER = layerNames[selectedLayerIndex];
    }
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        var element = doc.layers[i];
        if (element.name == SELECTED_LAYER) {
            layer = element;
            break;
        }
    }
    return layer;
}

function processLink(link, temp) {
    var itemLink = link.itemLink;
    var page = link.page;
    if (itemLink.filePath && itemLink.filePath.match(/\.indd$/i)) {
        try {
            var pageNumber = itemLink.parent.pageNumber;
            if (!isNaN(pageNumber)) {
                var linkedDocFile = new File(itemLink.filePath);
                if (linkedDocFile.exists) {
                    var linkedDoc = app.open(linkedDocFile);
                    temp = processPages(linkedDoc, temp, page, pageNumber);
                    temp = processStories(linkedDoc, temp, page, pageNumber);
                    temp = processTextframes(linkedDoc, temp, page, pageNumber);
                    linkedDoc.close(SaveOptions.NO);
                } else {
                    alert("Datei existiert nicht: " + linkedDocFile.fullName);
                }
            }
        } catch (e) {
            alert("Fehler beim Öffnen der Datei: " + e.toString());
        }
    }
    return temp;
}

function exportBK(doc, targetFolder, filename) {
    var temp = "";
    try {
        temp = createHeader(doc);
        var layer = determineLayer(doc);
        var linksInLayer = determineLinksInLayer(doc, layer, temp);

        if (layer != null) {
            for (var i = 0; i < linksInLayer.length; i++) {
                temp = processLink(linksInLayer[i], temp)
            }
        }
    } catch (exception) {
        alert("Fehler bei Export aus Indesign " + filename + "\nFehler: " + exception + "\nBitte Rücksprache mit support@blaetterkatalog.de");
    }
    writeToCSV(doc, temp, targetFolder);
}

function determineLinksInLayer(doc, layer) {
    var linksInLayer = [];
    for (var i = 0; i < doc.pages.length; i++) {
        var currentPage = doc.pages[i];
        for (var pageItemCounter = 0; pageItemCounter < currentPage.allPageItems.length; pageItemCounter++) {
            var currentPageItem = currentPage.allPageItems[pageItemCounter];
            if (currentPageItem instanceof Rectangle)
            {
                var rect = currentPageItem;
                if (rect.hasOwnProperty("importedPages")) {
                    for (var j = 0; j < rect.importedPages.length; j++)
                    {
                        if (rect.importedPages[j].itemLayer == layer) {
                            var importedPage = rect.importedPages[j];
                            linksInLayer.push({
                                itemLink: importedPage.itemLink,
                                page: currentPage
                            });
                        }
                    }
                }
            }
        }
    }
    
    return linksInLayer;
}

function getAllLayerNames(doc) {
    var layers = doc.layers;
    var layerNames = [];
    for (var i = 0; i < layers.length; i++) {
        var element = layers[i];
        layerNames.push(element.name);
    }
    return layerNames;
}

function showLayerSelectionDialog(layerNames) {
    var dialog = new Window("dialog", "Ebene auswählen");
    dialog.orientation = "column";
    dialog.add("statictext", undefined, "Wähle eine Ebene:");
    var dropdown = dialog.add("dropdownlist", undefined, layerNames);
    dropdown.selection = 0;
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    var selectedLayerIndex = null;

    okButton.onClick = function () {
        selectedLayerIndex = dropdown.selection.index;
        dialog.close();
    };

    cancelButton.onClick = function () {
        dialog.close();
    };

    dialog.show();
    return selectedLayerIndex;
}

function createHeader(doc) {
    return doc.zeroPoint + ","
        + doc.documentPreferences.pageWidth + ","
        + doc.documentPreferences.pageHeight + ","
        + ID_VERSION + "\n";
}

function processPages(doc, temp, page, pageNumber) {
    var testCounter = 0;
    var currentPage = doc.pages[pageNumber -1];
    for (var j = 0; j < currentPage.allPageItems.length; j++) {
        var currentPageItem = currentPage.allPageItems[j];
        if (EXPORT_GRAPHICS) {
            temp = processGraphics(currentPageItem, temp, testCounter, page);
        }
        if (EXPORT_PAGE_ITEMS) {
            temp = processPageItems(currentPageItem, temp, testCounter, page);
        }
    }
    return temp;
}

function processStories(doc, temp, page, pageNumber) {
    for (var currentStoryId = 0; currentStoryId < doc.stories.length; currentStoryId++) {
        var currentStory = doc.stories.item(currentStoryId);
		var currentPage = doc.pages[pageNumber -1];
        for (var textFrameCounter = 0; textFrameCounter < currentStory.textFrames.length; textFrameCounter++) {
            var currentTextFrame = currentStory.textFrames.item(textFrameCounter);
			if (currentTextFrame.parentPage.name == currentPage.name) {
                if (EXPORT_TABLES) {
                    temp = processTablesInStory(currentTextFrame, temp, page);
                }
                if (EXPORT_TEXT_FRAMES_IN_STORIES) {
                    temp = processTextFramesInStory(currentTextFrame, temp, page);
                }
			}
        }
    }
    return temp;
}

function processTextframes(doc, temp, page, pageNumber) {
    if (EXPORT_TEXT_FRAMES) {
        var currentPage = doc.pages[pageNumber -1];
        for (var i = 0; i < currentPage.allPageItems.length; i++) {
            var currentPageItem = currentPage.allPageItems[i];
            if (currentPageItem instanceof TextFrame) {
                var lines = currentPageItem.lines;
                temp = processLines(lines, temp, page.name, REGEX_TEXTFRAME_IN_STORY, "T");
            }
        }
    }

    return temp;
}

function processPageItemsTextFrames(currentPage, temp, pageNumber) {
    for (var pageItemCounter = 0; pageItemCounter < currentPage.allPageItems.length; pageItemCounter++) {
        var currentPageItem = currentPage.allPageItems[pageItemCounter];
        if (currentPageItem instanceof TextFrame)
        {
			var lines = currentPageItem.lines;
			temp = processLines(lines, temp, pageNumber, REGEX_TEXTFRAME_IN_STORY, "T");
        }
    }
    
    return temp;
} 

function processGraphics(currentPage, temp, testCounter, page) {
    for (var i = 0; i < currentPage.allGraphics.length; i++) {
        var graphic = currentPage.allGraphics[i];
        if (graphic !== null && graphic.itemLink !== null && graphic.itemLink.name !== null && graphic.visibleBounds !== null) {
            var parentObject = graphic.parent;
            var graphicFileName = File.encode(graphic.itemLink.name);
            graphicFileName = graphicFileName.replace(",", "_");
            if (!isNaN(page.name) && page.name > 0) {
                pages.push(page.name);
                temp = temp + "G," + page.name + "," + page.index + "," + graphicFileName + "," + graphic.visibleBounds;
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

function processPageItems(currentPage, temp, testCounter, page) {
    for (var i = 0; i < currentPage.allPageItems.length; i++) {
        var currentPageItem = currentPage.allPageItems[i];
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
            if (!isNaN(page.name) && page.name > 0) {
                pages.push(page.name);
                temp = temp + "W," + page.name + ",\"" + File.encode(linkType) + "\",\"" + File.encode(currentPageItem.label) + "\"," + visibleBounds;
                var parentObject = currentPageItem.parent;
                temp = concatTempWithPageItems(parentObject, temp);
                temp = temp + "\n";
                if (TEST_MODE === 1) {
                    testCounter++;
                    if (testCounter > 10) break;
                }
            }
        }
        if (currentPageItem instanceof TextFrame)  {
                var parentObject = currentPageItem.parent;
                var pageName = getPageName(parentObject);
                var lines = currentPageItem.lines;
                temp = processLines(lines, temp, pageName, REGEX_TEXTFRAME_IN_STORY, "T");
        }
    }
    return temp;
}

function processTablesInStory(currentStory, temp, page) {
    for (var tableCounter = 0; tableCounter < currentStory.tables.length; tableCounter++) {
        var table = currentStory.tables.item(tableCounter);
        temp = processTableCells(table, temp, page.name);
    }
    return temp;
}

function processTextFramesInStory(currentStory, temp, page) {
    for (var textFrameCounter = 0; textFrameCounter < currentStory.textFrames.length; textFrameCounter++) {
        var lines = currentStory.textFrames.item(textFrameCounter).lines;
        temp = processLines(lines, temp, page.name, REGEX_TEXTFRAME_IN_STORY, "T");
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
            if (result === undefined || result === null) {
                continue;
            }
            var artNr = result[0];
            pages.push(pageName);
            temp = temp + letter + "," + pageName + "," + artNr + "," + (characters[0].baseline - line.ascent - line.descent) + "," + horizontalOffset + "," + (characters[0].baseline) + "," + (characters[characters.length - 1].horizontalOffset) + "\n";
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

function getUniqueLines(temp) {
    var lines = temp.split('\n');
    var uniqueLinesMap = {};
    var uniqueLines = [];

    for (var i = 0; i < lines.length; i++) {
        if (!uniqueLinesMap.hasOwnProperty(lines[i])) {
            uniqueLinesMap[lines[i]] = true;
            uniqueLines.push(lines[i]);
        }
    }

    return uniqueLines.join('\n');
}

function writeToCSV(doc, temp, targetFolder) {
    pages.sort(function (a, b) {
        return a - b;
    });

    temp = getUniqueLines(temp);

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
