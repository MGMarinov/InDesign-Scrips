# InDesign UXP Script Fehleranalyse und Korrekturvorschläge

Dieses Dokument analysiert das bereitgestellte InDesign UXP-Script, identifiziert potenzielle Fehler und Betriebsprobleme und gibt Lösungsvorschläge. Die UXP (Unified Extensibility Platform) hat spezifische Anforderungen, die sich von der älteren ExtendScript-Umgebung unterscheiden können, was bei der Entwicklung zu berücksichtigen ist.

---

## 1. Fehler: Modulimport

### Beschreibung
Das Skript versucht, Module mit der herkömmlichen `require()`-Syntax zu laden, was in der UXP-Umgebung in bestimmten Fällen problematisch sein kann.

Problematischer Code:
const { app } = require("indesign");
const fs = require("uxp").storage.localFileSystem;

Das UXP-Modulsystem unterscheidet sich von Node.js. Während `require("indesign")` korrekt ist, erfordert der Zugriff auf die Submodule des `uxp`-Moduls (`storage.localFileSystem`) eine andere Syntax.

### Vorschlag
Verwenden Sie die destrukturierende Zuweisung aus dem `storage`-Objekt des `uxp`-Moduls.

Korrigierter Code:
const { app } = require("indesign");
const { fs } = require("uxp").storage;
// Oder falls localFileSystem direkt benötigt wird:
// const { storage } = require("uxp");
// const fs = storage.localFileSystem;

Dieser Ansatz stellt sicher, dass das korrekte Dateisystem-Handler-Objekt abgerufen wird.

---

## 2. Fehler: Deklaration von UI-Komponenten (Spectrum)

### Beschreibung
Das Skript verwendet Spectrum-UI-Komponenten (`<sp-dialog>`, `<sp-picker>` usw.), deklariert diese aber nicht als Abhängigkeit in der `manifest.json`-Datei. Ohne diese Deklaration lädt die UXP-Umgebung diese Elemente nicht und zeigt sie nicht an.

### Vorschlag
Fügen Sie das `spectrum`-Paket zum `dependencies`-Abschnitt der `manifest.json`-Datei hinzu.

Beispiel für die `manifest.json`-Datei:
{
  "manifestVersion": 5,
  "id": "com.example.katalogseiten-austausch",
  "name": "Katalog Seiten austausch",
  "version": "0.1.0",
  "host": {
    "app": "ID",
    "minVersion": "18.0"
  },
  "dependencies": [
    {
      "id": "spectrum",
      "version": "1.0.0"
    }
  ],
  "entrypoint": "./index.html",
  "requiredPermissions": [
    "localFileSystem",
    "fileSystem:read",
    "fileSystem:write"
  ]
}

Hinweis: Der Abschnitt `requiredPermissions` wird im nächsten Punkt näher erläutert.

---

## 3. Fehler: Fehlende Dateisystemberechtigungen

### Beschreibung
Das Skript versucht, Dateien zu lesen (`leseProduktliste`) und zu schreiben (`schreibeReport`). Das UXP-Sicherheitsmodell (Sandbox) erfordert, dass das Plugin explizit Berechtigungen für Dateisystemvorgänge in der `manifest.json`-Datei anfordert. Ohne diese Berechtigungen schlagen die Aufrufe mit einem `Permission denied`-Fehler fehl.

### Vorschlag
Fügen Sie die erforderlichen Berechtigungen zur `requiredPermissions`-Liste in der `manifest.json`-Datei hinzu.

Beispiel für den `requiredPermissions`-Abschnitt:
"requiredPermissions": [
  "localFileSystem",
  "fileSystem:read",
  "fileSystem:write"
]

- `localFileSystem`: Bietet grundlegenden Zugriff auf das Dateisystem.
- `fileSystem:read`: Erlaubt das Lesen von Dateien.
- `fileSystem:write`: Erlaubt das Schreiben und Erstellen von Dateien.

---

## 4. Fehler: Verwendung der InDesign-API und Änderungen

### Beschreibung
Obwohl der Großteil des Skripts die InDesign-API korrekt verwendet, können an einigen Stellen Inkompatibilitäten oder Unterschiede zwischen UXP und dem älteren ExtendScript auftreten.

Potenziell problematische Bereiche:
1. `app.pdfPlacePreferences`: Obwohl dieses Objekt existiert, muss seine korrekte Funktion unter UXP überprüft werden.
2. `seite.place(...)`: Die `place()`-Methode gibt in UXP ein Array der platzierten Objekte zurück. Das Skript behandelt dies korrekt (`[0]`), aber es ist wichtig zu wissen, dass wenn das Platzieren fehlschlägt, das Array leer sein kann, was beim Zugriff auf `[0]` zu einem `undefined`-Fehler führt.

### Vorschlag
1. **Dokumentation prüfen**: Überprüfen Sie immer die neueste [Adobe UXP for InDesign API-Dokumentation](https://www.adobe.com/devnet-docs/uxp/uxp/reference-js/Modules/indesign/index.html), um die korrekte Funktionsweise der verwendeten Methoden und Eigenschaften sicherzustellen.
2. **Robustere Fehlerbehandlung**: Überprüfen Sie nach dem `place()`-Aufruf, ob das zurückgegebene Array nicht leer ist.

Korrigierter Code-Ausschnitt in der Funktion `verarbeiteGanzeSeiten`:
// ...
const platzierteObjekte = seite.place(neueDatei, [alteBounds[1], alteBounds[0]]);
if (platzierteObjekte && platzierteObjekte.length > 0) {
    const neuesObjekt = platzierteObjekte[0];
    neuesObjekt.geometricBounds = alteBounds;
    // ... Report aktualisieren
} else {
    report.details.fehlgeschlagenePlatzierungen.push({
        seitennummer: item.seitennummer,
        produktcode: item.produktcode,
        fehlergrund: "Fehler beim Platzieren: Die place()-Methode gab kein Objekt zurück."
    });
}
// ...

---

## 5. Fehler: Behandlung von Dateipfaden

### Beschreibung
Das Skript verwendet `datei.nativePath`, um die Pfade der Dateien anzuzeigen. Aus Sicherheitsgründen schränkt UXP den Zugriff auf die `nativePath`-Eigenschaft ein, und sie funktioniert nicht auf allen Plattformen zuverlässig. Plugins sollten stattdessen mit UXP-`Entry`-Objekten arbeiten, die die Dateisystempfade abstrahieren.

### Vorschlag
Vermeiden Sie die Verwendung von `nativePath`, wo immer es möglich ist. Wenn der Pfad dem Benutzer angezeigt werden muss, verwenden Sie die `name`-Eigenschaft des `Entry`-Objekts. Die Erstellung der Berichtsdatei in der Funktion `schreibeReport` ist kompliziert und fehleranfällig:
const ordner = await dokument.filePath.getEntries();
const ordnerReferenz = await fs.getEntryForPath(ordner[0].nativePath.replace(ordner[0].name, ''));

### Vorschlag
Vereinfachen Sie die Erstellung der Berichtsdatei. `dokument.filePath` ist bereits ein `Folder`-`Entry`-Objekt.

Korrigierter Code-Ausschnitt in der Funktion `schreibeReport`:
async function schreibeReport(reportDaten, dokument) {
    if (!dokument.saved) {
        throw new Error("Das Dokument muss gespeichert sein, um einen Report zu erstellen.");
    }
    const jetzt = new Date();
    const zeitstempel = `${jetzt.getFullYear()}${(jetzt.getMonth() + 1).toString().padStart(2, '0')}${jetzt.getDate().toString().padStart(2, '0')}${jetzt.getHours().toString().padStart(2, '0')}${jetzt.getMinutes().toString().padStart(2, '0')}`;
    const basisName = dokument.name.replace(/\.indd$/i, '');
    const reportName = `${basisName}${DATEINAME_TRENNER}${zeitstempel}.json`;
    
    // Dokumentenordner einfacher abrufen
    const dokumentOrdner = dokument.filePath; 
    if (!dokumentOrdner) {
        throw new Error("Konnte den Dokumentenordner nicht finden.");
    }

    const reportDatei = await dokumentOrdner.createFile(reportName, { overwrite: true });
    await reportDatei.write(JSON.stringify(reportDaten, null, 4), { format: "utf-8" });
    return reportName;
}

---

## 6. Fehler: Allgemeine Fehlerbehandlung

### Beschreibung
Das Skript enthält `try...catch`-Blöcke, aber nicht für jede asynchrone Operation ist eine umfassende Fehlerbehandlung implementiert. Ein Fehler in einer asynchronen Kette kann dazu führen, dass das Skript ohne klare Meldung an den Benutzer abstürzt.

### Vorschlag
Stellen Sie sicher, dass jeder `await`-Aufruf, der fehlschlagen könnte (z. B. Dateizugriffe, API-Aufrufe), in einem `try...catch`-Block steht, um den Fehler abzufangen und dem Benutzer eine verständliche Nachricht zu zeigen.

Beispiel für eine verbesserte asynchrone Fehlerbehandlung:
try {
    const produktliste = await leseProduktliste(einstellungen.datei);
    // ... weiterer Code
} catch (error) {
    console.error("Fehler beim Lesen der Produktliste:", error);
    await zeigeDialog(`Die Produktliste konnte nicht gelesen werden: ${error.message}`, "Lesefehler");
    return; // Workflow abbrechen
}

---

## Zusammenfassung

Die wahrscheinlichsten Probleme des Skripts sind:
1.  **Fehlende oder inkorrekte Modulimporte.**
2.  **Nicht deklarierte UI-Abhängigkeiten (Spectrum) in der `manifest.json`.**
3.  **Fehlende Dateisystemberechtigungen in der `manifest.json`.**
4.  **Änderungen und Nuancen in der InDesign-API unter UXP.**

Es wird dringend empfohlen, zuerst die `manifest.json`-Datei zu überprüfen und zu korrigieren, um sicherzustellen, dass alle notwendigen Berechtigungen und Abhängigkeiten enthalten sind. Anschließend sollten die Modulimporte und die Fehlerbehandlung im Skript selbst überarbeitet werden. Wenn das Skript danach immer noch nicht funktioniert, ist es ratsam, die Adobe UXP-Dokumentation oder die Community-Foren zu konsultieren.