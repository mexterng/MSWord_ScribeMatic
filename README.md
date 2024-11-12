# MSWord_ScribeMatic
*english version below*

Mit Hilfe des ScribeMatic-Makros können zwei Texte in Microsoft Word miteinander verglichen und Unterschiede gekennzeichnet werden. Es eignet sich besonders zur Korrektur und Überprüfung von Textabschriften.

# Voraussetzungen

## Software
*english version below*

- **Microsoft Word** (getestet ab Version 2010)

## Dokumentenaufbau
*english version below*

- Das zu prüfende Dokument muss reinen Text enthalten (nicht in Tabellen oder Textfeldern).
- Der Referenztext muss als Textdatei im ANSI-Format vorliegen. Zeilenumbrüche werden dabei ignoriert.

# Installation
*english version below*

1. Quellcode herunterladen.
   - Darauf achten, dass die Codierung korrekt ist. (Sonderzeichen in etwa Zeile 281). Ggf. den Code direkt aus dem Editor in GitHub kopieren.
2. Makro in Microsoft Word importieren.

## Makro hinzufügen
*english version below*

### Option 1: Global (für alle Dokumente)
- Entwicklertools → Visual Basic → `Normal` → `Module` → Rechtsklick → Datei importieren → `ScribeMatic.bas` auswählen.

### Option 2: Lokal (für aktuelles Dokument)
- Entwicklertools → Visual Basic → `Project (<Dokumentname>)` → `Module` → Rechtsklick auf `Module` → Datei importieren → `ScribeMatic.bas` auswählen.

⚠️ Word-Dokument anschließend im Format `.docm` (Dokument mit Makros) speichern.

## Makro zum Menüband hinzufügen
*english version below*

1. **Word-Optionen** → Menüband anpassen.
2. **Befehle auswählen**: Makros und gewünschte Registerkarte und Gruppe auswählen.
   - Optional: Beschriftung und Icon anpassen. ![grafik](https://github.com/user-attachments/assets/2b452ede-53f2-4729-a8fd-dfbc674d2fc0)

# Verwendung
*english version below*

1. Textabschnitt markieren, der mit dem Referenztext verglichen werden soll. Dieser sollte mit dem Referenztext übereinstimmen, kann jedoch kürzer sein.
2. Im sich öffnenden Fenster die Referenzdatei (im ANSI-Format) auswählen.
3. Anzahl der erwarteten Zeichenanschläge angeben.
   - Falls der Referenztext mehr Zeichen enthält, wird nur der markierte Textabschnitt verglichen. Fehlende Zeichen im Dokument werden als solche markiert.
4. Alle nötigen Änderungen werden in Dokument markiert.

## Funktionsweise des Makros
*english version below*

Das Makro `ScribeMatic()` markiert Unterschiede im Vergleich zum Referenztext. Die Korrekturzeichen sind wie folgt:

- **Insert**: Einfügen eines Zeichens, markiert durch senkrechte und waagerechte Linien.
- **Replace**: Ersetzen eines Zeichens, unterstrichen, mit dem richtigen Zeichen darüber notiert.
- **Delete**: Löschen eines überflüssigen Zeichens, durchgestrichen.

# Anwendungsfälle
*english version below*

- **Abschriftkorrektur**: Für Übungen mit zeitgesteuerten Abschriften (z. B. 10 Minuten für 700 Anschläge).
  - Schülertext markieren und Makro starten.
  - Referenzdatei auswählen und Zielanzahl angeben.
  - Fehler werden automatisch in Textfeldern markiert, eine manuelle Korrektur ist möglich.

# Mitwirken

Es kann ein Issue oder ein Pull Request geöffnet werden.

-------
-------
-------

# MSWord_ScribeMatic

The ScribeMatic macro allows for comparing two texts in Microsoft Word and marking differences. It's particularly useful for correcting and reviewing transcribed texts.

# Requirements

## Software

- **Microsoft Word** (tested starting from version 2010)

## Document Structure

- The document to be checked must contain plain text only (no tables or text boxes).
- The reference text must be in an ANSI text file. Line breaks are ignored.

# Installation

1. Download the source code.
   - Ensure correct encoding, as special characters may appear in the code starting around line 281.
2. Import the macro into Microsoft Word.

## Adding the Macro

### Option 1: Globally (for all documents)
- Developer Tools → Visual Basic → `Normal` → `Module` → Right-click → Import File → select `ScribeMatic.bas`.

### Option 2: Locally (for the current document)
- Developer Tools → Visual Basic → `Project (<Document name>)` → `Module` → Right-click on `Module` → Import File → select `ScribeMatic.bas`.

⚠️ Save the Word document in `.docm` format (document with macros) afterward.

## Adding the Macro to the Ribbon

1. **Word Options** → Customize Ribbon.
2. **Choose commands**: Select Macros and add to the desired tab and group.
   - Optional: Customize label and icon. ![grafik](https://github.com/user-attachments/assets/2b452ede-53f2-4729-a8fd-dfbc674d2fc0)

# Usage

1. Select the text section to compare with the reference text. It should match the reference text but can be shorter.
2. In the dialog box, select the reference file (in ANSI format).
3. Enter the expected number of keystrokes.
   - If the reference text has more characters, only the selected text section will be compared. Missing characters in the document will be marked accordingly.
4. All necessary changes will be marked in the document.

## Macro Functionality

The `ScribeMatic()` macro marks differences compared to the reference text. Correction symbols are as follows:

- **Insert**: Insertion of a character, marked by vertical and horizontal lines.
- **Replace**: Replacement of a character, underlined with the correct character written above.
- **Delete**: Deletion of an extra character, marked with a strikethrough.

# Use Cases

- **Transcription Correction**: For timed transcription exercises (e.g., 10 minutes for 700 keystrokes).
  - Select the student’s text and start the macro.
  - Choose the reference file and enter the target keystroke count.
  - Errors are automatically marked in text boxes, allowing manual correction.

# Contributing

Feel free to open an issue or a pull request.
