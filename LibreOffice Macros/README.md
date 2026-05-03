# LibreOffice Sanskrit Conversion Macros

LibreOffice Basic macros for converting Sanskrit transliteration between Balaram ASCII encoding, IAST Unicode diacritics, Harvard-Kyoto, Russian Cyrillic, and Ukrainian Cyrillic.

## Available macros

| Macro | Description |
|---|---|
| `BalaramToUnicode` | Balaram → IAST Unicode (ä → ā, ö → ṭ, etc.) |
| `UnicodeToBalaram` | IAST Unicode → Balaram |
| `BalaramToUkrainianCyrillic` | Balaram → Ukrainian Cyrillic transliteration |
| `UkrainianCyrillicToBalaram` | Ukrainian Cyrillic → Balaram |
| `HarvardKyotoToUkrainianCyrillic` | Harvard-Kyoto → Ukrainian Cyrillic |
| `BalaramToNoDiacritics` | Balaram → plain ASCII (strips all diacritics) |
| `BalaramToRussianCyrillic` | Balaram → Russian Cyrillic transliteration |
| `RussianCyrillicToBalaram` | Russian Cyrillic → Balaram |
| `HarvardKyotoToRussianCyrillic` | Harvard-Kyoto → Russian Cyrillic |

## Selection support

All macros work in two modes:

- **Text selected** — converts only the selected text
- **No selection** — converts the entire document

## Files

- `SansConverterMacro_selections.bas` — full version with selection support (recommended)
- `SansConverterMacro_whole_file.bas` — experimental version, always converts the whole document

## Installation

LibreOffice macros must be installed by pasting the code directly into the LibreOffice Basic IDE. The install script in the repo root does **not** work reliably due to LibreOffice's internal module format.

**Steps:**

1. Open LibreOffice Writer
2. Go to **Tools → Macros → Edit Macros** (or press Alt+F8 then click Edit)
3. In the left panel, expand **My Macros → Standard**
4. Click on an existing module (e.g. `Module1`), or create a new one via **Insert → Module**
5. Select all existing code in the editor (Cmd+A / Ctrl+A) and delete it
6. Open `SansConverterMacro_selections.bas` in a text editor, copy all contents
7. Paste into the LibreOffice Basic IDE editor
8. Click the **Compile** button (checkmark icon) to verify no errors
9. Close the IDE

**Running a macro:**

- Go to **Tools → Macros → Run Macro**
- Navigate to **My Macros → Standard → Module1** (or whatever module name you used)
- Select the macro and click **Run**

**Tip:** Assign macros to keyboard shortcuts via **Tools → Customize → Keyboard** for faster access.

## Adding macros to the toolbar

For one-click access, you can add buttons for each macro to the LibreOffice toolbar:

1. Go to **Tools → Customize → Toolbars**
2. In the **Toolbar** dropdown, select the toolbar you want to add buttons to (e.g. *Standard*)
3. Click **Add Command...**
4. In the **Category** list, scroll down and select **Macros → My Macros → Standard → Module1**
5. In the **Commands** list, select the macro you want (e.g. `BalaramToUnicode`)
6. Click **Add**, then repeat for each macro you want on the toolbar
7. Click **OK**

To change the button icon: right-click the new button in the toolbar → **Customize Toolbar** → select the button → **Change Icon**, then pick or import an icon.

The macro buttons will appear in the toolbar and can be clicked directly — with text selected they convert the selection, otherwise they convert the whole document.
