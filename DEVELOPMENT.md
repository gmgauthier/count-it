# Count-It development plan

A gtkmm-3 spreadsheet for LCOS. The window is Excel 97’s grid. The file is a simple `.xlsx`.

Display name: **Count-It**  
Binary / repo / package: `count-it`  
APP_ID: `org.gmgauthier.CountIt`  
License: The Unlicense (`UNLICENSE`)  
Repos: https://gitea.scriptorium/gmgauthier/count-it (origin), https://github.com/gmgauthier/count-it

## Status (2026-10-07)

**Specification.** This repository holds the plan. Source begins at M0, after Write-It 1.0 has been lived with. 1.0 is M0 through M5. Live with that release before adding a function. Tag `v1.0.0` at M5.

The 960×700 first-launch mockup is [brand/window.png](brand/window.png). The sample workbook in that picture is `household.xlsx`. The formula bar shows `B6` and `=AVG(B2:B4)`.

## 1. Locked decisions

| Decision | Choice |
|---|---|
| Product | Original. The window is Excel 97. Gnumeric is the feature floor to study |
| Name | Count-It. Binary `count-it` |
| Toolkit | C++17, gtkmm-3.0, GTK3 CSS, Meson (`warning_level=2`) |
| Look | One decorated window. Clearlooks-Phenix draws the controls. The window manager draws the title bar |
| File | A simple `.xlsx` (SpreadsheetML, ECMA-376), written with libarchive and libxml2 |
| Formulas | Start with `=`. The bar shows `AVG`. The file stores `AVERAGE`. `AVG` is the arithmetic mean only |
| Fallback | CSV of displayed values, one sheet. Save still writes `.xlsx` |
| Init | No systemd. Config `~/.config/count-it/count-it.ini` |
| License | The Unlicense |
| Versioning | Semantic (`MAJOR.MINOR.PATCH`). When source exists, `meson.build` is the source of truth. The Debian changelog and git tag `vX.Y.Z` match it |

## 2. Place in the suite

Retro-Office is three applications with one window language: Write-It (`write-it`), Count-It, and Show-It (`show-it`). Each is its own repository. The umbrella note lives in the local `lcos-projects` folder as `RETRO-OFFICE.md` and is not part of this repository. This file is the Count-It specification.

LCOS already ships AbiWord and Gnumeric. The gap this suite fills is coherence. Study Gnumeric. Do not fork it, and do not shell out to `ssconvert`. Gnumeric is GPL. The house license is the Unlicense. Its weight sits in the function library and the import filters.

Write-It is the first codebase. This plan is the second. Show-It waits until Count-It 1.0 has been lived with.

A calculator stays galculator. Count-It is a sheet. Organized notes stay in the Ephemeris Notepad.

## 3. House rules

- Devuan Excalibur / LCOS, XLibre, XFCE, Clearlooks-Phenix
- No systemd, no PackageKit, no custom title bar, no daemon, no online account, no AI
- Local files only
- Borrow the LCOS palette. Do not use Bryan’s seal
- Ship `.deb`, source tarball, and AppImage at M5
- Tests headless and offline. Lint covers `src/` only

## 4. Window

One workbook, one window. The title is `Count-It - household.xlsx`. A dirty workbook adds a trailing `*`. A new workbook is `Count-It - Untitled`.

Closing a dirty workbook asks one question. The buttons, in order, are **Save**, **Don’t Save**, **Cancel**. Save is the default. **Close** (Ctrl+W) returns to Untitled. **Exit** (Ctrl+Q) leaves the program. The first launch is 960×700, not maximized. The window remembers its size.

The grid fills its area and uses the theme’s base colour. Write-It’s gray pasteboard is not this surface.

```
+------------------------------------------------------------------+
| File  Edit  View  Insert  Format  Tools  Data  Help              |
+------------------------------------------------------------------+
| [New] [Open] [Save] | [Print] | [Cut] [Copy] [Paste] | [Undo] [Redo]
| [Font ▾] [Size ▾] [B] [I] [U] | [Left] [Center] [Right] | [General ▾]
| Name [B6 ▾]   fx [ =AVG(B2:B4)                                  ] |
+----+----------+------------------+-------------------------------+
|    | A        | B                | C                             |
|  1 |          |                  |                               |
|  2 | Groceries|             86.40|                               |
+----+----------+------------------+-------------------------------+
| Sheet1 | Sheet2                                                  |
+------------------------------------------------------------------+
| Ready                            Sheet1          B6         100% |
+------------------------------------------------------------------+
```

Icons come from the desktop icon theme, by freedesktop name: `document-new`, `document-open`, `document-save`, `document-print`, `edit-cut`, `edit-copy`, `edit-paste`, `edit-undo`, `edit-redo`, `format-text-bold`, `format-text-italic`, `format-text-underline`, `format-justify-left`, `format-justify-center`, `format-justify-right`. Toolbars are icons. The menu’s words are the tooltip. A missing icon falls back to that short word. A toolbar combo that applies a format returns focus to the grid afterward.

Cut, Copy, Paste, Undo, and Redo are insensitive when there is nothing to do. Save stays sensitive. The right-click menu starts with Cut, Copy, Paste, then a separator, then this app’s own items.

### Menus

The menus are File, Edit, View, Insert, Format, Tools, Data, Help. Mnemonics: **F**ile, **E**dit, **V**iew, **I**nsert, F**o**rmat, **T**ools, **D**ata, **H**elp. A menu item that opens a dialog ends with `…`. Accelerators are visible in the menu.

**File.** Open’s filter lists `.xlsx`, CSV, and plain text. Save writes `.xlsx`. Export writes CSV of displayed values.

| Item | Keys | What it does |
|---|---|---|
| New | Ctrl+N | A blank untitled workbook |
| New from Template… | | Pick a starter workbook of the native type |
| Open… | Ctrl+O | Remember the last directory |
| Open Recent | | Up to eight basename items. The tooltip is the full path. A missing file uses one sentence: “That file is missing.” |
| Save | Ctrl+S | Write the whole `.xlsx` |
| Save As… | | |
| Export… | | CSV of displayed values |
| Print… | Ctrl+P | The system print dialog |
| Page Setup… | | Paper, orientation, margins |
| Close | Ctrl+W | Back to Untitled |
| Exit | Ctrl+Q | |

**Edit.**

| Item | Keys |
|---|---|
| Undo | Ctrl+Z |
| Redo | Ctrl+Y, and Ctrl+Shift+Z |
| Cut | Ctrl+X |
| Copy | Ctrl+C |
| Paste | Ctrl+V |
| Delete | Delete. Clears the selected cell contents |
| Select All | Ctrl+A |
| Find… | Ctrl+F |
| Replace… | Ctrl+H |

Find and Replace are one modal dialog. Fields, in order: Find, Replace, a Match case check, “Look in formulas,” Next, Replace, Close.

**View.**

| Item | Behaviour |
|---|---|
| Standard Toolbar | Check. On by default |
| Format Toolbar | Check. On by default |
| Status Bar | Check. On by default |
| Zoom | Submenu: 50%, 75%, 100%, 150%, 200% |
| Gridlines | Check. On by default |

**Insert.** Sheet, Chart…, Comment. Count-It has no picture in v1.

**Format.** Font…, Bold (Ctrl+B), Italic (Ctrl+I), Underline (Ctrl+U), Align Left, Center, Align Right, then Number Format…. Font… is family, size, bold, italic, underline. No colour in v1.

The font list and the size list match the other two apps. Sizes are 8, 9, 10, 11, 12, 14, 16, 18, 24, 36. A new workbook starts at Sans 11.

**Tools.** Options…: default font family, default size, recent-file count (4, 8, or 12). Spelling is Write-It’s item. The Tools menu stays so the menu bar does not shift.

**Data.** Sort…, Autofilter.

**Help.** About Count-It. The dialog shows the program name, the version, one sentence, the Unlicense, and Close.

### Keyboard

The only keyboard is the Excel 97 map in the menus above. The Lotus 1-2-3 slash menu stays out of v1. There is no Options switch for it. Formulas start with `=`. The `@` prefix stays out. F10 is the GTK menu key. Alt+F4 closes the window. F1 is Help. F5 is unused here. F7 is unused here.

### Toolbars

The standard toolbar never grows an app-specific button. Groups, left to right: New Open Save, Print, Cut Copy Paste, Undo Redo.

The format toolbar: font, size, bold, italic, underline, align left, align center, align right, then a separator, then the number-format combo: General, Number, Currency, Date, Percent, Text.

The formula bar is a full-width band directly under the format toolbar: a name box of fixed width, an `fx` label, and a formula entry that takes the remaining width. Enter commits and moves the selection down. Esc cancels the edit. The font combo does not behave this way.

### Status bar

The left side is a message (“Ready”) that stays until the next message. The rightmost cell is the zoom, and it pops the same list as View. To its left is the current cell (`B6`). To the left of that is the sheet name (`Sheet1`).

### Config

`~/.config/count-it/count-it.ini`

Keys: `window-width`, `window-height`, `recent`, `last-dir`, `default-font`, `default-size`, `show-standard-toolbar`, `show-format-toolbar`, `show-statusbar`, `zoom`.

## 5. Feature floor

Taken from Gnumeric’s own surface:

- A grid, several sheets, a formula bar, and a name box
- Fill down, series fill, and the fill handle
- Named ranges
- Number, currency, date, time, percent, and text formats
- Fonts, borders, alignment, wrap, and merged cells
- Comments
- Sort and autofilter
- One chart, column or pie
- Print, including gridlines and headings
- CSV in both directions

Column letters, row numbers, sheet tabs, and the fill handle are the grid.

Gnumeric aims to cover the North American Excel worksheet functions and adds many of its own. It has no pivot tables and no VBA. Count-It follows that refusal from the start. Full Excel parity is a later program.

## 6. Formulas

A formula starts with `=`. The v1 set is `SUM`, `AVG`, `MIN`, `MAX`, `COUNT`, `IF`, `ROUND`, `DATE`, `TODAY`, `YEAR`, `MONTH`, and `DAY`.

`AVG` and `AVERAGE` are two spellings of one spreadsheet function. It returns the arithmetic mean: add the numbers, divide by how many numbers there are. Text and blank cells are skipped. In mathematics the unqualified word “mean” also covers the geometric mean and the harmonic mean, and the unqualified word “average” also covers a median or a mode. Those are different operations. They are not this function, and they are not in v1.

The formula bar shows one spelling. The `.xlsx` stores the name Excel and Gnumeric calculate. On commit, an accepted alias is rewritten to the shown name, and the file receives the stored name.

| Shown after `=` | Also accepted | Stored in the `.xlsx` |
|---|---|---|
| `SUM` | | `SUM` |
| `AVG` | `AVERAGE` | `AVERAGE` |
| `MIN` | | `MIN` |
| `MAX` | | `MAX` |
| `COUNT` | | `COUNT` |
| `IF` | | `IF` |
| `ROUND` | | `ROUND` |
| `DATE` | | `DATE` |
| `TODAY` | | `TODAY` |
| `YEAR` | | `YEAR` |
| `MONTH` | | `MONTH` |
| `DAY` | | `DAY` |

`SUM`, `MIN`, `MAX`, and `DAY` are already three letters in both programs. `COUNT`, `IF`, `ROUND`, `DATE`, `TODAY`, `YEAR`, and `MONTH` keep those names. Lotus used them too, and they are longer than three letters. Invented squeezes such as `MEN`, `CNT`, `RND`, `TDY`, and `MON` stay out. The `@` prefix stays out. `=AVG(B2:B9)` is the arithmetic mean.

## 7. Format

The file Count-It saves is **`.xlsx`** (SpreadsheetML, ECMA-376). It is a zip of XML, written with libarchive and libxml2. One part lists the sheets. Each sheet is a grid. A formula is stored as text, such as `SUM(B2:B9)`, beside the last calculated value. Number formats, fonts, borders, alignment, merged cells, named ranges, and an autofilter live in that same package. Comments are a comments part in the same workbook. One chart, a single column series or a single pie, is one chart part. Excel, Gnumeric, and LibreOffice Calc open this slice.

The v1 contract is the `.xlsx` Count-It writes, plus a straightforward workbook that uses that same slice. Pivot tables, macros, and the rest of the function catalogue stay out. `.xls`, `.ods`, and `.gnumeric` wait. An existing `.gnumeric` file is a later importer. `ssconvert` is not the product.

**CSV is the fallback, and it is lossy.** Export writes the values as displayed, one sheet, commas between cells. Import lays those values onto a sheet. Formulas, further sheets, formatting, and the chart stay behind. CSV export is a separate command, so Save does not drop them. Plain text is not a second formula language.

## 8. Work plan

1.0 is M0 through M5, in this order. Count-It is the second Retro-Office codebase. Implementation starts after Write-It 1.0 has been lived with. The next milestone starts when the current one's done line is true. Live with that release before adding a function. M5 cuts `v1.0.0`.

Each milestone is a branch `feature/mN-short-name` from `master`. A milestone that owns a file format, a grid operation, or a formula brings a headless offline test for that slice. The CHECK harness is the one the other guests use. Lint covers `src/` only.

The sections above are the specification. This section is the order of work. [brand/window.png](brand/window.png) is the chrome target at M0. The household numbers and `=AVG(B2:B4)` arrive with the milestones that own values and formulas.

| Milestone | Done when |
|---|---|
| **M0 — Window** | Menus, toolbars, formula bar, name box, empty grid, one sheet tab, About. Matches the sketch. |
| **M1 — File** | New / Open / Save `.xlsx` for values. CSV import and export of displayed values. Recent files. |
| **M2 — Grid** | Edit, fill down, series fill, formats, fonts, borders, alignment, wrap, merge, comments. |
| **M3 — Formulas** | The function set above. The bar shows `AVG`. The file stores `AVERAGE`. Named ranges. Several sheets. |
| **M4 — Arrange and print** | Sort, autofilter, one column or pie chart written into the `.xlsx`, print with grid and headings. |
| **M5 — 1.0** | `debian/`, `scripts/release.sh` → `.deb`, tarball, AppImage. Tag `v1.0.0` and publish it. |

### M0 — Window

The Meson tree, the gtkmm window, and `scripts/lint.sh`. No workbook on disk.

- Menus in order: File, Edit, View, Insert, Format, Tools, Data, Help, with the mnemonics from the window section. Items are visible. Commands that need a workbook are insensitive. Save stays sensitive.
- Standard toolbar, then the format toolbar through alignment, then the number-format combo showing General. The combo’s formats wait for M2.
- Formula bar under that toolbar: name box, `fx`, and the entry. The entry does not commit yet.
- An empty grid with column letters, row numbers, and one sheet tab, Sheet1. Arrow keys move the selection. The name box and the status cell follow it, starting at A1. Gridlines are on.
- Title `Count-It - Untitled`. First launch 960×700. The ini remembers `window-width` and `window-height`.
- Status message, then the sheet name, then the cell, then the zoom.
- About Count-It: name, version, one sentence, the Unlicense, Close.
- Close (Ctrl+W) and Exit (Ctrl+Q). The right-click menu starts with Cut, Copy, Paste.

**Done when** the line in the table is true and the window matches [brand/window.png](brand/window.png) with an empty grid.

### M1 — File

Values round-trip. Typing into a cell waits until M2, so this milestone is opened and saved workbooks, not the editor.

- New, Open, Save, and Save As write a `.xlsx` of cell values, with libarchive and libxml2. One sheet. A formula is not stored yet. The grid paints the stored values.
- Dirty state is a trailing `*` on the title. Closing a dirty workbook asks Save, Don’t Save, Cancel, with Save as the default.
- Export writes CSV of the displayed values, one sheet. Import lays CSV values onto a sheet. Export is a different command from Save.
- Open Recent, up to eight names, tooltip the full path, and the sentence “That file is missing.” Options… can set the recent-file count to 4, 8, or 12.

**Done when** the M1 line in the table is true. A headless test writes values, reads them back from the `.xlsx`, exports the displayed CSV, and imports that CSV onto a sheet.

### M2 — Grid

- Edit a cell in the grid and from the formula bar. Enter commits and moves the selection down. Esc cancels. A committed literal is a value, not a formula.
- Fill down, series fill, and the fill handle.
- Number, currency, date, time, percent, and text formats, from Number Format… and from the format-toolbar combo (General, Number, Currency, Date, Percent, Text).
- Fonts, borders, alignment, wrap, and merged cells. A new workbook starts at Sans 11. No font colour. Options… gains the default font family and size.
- Comments. Insert → Comment. The comment is a comments part in the workbook.
- Find and Replace, one dialog, in the locked field order, with “Look in formulas” present and idle until M3.
- The `.xlsx` from M1 gains these formats, the merge, and the comment.

**Done when** the M2 line in the table is true. A headless test round-trips a formatted, merged, commented sheet, and checks fill down and a series fill.

### M3 — Formulas

The function set in the formulas section: `SUM`, `AVG`, `MIN`, `MAX`, `COUNT`, `IF`, `ROUND`, `DATE`, `TODAY`, `YEAR`, `MONTH`, and `DAY`.

- A formula starts with `=`. The bar shows `AVG`. The file stores `AVERAGE`. Typing `AVERAGE` rewrites to `AVG` on commit. `AVG` is the arithmetic mean only: add the numbers, divide by how many numbers there are, skip text and blanks.
- The formula is stored as text beside the last calculated value. Dependents recalculate when a value they read changes.
- Named ranges. The name box shows a name when the selection is that range, and choosing a name selects it.
- Several sheets. Insert → Sheet. Sheet tabs switch the grid. The status bar names the current sheet.
- “Look in formulas” searches the stored formula text.

**Done when** the M3 line in the table is true. A headless test calculates each function, including `=AVG` over a range that contains a blank and a text cell, writes `AVERAGE` into the `.xlsx`, and reads that file back onto the bar as `AVG`.

### M4 — Arrange and print

- Data → Sort… and Data → Autofilter. The autofilter is stored in the workbook.
- Insert → Chart…. One chart, a single column series or a single pie, written as one chart part.
- Print… opens the system print dialog and can include gridlines and headings. Page Setup… sets paper, orientation, and margins.

**Done when** the M4 line in the table is true. A headless test round-trips a sort, an autofilter, and both the column chart and the pie chart in the `.xlsx`.

### M5 — 1.0

- `debian/`, a desktop file for `org.gmgauthier.CountIt`, and `scripts/release.sh`.
- The script produces the source tarball, the amd64 `.deb`, and the AppImage. The desktop `Name=` is the AppImage’s name.
- Tag `v1.0.0` after `meson test` and lint are green. Publish the tag and the three artifacts to Gitea and GitHub.

**Done when** `v1.0.0` is tagged and the three artifacts are on both remotes. Live with that release before adding a function. `.xls`, `.ods`, `.gnumeric`, pivot tables, and VBA stay out of 1.0.

## 9. Traps

- Forking Gnumeric, or shelling out to `ssconvert` as the product
- The Excel function catalog, pivot tables, VBA, array formulas, solver
- Treating geometric mean, harmonic mean, median, or mode as `AVG`
- Invented three-letter names (`MEN`, `CNT`, `RND`, `TDY`, `MON`), or an `@` before a function name
- A Lotus 1-2-3 slash menu
- `.xls`, `.ods`, or `.gnumeric` as a v1 deliverable
- Reading every `.xlsx` Excel can save
- A second chart type, before column and pie have been lived with
