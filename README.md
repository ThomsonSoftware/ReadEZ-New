# ReadEZ 3

ReadEZ screens for visual stress and finds the preferred colour for ReadEZ overlays, ReadEZ tinted
lenses and computer screens. This version is a rewrite of ReadEZ 2020 in **VB.NET / .NET 8 WinForms**
on the Eye Dynamix framework: same look and feel (navigation rail, top bar, cards, Light / Dark / Blue
themes), same languages (English, Spanish, German, Italian), same one-or-two-screen handling,
installer, updates, licence, backups and help — without eye tracking.

## Folders

| Folder | What it is |
|---|---|
| `ReadEZ\` | The program (`ReadEZ.vbproj`) |
| `ReadEZSetup\` | WiX v5 installer (`ReadEZSetup.wixproj`, `Package.wxs`, `License.rtf`) — see its README |
| `Help\ReadEZ Help\` | Help manual: topic HTML, HTML Help Workshop project (`ReadEZHelp.hhp`) and HelpNDoc import list |
| `ReadEZ.sln` / `ReadEZ.slnx` | Solution (Debug builds only the program; Release also builds the installer) |

## Building

Visual Studio 2022 (17.12+) with the .NET 8 SDK, DevExpress 24.2 (printing only; set `DevExpressBin`
in `ReadEZ.vbproj` if it is not in the default folder) and, for the installer, the HeatWave extension.
Open `ReadEZ.sln` and build. Releases are made from inside the program: Settings > Admin >
Publish update (see `ReadEZSetup\README.md`).

Before the first release:

* build the help (`Help\ReadEZ Help\How to build the help file.html`) and copy `ReadEZHelp.chm` to `ReadEZ\Deploy`;
* review `ReadEZSetup\License.rtf` (copied from Eye Dynamix with the product name changed);
* upload releases to `https://www.thomson-software-solutions.com/updates/readez/` (`UpdateService.DefaultUpdateUrl`).

## Program structure (`ReadEZ\`)

| Folder | Contents |
|---|---|
| `AppHost.vb` | Start-up: settings, licence, database (and the offer to import ReadEZ 2020), main window |
| `Core\` | Models (`ReadEZResult`, patients, organisations, screeners), `SessionState` (current patient / test / screener), `ScreenLayout` (one or two screens), log |
| `Data\` | `ReadEZDatabase` (SQLite), `LegacyDatabaseImport` (ReadEZ 2020 database), `LegacyImport` (RZSettings.ini, palette), settings, backups, starter files |
| `UI\Shell\` | Main window, top bar, menus, commands, splash |
| `UI\Screens\` | Patients, Tests, Results, Colours (calibration), Settings, Help |
| `UI\Tests\` | Test runner (`TestRunScreen`, `TestRun`) and the steps: Symptoms, Reading (Wilkins rate of reading), Colour (automatic staircase / manual wheels) |
| `UI\Stimulus\` | The patient's screen (`StimulusWindow`), reading text layout (`ReadingScene`), the 15 standard reading texts |
| `UI\Dialogs\`, `UI\Framework\` | Dialogs and the shared Eye Dynamix controls/theme; `ReadEZColourCheck` is the colour calibration |
| `Reports\` | Report and parent/teacher letters (`ReadEZReport`, DevExpress XtraPrinting) |
| `Services\` | Licence (`RZ-` serials, imports `tssRZ.dll`), updates, help file and help topics, code signing |
| `Localisation\Strings.json` | All UI text in en / es / de / it |

### One or two screens

* **Two screens** — the patient's screen (`StimulusWindow`) covers its own monitor and stays in front while a
  test runs; the test screen shows an exact, clickable copy.
* **One screen** — the test screen shows the patient's text in a panel; **F11** shows it full screen with a
  small control strip (as ReadEZ 2020's single-screen bar), **Esc** returns.

### The test procedure

Baseline (symptoms → rate of reading) and colour tests for Overlay / Lenses / Computer
(colour → rate of reading with colour → symptoms with colour). The automatic colour test is a port of
ReadEZ 2020's algorithm: three interleaved staircases from A, E and H (≈33 presentations), comparison of
the winners, winner vs grey, and light vs dark tint for lenses/computer. Overlays are shown side by side;
lenses and computer successively (colour 1, colour 2 at the set presentation time, then white).
Reading speed = floor(60 × (words − 15 × rows skipped) / seconds).

## Database (`Documents\ReadEZ Data\Database\ReadEZ.db`, schema version 1)

| Table | Key columns |
|---|---|
| `Meta` | Key/Value: `SchemaVersion`, `ImportedFrom.<file>` |
| `Organisations` | Id, Name, Tel, Address1/2, Town, County, Postcode |
| `Screeners` | Id, Title, Firstname, Surname, Active (deleting only hides) |
| `Patients` | Id, names, DoB, Reference, contact details, OrganisationId → Organisations, LegacySource/LegacyId |
| `Tests` | Id, PatientId → Patients (ON DELETE CASCADE), ScreenerId, TestDate, **Mode** (0 baseline, 1 overlay, 2 computer, 3 lenses), ColourIndex, ColourSet (1 overlay, 2 lens, 3 dark lens), ColourRgb, GreyPreferred, Method (0 auto, 1 manual, 2 bespoke), AutoScore, Symptom1–8, SymptomsDone, SymptomNotes, Wpm, ReadSeconds, WordsRead, Errors, RowsSkipped, Transcript, text layout (TextIndex, TextName, FontName, FontSize, LineSpacing, RowCount, OverlaySize), PresentationMs, ScreenWidth/Height, Notes, LegacySource/LegacyId |

Baseline and colour tests share one table (ReadEZ 2020 had `Baseline` and `Results`), so a patient's
history is one indexed query. WAL journal, foreign keys on, indexes on patient name, organisation and
(PatientId, TestDate).

### Importing ReadEZ 2020

`LegacyDatabaseImport` reads `ReadEZ2020.db` (tables `Clients`, `Organisations`, `Tester`, `Baseline`,
`Results`) read-only and copies everything into the new database in one transaction, mapping the old
ids. Each imported row keeps `LegacySource` (the old file) and `LegacyId`, with unique indexes, so
importing the same file again adds only new rows. It is offered on first run (found via `RZSettings.ini`
or the usual ReadEZ 2020 folders) and available from File > Import a ReadEZ 2020 database, Settings >
Data and Admin. A backup is made first. `RZSettings.ini` (text, timings, screens, palette) is imported
from Admin > Maintenance.

## Data on the user's PC

| What | Where |
|---|---|
| Settings, admin settings, licence, logs | `%AppData%\ReadEZ` |
| Database, reading texts, letter templates, report images | `Documents\ReadEZ Data` |
| Program, help, starter set | `Program Files\Thomson Software Solutions\ReadEZ` |
