# ReadEZSetup – ReadEZ installer and updates

WiX Toolset v5 project that builds `ReadEZ-Setup.msi` (same approach as Eye Dynamix's installer).
Visual Studio needs the free **HeatWave** extension to open `.wixproj` projects.

## Releasing a new version

1. Settings > Admin > **Publish update…**
2. Select **Increase** (or type the new version). This writes `<Version>` in `ReadEZ\ReadEZ.vbproj`.
3. Type what has changed, check the certificate details, and select **Build and publish**.
   This publishes ReadEZ, signs `ReadEZ.exe`/`ReadEZ.dll`, packs and signs the MSI, then writes
   `ReadEZ-Setup-<version>.msi` and `update.json` to the output folder and opens it.
4. Upload **both files** to `https://www.thomson-software-solutions.com/updates/readez/`
   (`UpdateService.DefaultUpdateUrl`), replacing the old `update.json`.

Installed copies check `update.json` when they start (and from Help > Check for updates), download
the MSI, check its SHA-256 and signature, close, install and restart.

## What an update changes

Only the program folder. Settings and the licence (`%AppData%\ReadEZ`) and the data folder
(`Documents\ReadEZ Data`: database, reading texts, letter templates) are never touched. The starter
set in the program's `Starter` folder is copied into the data folder only where a file is missing.

## Starter set

Put the files for a new installation in `ReadEZ\Deploy\Starter`:

    Starter\Reading texts\...        extra reading texts (the 15 standard texts are built in)
    Starter\ReadEZ Templates\...     ParentOpeningPara.txt, ParentClosingPara.txt,
                                     TeacherOpeningPara.txt, TeacherClosingPara.txt
    Starter\Report images\Logo.jpg   optional
    Starter\Database\ReadEZ.db       optional starter database (example data only)

## Help file

Build `Help\ReadEZ Help\HTML Help project\ReadEZHelp.hhp` with HTML Help Workshop (or import the topics into HelpNDoc - see `How to build the help file.html`)
and copy `ReadEZHelp.chm` into `ReadEZ\Deploy` so it is installed beside `ReadEZ.exe`.

## Building by hand

    set READEZ_SIGN_PASSWORD=...
    dotnet build ReadEZSetup\ReadEZSetup.wixproj -c Release -p:SignPfx="C:\path\cert.pfx" -p:SignTool="C:\...\signtool.exe"

Without `SignPfx` the MSI is unsigned (fine for testing; installed copies then refuse signed
updates and vice versa).
