Files in this folder are copied next to ReadEZ.exe when the program is built.

  ReadEZHelp.chm      the help file. Build it from "Help\ReadEZ Help" (HTML Help
                      Workshop: HTML Help project\ReadEZHelp.hhp, or HelpNDoc) and copy the .chm here. Without it,
                      Help opens the online help (Admin > Online help) or shows "help file missing".

Starter set (installed in "<program>\Starter"; copied into the user's data folder,
Documents\ReadEZ Data, only where files are missing, so updates never replace changed files):
  Starter\ReadEZ Templates\   the four letter paragraphs (from ReadEZ 2020's install files)
  Starter\Reading texts\      extra reading texts (the 15 standard texts are built in)
  Starter\Report images\      optional Logo.jpg
  Starter\Database\ReadEZ.db  optional starter database (example data only - Admin > Publish
                              update can put one in the release for you)

Product licence (activation): stored in %AppData%\ReadEZ\Licence.json. If ReadEZ 2020's tssRZ.dll
is found (next to the program or in ReadEZ 2020's folder) it is imported the first time.
