Option Strict On

Imports System.IO

' ====================================================================================================
'   HelpFile.vb - the help manual (ReadEZHelp.chm, made with HelpNDoc)
'
'   The CHM file is installed in the program folder: put it in ReadEZ\Deploy (everything in
'   Deploy is copied next to the program when it is built, and the installer installs the whole
'   program folder). ReadEZHelp.chm is used, or failing that any other .chm file there.
'
'   F1 and "Read in the help" open the topic for what is on the screen. The program asks for
'   topics by its own ids (screen.patients, test.colour, test.reading ... - the HelpCatalog ids,
'   also used in the help project made for HelpNDoc); HelpNDoc names each page of the CHM after
'   the topic's Help ID (Help ID + ".htm"), which it made from the topic title on import. The
'   table below links the two. If a topic's Help ID is changed in HelpNDoc, change it here too.
'   A topic that isn't in the table (e.g. one added in Admin) opens the topic of the group it is
'   in, or the contents.
' ====================================================================================================

Public NotInheritable Class HelpFile
    Private Sub New()
    End Sub

    Public Const FileName As String = "ReadEZHelp.chm"

    ''' <summary>Program topic id -> HelpNDoc Help ID (the page name in the CHM, without .htm).</summary>
    Private Shared ReadOnly Pages As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"welcome", "WelcometoReadEZ"},
        {"start", "Gettingstarted"},
        {"start.install", "Installingandactivating"},
        {"start.setup", "Settingupthescreens"},
        {"start.tour", "Tourofthemainwindow"},
        {"start.environment", "Thetestenvironment"},
        {"screens", "Tourofthemainwindow"},
        {"screen.patients", "Patients"},
        {"screen.tests", "Runningatest"},
        {"tests.guide", "Astepbystepguide"},
        {"tests.run", "Oneortwoscreens"},
        {"tests", "Runningatest"},
        {"test.symptoms", "Symptoms"},
        {"test.reading", "Rateofreading"},
        {"test.colour", "Determinationofthepreferredcolou"},
        {"screen.results", "Results"},
        {"results.report", "Reportsandletters"},
        {"screen.colours", "Colourcalibration"},
        {"screen.settings", "Settings"},
        {"settings.general", "General"},
        {"settings.appearance", "Appearanceandlanguage"},
        {"settings.data", "Dataandbackup"},
        {"settings.screens", "Screens"},
        {"settings.testsetup", "Testsetup"},
        {"import", "ImportingfromReadEZ2020"},
        {"licence", "Licenceandupdates"},
        {"help", "Helpandsupport"},
        {"admin", "Administratortools"},
        {"reference", "Reference"},
        {"ref.keys", "Keyboardshortcuts"},
        {"ref.trouble", "Troubleshooting"},
        {"ref.files", "WhereReadEZkeepsitsfiles"}
    }

    ''' <summary>The help file in the program folder ("" if there is none).</summary>
    Public Shared ReadOnly Property FilePath As String
        Get
            Try
                Dim f = AppHost.FileInAppFolder(FileName)
                If File.Exists(f) Then Return f
                Return If(Directory.EnumerateFiles(AppHost.AppFolder, "*.chm").FirstOrDefault(), "")
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property

    Public Shared ReadOnly Property Available As Boolean
        Get
            Return FilePath <> ""
        End Get
    End Property

    ''' <summary>The CHM page for a program topic id (its group's page if it has none), or "".</summary>
    Public Shared Function PageFor(topicId As String) As String
        Dim id = If(topicId, "")
        Dim page As String = Nothing
        While id <> ""
            If Pages.TryGetValue(id, page) Then Return page.ToLowerInvariant() & ".htm"     ' HelpNDoc writes the page names in lower case
            Dim parent = HelpCatalog.Find(id)?.Parent
            id = If(parent Is Nothing, "", parent.Id)
        End While
        Return ""
    End Function

    ''' <summary>Whether the help has its own page for this topic (or for its group).</summary>
    Public Shared Function HasPage(topicId As String) As Boolean
        Return Available AndAlso PageFor(topicId) <> ""
    End Function

    ''' <summary>
    ''' Opens the help at a topic (the contents when topicId is "" or has no page). Returns False
    ''' if there is no help file.
    ''' </summary>
    Public Shared Function Show(owner As IWin32Window, Optional topicId As String = "") As Boolean
        Dim f = FilePath
        If f = "" Then Return False
        Dim parent = If(TryCast(owner, Control), CType(AppHost.MainWindow, Control))
        Try
            Dim page = PageFor(topicId)
            If page <> "" Then
                Help.ShowHelp(parent, f, HelpNavigator.Topic, page)
            Else
                Help.ShowHelp(parent, f, HelpNavigator.TableOfContents)
            End If
        Catch ex As Exception
            Log.Error("Help file " & f, ex)
            AppHost.OpenUrl(f)
        End Try
        Return True
    End Function

    ''' <summary>The help's index, for searching by keyword.</summary>
    Public Shared Function ShowIndex(owner As IWin32Window) As Boolean
        Dim f = FilePath
        If f = "" Then Return False
        Try
            Help.ShowHelpIndex(If(TryCast(owner, Control), CType(AppHost.MainWindow, Control)), f)
        Catch ex As Exception
            Log.Error("Help file " & f, ex)
        End Try
        Return True
    End Function
End Class
