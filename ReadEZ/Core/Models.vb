Option Strict On

' ====================================================================================================
'   Models.vb - the data ReadEZ keeps
'
'   Stored in the ReadEZ database (ReadEZDatabase.vb):
'     PatientStructure      a row of Patients
'     OrganisationStructure a row of Organisations (school, practice ...)
'     ScreenerStructure     a row of Screeners (the person running the test)
'     ReadEZResult          a row of Tests: a baseline (white) or a colour test
'   ReadEZ 2020's Clients, Tester, Baseline and Results are brought across by LegacyDatabaseImport;
'   the field names below follow ReadEZ 2020 so the ported code reads the same.
' ====================================================================================================

''' <summary>A patient (table Patients).</summary>
Public Structure PatientStructure
    Public ID As Integer
    Public Title As String
    Public Firstname As String
    Public Surname As String
    Public DoB As Date
    Public Reference As String
    Public Address1 As String
    Public Address2 As String
    Public Town As String
    Public County As String
    Public Postcode As String
    Public Tel As String
    Public Mobile As String
    Public Email As String
    ''' <summary>Organisation the patient belongs to (0 = none).</summary>
    Public Organisation_ID As Integer
End Structure

''' <summary>A school, practice or other organisation (table Organisations).</summary>
Public Structure OrganisationStructure
    Public ID As Integer
    Public Name As String
    Public Address1 As String
    Public Address2 As String
    Public Town As String
    Public County As String
    Public Postcode As String
    Public Tel As String
End Structure

''' <summary>The person running a test (table Screeners).</summary>
Public Structure ScreenerStructure
    Public ID As Integer
    Public Title As String
    Public Firstname As String
    Public Surname As String

    Public ReadOnly Property FullName As String
        Get
            Return String.Join(" ", {Title, Firstname, Surname}.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))
        End Get
    End Property
End Structure

''' <summary>Baseline (white) or a colour test with one of the three ways of showing colour.</summary>
Public Enum ReadEZTestMode
    Baseline = 0
    Overlay = 1
    Computer = 2
    Lenses = 3
End Enum

''' <summary>
''' One test (table Tests): a baseline (white background) or a colour test.
''' Symptoms are graded 0 (none) to 5 (severe). Rate of reading = words per minute.
''' </summary>
Public NotInheritable Class ReadEZResult
    ''' <summary>Row id (0 = not saved yet).</summary>
    Public ID As Integer
    Public IsBaseline As Boolean
    Public Mode As ReadEZTestMode
    Public Client_ID As Integer
    Public Screener_ID As Integer
    Public Test_Date As Date

    ' ---- Colour (colour tests only) ----
    ''' <summary>Preferred colour: 1 - 12 = A - L, 0 = grey / none.</summary>
    Public Colour1 As Integer
    ''' <summary>1 overlays, 2 lenses, 3 strong (dark) lenses - as ReadEZ 2020.</summary>
    Public PaletteSet As Integer = 1
    Public Colour_R, Colour_G, Colour_B As Integer
    ''' <summary>Grey was preferred to the chosen colour (automatic test).</summary>
    Public GreyPreferred As Boolean
    ''' <summary>0 = automatic, 1 = manual, 2 = bespoke colour.</summary>
    Public Screening_Mode As Integer
    ''' <summary>Automatic test: how often each colour was chosen ("A:5,B:0,..."), for the colour wheel.</summary>
    Public AutoScore As String = ""

    ' ---- Symptoms ----
    Public Symptoms As Integer() = New Integer(7) {}
    Public Symptoms_Notes As String = ""
    ''' <summary>True once the symptoms questions were answered.</summary>
    Public SymptomsDone As Boolean

    ' ---- Rate of reading ----
    Public ROR As Integer
    Public ROR_RecordTime As Integer
    Public ROR_NoErrors As Integer
    Public ROR_RowsSkipped As Integer
    Public ROR_NoWordsRead As Integer
    Public ROR_Transcript As String = ""
    ''' <summary>Name of the text read.</summary>
    Public ROR_Text As String = ""
    ''' <summary>The clinician's notes on this test.</summary>
    Public Notes As String = ""
    ''' <summary>"ReadEZ 2020" for a test imported from the previous version ("" otherwise).</summary>
    Public ImportedFrom As String = ""

    ' ---- How the text was shown ----
    Public Test_Left, Test_Top, Test_Width, Test_Height As Integer
    Public RoR_LineSpacing As Integer
    Public RoR_FontSize As Integer
    Public RoR_FontName As String = "Times New Roman"
    Public RoR_NoRows As Integer = 20
    Public RoR_TextOptions As Integer
    Public RoR_OverlaySize As Integer
    Public RoR_PresentationTime As Integer

    ''' <summary>The ReadEZ set of the preferred colour.</summary>
    Public ReadOnly Property ColourSet As ReadEZSet
        Get
            Select Case PaletteSet
                Case 2 : Return ReadEZSet.Lens
                Case 3 : Return ReadEZSet.LensStrong
                Case Else : Return ReadEZSet.Overlay
            End Select
        End Get
    End Property

    ''' <summary>The colour as shown on the screen during the test.</summary>
    Public ReadOnly Property ScreenColour As Color
        Get
            If IsBaseline Then Return Color.White
            If Colour_R + Colour_G + Colour_B > 0 Then Return Color.FromArgb(255, Colour_R, Colour_G, Colour_B)
            Return ReadEZColours.Colour(ColourSet, Math.Max(0, Math.Min(12, Colour1)))
        End Get
    End Property

    ''' <summary>"Overlay F", "Lens C (dark)", "Bespoke colour" or "White".</summary>
    Public ReadOnly Property ColourName As String
        Get
            If IsBaseline Then Return Loc.T("result.white")
            If Screening_Mode = 2 Then Return Loc.T("result.bespoke")
            Return ReadEZColours.Name(ColourSet, Math.Max(0, Math.Min(12, Colour1)))
        End Get
    End Property

    ''' <summary>The sum of the symptom grades (0 - 20 for the four questions).</summary>
    Public ReadOnly Property SymptomTotal As Integer
        Get
            Return Symptoms.Take(4).Sum()
        End Get
    End Property

    Public ReadOnly Property HasReading As Boolean
        Get
            Return ROR_RecordTime > 0 AndAlso ROR_NoWordsRead > 0
        End Get
    End Property

    ''' <summary>Words per minute: 60 x (words read - 15 x rows skipped) / time (as ReadEZ 2020).</summary>
    Public Shared Function WordsPerMinute(words As Integer, rowsSkipped As Integer, seconds As Integer) As Integer
        If seconds <= 0 Then Return 0
        Return Math.Max(0, CInt(Math.Floor(60.0 * (words - 15 * rowsSkipped) / seconds)))
    End Function

    Public Sub UpdateRate()
        ROR = WordsPerMinute(ROR_NoWordsRead, ROR_RowsSkipped, ROR_RecordTime)
    End Sub

    Public Function Clone() As ReadEZResult
        Dim c = DirectCast(MemberwiseClone(), ReadEZResult)
        c.Symptoms = CType(Symptoms.Clone(), Integer())
        Return c
    End Function
End Class

