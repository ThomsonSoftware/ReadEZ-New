Option Strict On

' ====================================================================================================
'   ReadingStep.vb - rate of reading on white (baseline) or with the chosen colour
'
'   The Wilkins Rate of Reading test: the patient reads the words aloud as quickly and accurately
'   as they can. As ReadEZ 2020:
'     - Start clock (Space) starts the timer; it stops by itself after the reading time (60 s by
'       default, with a beep) or with Stop clock (Space).
'     - The examiner marks the transcript on the copy of the text in this window (the patient
'       doesn't see the marks):
'         click             a misread word (click again to remove)
'         Shift + click     a skipped row
'         right-click       the last word read - the words read are counted from it
'     - Words per minute = 60 x (words read - 15 x rows skipped) / seconds.
'   The counts can also be typed in (for example from a paper copy marked during the test).
' ====================================================================================================

Friend NotInheritable Class NativeSound
    Private Sub New()
    End Sub

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function MessageBeep(type As UInteger) As Boolean
    End Function

    ''' <summary>The Windows "exclamation" sound.</summary>
    Public Shared Sub Beep()
        Try
            MessageBeep(&H30)
        Catch
        End Try
    End Sub
End Class

Public NotInheritable Class ReadingStep
    Inherits TestStep

    Private ReadOnly _clock As New Timer With {.Interval = 100}
    Private _started As Date
    Private _running As Boolean
    Private ReadOnly lblClock As New Label With {.AutoSize = True, .Font = Fonts.Display}
    Private ReadOnly lblWpm As New Label With {.AutoSize = True, .Font = Fonts.Title}
    Private ReadOnly btnClock As ModernButton = ModernButton.Create("reading.start", ButtonKind.Primary, Glyphs.Play)
    Private ReadOnly nudTime As NumericUpDown = Ui.NumberBox(0, 600, 0)
    Private ReadOnly nudWords As NumericUpDown = Ui.NumberBox(0, 2000, 0)
    Private ReadOnly nudErrors As NumericUpDown = Ui.NumberBox(0, 500, 0)
    Private ReadOnly nudSkipped As NumericUpDown = Ui.NumberBox(0, 20, 0)
    Private ReadOnly lblHelp As New Label With {.AutoSize = True, .Font = Fonts.Small, .MaximumSize = New Size(310, 0)}
    Private _syncing As Boolean

    Public Overrides ReadOnly Property TitleKey As String
        Get
            Return If(Run IsNot Nothing AndAlso Not Run.Result.IsBaseline, "step.readingColour", "step.reading")
        End Get
    End Property

    Public Overrides ReadOnly Property HelpTopic As String
        Get
            Return "test.reading"
        End Get
    End Property

    Public Overrides Function CreatePanel() As Control
        Dim col = Ui.Column()
        Dim intro = Ui.Label("reading.intro", muted:=True)
        intro.MaximumSize = New Size(310, 0)
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)

        lblClock.Margin = New Padding(0, 0, 0, 4)
        col.Controls.Add(lblClock)
        btnClock.Margin = New Padding(0, 0, 0, 12)
        AddHandler btnClock.Click, Sub() ToggleClock()
        col.Controls.Add(btnClock)

        Dim grid As New TableLayoutPanel With {.ColumnCount = 2, .AutoSize = True, .BackColor = Color.Transparent, .Margin = New Padding(0)}
        grid.Controls.Add(Ui.Stack("reading.seconds", nudTime), 0, 0)
        grid.Controls.Add(Ui.Stack("reading.words", nudWords), 1, 0)
        grid.Controls.Add(Ui.Stack("reading.errors", nudErrors), 0, 1)
        grid.Controls.Add(Ui.Stack("reading.skipped", nudSkipped), 1, 1)
        col.Controls.Add(grid)
        Dim wpmRow = Ui.Row()
        wpmRow.Controls.Add(lblWpm)
        Dim unit = Ui.Label("reading.wpmUnit", muted:=True)
        unit.Margin = New Padding(4, 12, 0, 0)
        wpmRow.Controls.Add(unit)
        col.Controls.Add(wpmRow)

        LabelStyle.SetMuted(lblHelp)
        Loc.Bind(lblHelp, "reading.marking")
        lblHelp.Margin = New Padding(0, 8, 0, 4)
        col.Controls.Add(lblHelp)
        Dim clearMarks = ModernButton.Create("reading.clearMarks", ButtonKind.Ghost, Glyphs.Delete)
        AddHandler clearMarks.Click, Sub()
                                         Scene.Errors.Clear() : Scene.SkippedRows.Clear() : Scene.LastWord = Nothing
                                         _syncing = True
                                         nudErrors.Value = 0 : nudSkipped.Value = 0 : nudWords.Value = 0
                                         _syncing = False
                                         ValuesChanged()
                                     End Sub
        col.Controls.Add(clearMarks)

        col.Controls.Add(TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=Run.Result.Mode = ReadEZTestMode.Overlay,
                                            enabledWhen:=Function() Not _running AndAlso Scene.Errors.Count = 0 AndAlso Not Scene.LastWord.HasValue))
        Dim nextBtn = ModernButton.Create(If(IsLastStep(), "run.saveFinish", "run.nextStep"), ButtonKind.Primary, If(IsLastStep(), Glyphs.Save, Glyphs.Forward))
        nextBtn.Margin = New Padding(0, 14, 0, 0)
        AddHandler nextBtn.Click, Sub() Finish()
        col.Controls.Add(nextBtn)

        Dim r = Run.Result
        _syncing = True
        nudTime.Value = Math.Min(nudTime.Maximum, r.ROR_RecordTime)
        nudWords.Value = Math.Min(nudWords.Maximum, r.ROR_NoWordsRead)
        nudErrors.Value = Math.Min(nudErrors.Maximum, r.ROR_NoErrors)
        nudSkipped.Value = Math.Min(nudSkipped.Maximum, r.ROR_RowsSkipped)
        _syncing = False
        For Each n In {nudTime, nudWords, nudErrors, nudSkipped}
            AddHandler n.ValueChanged, Sub() If Not _syncing Then ValuesChanged()
        Next
        AddHandler _clock.Tick, Sub() ClockTick()
        UpdateClock(r.ROR_RecordTime)
        ValuesChanged()
        Return col
    End Function

    Private Function IsLastStep() As Boolean
        Return Run.Index = Run.Steps.Count - 1
    End Function

    Public Overrides Sub Begin()
        ApplyTextSettings()
        ApplyTestColour()
        Scene.ShowMarks = True
        MyBase.Begin()
    End Sub

    Public Overrides Sub [End]()
        _clock.Stop()
        _running = False
    End Sub

    Public Overrides Function IsComplete() As Boolean
        Return Run.Result.HasReading
    End Function

#Region "Clock"

    Private Sub ToggleClock()
        If _running Then
            StopClock(timeUp:=False)
        Else
            _started = Date.Now
            _running = True
            _clock.Start()
            btnClock.Text = Loc.T("reading.stop")
            btnClock.IconGlyph = Glyphs.Stop
            btnClock.Kind = ButtonKind.Danger
            ActionsChanged()
        End If
    End Sub

    Private Sub ClockTick()
        Dim secs = (Date.Now - _started).TotalSeconds
        UpdateClock(CInt(Math.Floor(secs)))
        If secs >= AppHost.Settings.RoR_ReadingTime Then StopClock(timeUp:=True)
    End Sub

    Private Sub StopClock(timeUp As Boolean)
        _clock.Stop()
        _running = False
        Dim secs = CInt(Math.Min(AppHost.Settings.RoR_ReadingTime, Math.Round((Date.Now - _started).TotalSeconds)))
        _syncing = True
        nudTime.Value = Math.Min(nudTime.Maximum, secs)
        _syncing = False
        UpdateClock(secs)
        btnClock.Text = Loc.T("reading.start")
        btnClock.IconGlyph = Glyphs.Play
        btnClock.Kind = ButtonKind.Primary
        ValuesChanged()
        ActionsChanged()
        If timeUp AndAlso AppHost.Settings.RoR_Beep Then NativeSound.Beep()
        Toast.Show(AppHost.MainWindow, Loc.T(If(timeUp, "reading.timeUp", "reading.stopped")), ToastKind.Info, 6000)
    End Sub

    Private Sub UpdateClock(seconds As Integer)
        lblClock.Text = $"{seconds \ 60}:{seconds Mod 60:00}"
    End Sub

#End Region

#Region "Transcript marks on the view"

    Private _hover As WordRef?

    Public Overrides ReadOnly Property PreviewCursor As Cursor
        Get
            Return If(_hover.HasValue, Cursors.Hand, Cursors.Default)
        End Get
    End Property

    Public Overrides Sub PreviewMouseMove(bounds As Rectangle, p As Point)
        Dim w = Scene.WordAt(bounds, p)
        If Not Nullable.Equals(w, _hover) Then
            _hover = w
            Scene.HoverWord = w
            Host?.RefreshView()
        End If
    End Sub

    Public Overrides Sub PreviewMouseClick(bounds As Rectangle, e As MouseEventArgs)
        Dim w = Scene.WordAt(bounds, e.Location)
        If Not w.HasValue Then Return
        If e.Button = MouseButtons.Right Then
            Scene.LastWord = w
            _syncing = True
            nudWords.Value = Math.Min(nudWords.Maximum, Scene.WordsReadTo(w.Value))
            _syncing = False
        ElseIf (Control.ModifierKeys And Keys.Shift) = Keys.Shift Then
            If Not Scene.SkippedRows.Remove(w.Value.Row) Then Scene.SkippedRows.Add(w.Value.Row)
            _syncing = True
            nudSkipped.Value = Math.Min(nudSkipped.Maximum, Scene.SkippedRows.Count)
            _syncing = False
        Else
            If Not Scene.Errors.Remove(w.Value) Then Scene.Errors.Add(w.Value)
            _syncing = True
            nudErrors.Value = Math.Min(nudErrors.Maximum, Scene.Errors.Count)
            _syncing = False
        End If
        ValuesChanged()
        Host?.RefreshView()
    End Sub

#End Region

    Private Sub ValuesChanged()
        Dim r = Run.Result
        r.ROR_RecordTime = CInt(nudTime.Value)
        r.ROR_NoWordsRead = CInt(nudWords.Value)
        r.ROR_NoErrors = CInt(nudErrors.Value)
        r.ROR_RowsSkipped = CInt(nudSkipped.Value)
        r.UpdateRate()
        If Scene.Errors.Count > 0 OrElse Scene.SkippedRows.Count > 0 OrElse Scene.LastWord.HasValue Then r.ROR_Transcript = Scene.TranscriptText()
        lblWpm.Text = If(r.HasReading, r.ROR.ToString(), "–")
    End Sub

    Public Overrides Function HandleKey(keyData As Keys) As Boolean
        If keyData = Keys.Space Then
            ToggleClock()
            Return True
        End If
        If keyData = Keys.Enter AndAlso Not _running Then
            Finish()
            Return True
        End If
        Return False
    End Function

    Public Overrides Function QuickActions() As List(Of QuickAction)
        Return New List(Of QuickAction) From {
            New QuickAction(If(_running, "reading.stop", "reading.start"), Sub() ToggleClock(), primary:=True),
            New QuickAction(If(IsLastStep(), "run.saveFinish", "run.nextStep"), Sub() Finish(), enabled:=Not _running)}
    End Function

    Public Overrides Sub Dispose()
        _clock.Dispose()
        MyBase.Dispose()
    End Sub
End Class
