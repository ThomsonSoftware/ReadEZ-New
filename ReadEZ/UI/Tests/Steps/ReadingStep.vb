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
    Private ReadOnly nudTime As NumericUpDown = Ui.NumberBox(0, 600, 0, width:=110)
    Private ReadOnly nudWords As NumericUpDown = Ui.NumberBox(0, 2000, 0, width:=110)
    Private ReadOnly nudErrors As NumericUpDown = Ui.NumberBox(0, 500, 0, width:=110)
    Private ReadOnly nudSkipped As NumericUpDown = Ui.NumberBox(0, 20, 0, width:=110)
    Private ReadOnly btnReset As ModernButton = ModernButton.Create("reading.reset", ButtonKind.Ghost, Glyphs.Refresh)
    Private ReadOnly btn60 As ModernButton = ModernButton.Create("reading.sixty", ButtonKind.Secondary)
    Private ReadOnly segView As New SegmentedControl
    Private _syncing As Boolean
    ''' <summary>Two screens: show the enlarged scoring copy (remembered while the program runs).</summary>
    Private Shared _scoringView As Boolean = True

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
        Ui.FillWidth(col)                                 ' labels wrap at, and cards stretch to, the panel's width
        Dim intro = Ui.Label("reading.intro", muted:=True)
        intro.Font = Fonts.Body
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)

        ' Two screens: the centre can show the text enlarged for scoring, or an exact copy of the patient's screen
        If ScreenLayout.IsTwoScreenMode() Then
            segView.SetOptions("reading.view.scoring", "reading.view.patient")
            segView.SelectedIndex = If(_scoringView, 0, 1)
            segView.Margin = New Padding(0, 0, 0, 10)
            AddHandler segView.SelectedIndexChanged, Sub()
                                                         _scoringView = segView.SelectedIndex = 0
                                                         If Host IsNot Nothing Then Host.ScoringView = _scoringView
                                                     End Sub
            col.Controls.Add(segView)
        End If

        ' Clock, with Reset (back to 0:00) and 60 s (sets the reading time to a full minute)
        Dim clockRow = Ui.Row()
        clockRow.Margin = New Padding(0, 0, 0, 4)
        lblClock.Margin = New Padding(0, 0, 12, 0)
        clockRow.Controls.Add(lblClock)
        btnReset.Margin = New Padding(0, 10, 6, 0)
        btn60.Margin = New Padding(0, 10, 0, 0)
        AddHandler btnReset.Click, Sub() ResetClock()
        AddHandler btn60.Click, Sub() SetFullMinute()
        clockRow.Controls.Add(btnReset)
        clockRow.Controls.Add(btn60)
        col.Controls.Add(clockRow)
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

        ' How to score: in a window, so the panel stays short
        Dim actions = Ui.Row(wrap:=True)
        actions.Margin = New Padding(0, 8, 0, 0)
        Dim how = HelpButton("reading.how.button", "reading.how.title",
                             "reading.how.1", "reading.how.2", "reading.how.3", "reading.how.4", "reading.how.5", "reading.how.6", "reading.how.7")
        how.Margin = New Padding(0, 0, 8, 0)
        actions.Controls.Add(how)
        col.Controls.Add(actions)
        Dim clearMarks = ModernButton.Create("reading.clearMarks", ButtonKind.Ghost, Glyphs.Delete)
        AddHandler clearMarks.Click, Sub()
                                         Scene.Errors.Clear() : Scene.SkippedRows.Clear() : Scene.LastWord = Nothing
                                         _syncing = True
                                         nudErrors.Value = 0 : nudSkipped.Value = 0 : nudWords.Value = 0
                                         _syncing = False
                                         ValuesChanged()
                                     End Sub
        actions.Controls.Add(clearMarks)

        col.Controls.Add(TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=Run.Result.Mode = ReadEZTestMode.Overlay,
                                            enabledWhen:=Function() Not _running AndAlso Scene.Errors.Count = 0 AndAlso Not Scene.LastWord.HasValue))

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
        AddHandler nudTime.ValueChanged, Sub() If Not _syncing AndAlso Not _running Then UpdateClock(CInt(nudTime.Value))
        AddHandler _clock.Tick, Sub() ClockTick()
        UpdateClock(r.ROR_RecordTime)
        ValuesChanged()
        Return col
    End Function

    Public Overrides Function InstructionKeys() As String()
        Return {"instr.reading.title", "instr.reading.1", "instr.reading.2", "instr.reading.3"}
    End Function

    ''' <summary>The text as it will be shown: start at the first word, read along each line.</summary>
    Public Overrides Function InstructionScreen() As InstructionScene
        Dim sample = SampleText()
        Dim r = Run.Result
        If Not r.IsBaseline AndAlso Run.HasColour Then
            sample.Colour1 = r.ScreenColour
            sample.Presentation = If(r.Mode = ReadEZTestMode.Overlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen)
        End If
        Dim sc As New InstructionScene(InstructionKeys()) With {.Sample = sample}
        sc.Add("instr.reading.callout1", sample.WordBounds(0, 0), InstructionScene.Side.Left)
        sc.Add("instr.reading.callout2", sample.RowBounds(1), InstructionScene.Side.Right)
        sc.Add("instr.reading.callout3", RectangleF.Empty, InstructionScene.Side.BottomCentre)
        Return sc
    End Function

    Public Overrides Function CreateFooter() As Control
        Return FooterRow(NextButton())
    End Function

    Private Function IsLastStep() As Boolean
        Return Run.Index = Run.Steps.Count - 1
    End Function

    Public Overrides Sub Begin()
        ApplyTextSettings()
        ApplyTestColour()
        Scene.ShowMarks = True
        If Host IsNot Nothing Then Host.ScoringView = _scoringView
        MyBase.Begin()
    End Sub

    Public Overrides Sub [End]()
        _clock.Stop()
        _running = False
        If Host IsNot Nothing Then Host.ScoringView = False
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
        ShowClockStopped()
        ValuesChanged()
        ActionsChanged()
        If timeUp AndAlso AppHost.Settings.RoR_Beep Then NativeSound.Beep()
        Toast.Show(AppHost.MainWindow, Loc.T(If(timeUp, "reading.timeUp", "reading.stopped")), ToastKind.Info, 6000)
    End Sub

    Private Sub ShowClockStopped()
        btnClock.Text = Loc.T("reading.start")
        btnClock.IconGlyph = Glyphs.Play
        btnClock.Kind = ButtonKind.Primary
    End Sub

    ''' <summary>Sets the time to the value given (stopping the clock quietly if it is running).</summary>
    Private Sub SetTime(seconds As Integer)
        If _running Then
            _clock.Stop()
            _running = False
            ShowClockStopped()
        End If
        _syncing = True
        nudTime.Value = Math.Max(nudTime.Minimum, Math.Min(nudTime.Maximum, seconds))
        _syncing = False
        UpdateClock(CInt(nudTime.Value))
        ValuesChanged()
        ActionsChanged()
    End Sub

    ''' <summary>Reset: the clock back to 0:00 (the marks stay; Clear marks removes them).</summary>
    Private Sub ResetClock()
        SetTime(0)
    End Sub

    ''' <summary>60 s: the reading time set to a full minute (e.g. the patient read for the whole minute).</summary>
    Private Sub SetFullMinute()
        SetTime(60)
    End Sub

    Private Sub UpdateClock(seconds As Integer)
        lblClock.Text = $"{seconds \ 60}:{seconds Mod 60:00}"
        Host?.UpdateBarValue("time", lblClock.Text)     ' one screen: the time on the bar
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
        MarkWord(bounds, e)
    End Sub

    ''' <summary>
    ''' One screen: the examiner can mark on the patient's screen too (same clicks as on the test
    ''' screen; the patient doesn't see the marks - the counts show in the bar at the bottom).
    ''' Not with a tablet or a second monitor: clicks there are the patient's.
    ''' </summary>
    Public Overrides Sub PatientScreenClick(bounds As Rectangle, e As MouseEventArgs)
        If ScreenLayout.IsTwoScreenMode() Then Return
        MarkWord(bounds, e)
        ActionsChanged()
    End Sub

    ''' <summary>click = misread word, Shift + click = skipped row, right-click = last word read.</summary>
    Private Sub MarkWord(bounds As Rectangle, e As MouseEventArgs)
        Dim w = Scene.WordAt(bounds, e.Location)
        If Not w.HasValue Then Return
        If e.Button = MouseButtons.Right Then
            Scene.LastWord = w
            _syncing = True
            nudWords.Value = Math.Min(nudWords.Maximum, Scene.WordsReadTo(w.Value))
            _syncing = False
            ' Acknowledged: a beep and a message (on the bar over the patient's view in one-screen mode)
            System.Media.SystemSounds.Asterisk.Play()
            Host?.ShowExaminerMessage(Loc.F("reading.lastWordMarked", CInt(nudWords.Value)))
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
        ' Space: starts the clock (before it has run), then records an error; S: a skipped row; C: start / stop the clock
        If keyData = Keys.Space Then
            If Not _running AndAlso nudTime.Value = 0 Then ToggleClock() Else Bump(nudErrors, 1)
            Return True
        End If
        If keyData = Keys.S Then
            Bump(nudSkipped, 1)
            Return True
        End If
        If keyData = Keys.C Then
            ToggleClock()
            Return True
        End If
        If keyData = Keys.Enter AndAlso Not _running Then
            Finish()
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' One screen, patient full screen: the bar at the bottom - the clock (Start/Stop, Reset, 60 s),
    ''' errors and skipped rows (− / +), the words read, and Next step.
    ''' </summary>
    Public Overrides Function QuickActions() As List(Of QuickAction)
        Return New List(Of QuickAction) From {
            New QuickAction("", Nothing) With {.Text = Loc.T("reading.strip.time"), .Value = lblClock.Text, .Id = "time", .GroupStart = True},
            New QuickAction("", Sub() ToggleClock(), primary:=True) With {.Text = Loc.T(If(_running, "reading.stop", "reading.start")) & "  (C)"},
            New QuickAction("reading.reset", Sub() ResetClock()),
            New QuickAction("reading.sixty", Sub() SetFullMinute()),
            New QuickAction("", Nothing) With {.Text = Loc.T("reading.errors"), .Value = CInt(nudErrors.Value).ToString(), .GroupStart = True},
            New QuickAction("", Sub() Bump(nudErrors, -1)) With {.Text = "−"},
            New QuickAction("", Sub() Bump(nudErrors, 1)) With {.Text = "+  (" & Loc.T("key.space") & ")"},
            New QuickAction("", Nothing) With {.Text = Loc.T("reading.skipped"), .Value = CInt(nudSkipped.Value).ToString(), .GroupStart = True},
            New QuickAction("", Sub() Bump(nudSkipped, -1)) With {.Text = "−"},
            New QuickAction("", Sub() Bump(nudSkipped, 1)) With {.Text = "+  (S)"},
            New QuickAction("", Nothing) With {.Text = Loc.T("reading.strip.wordsHint"), .GroupStart = True},
            New QuickAction("", Nothing) With {.Value = CInt(nudWords.Value).ToString(), .NumberMax = CInt(nudWords.Maximum),
                                               .NumberChanged = Sub(n) SetWordsRead(n)},
            New QuickAction(If(IsLastStep(), "run.saveFinish", "run.nextStep"), Sub() Finish(), enabled:=Not _running) With {.GroupStart = True}}
    End Function

    Public Overrides Function CreateTextOptions() As Control
        Return TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=Run.Result.Mode = ReadEZTestMode.Overlay,
                                  enabledWhen:=Function() Not _running AndAlso Scene.Errors.Count = 0 AndAlso Not Scene.LastWord.HasValue)
    End Function

    ''' <summary>The bar's words read, typed or stepped with the arrows (the bar is not rebuilt, so typing carries on).</summary>
    Private Sub SetWordsRead(n As Integer)
        If nudWords.Value = n Then Return
        _syncing = True
        nudWords.Value = Math.Max(nudWords.Minimum, Math.Min(nudWords.Maximum, n))
        _syncing = False
        ValuesChanged()
    End Sub

    ''' <summary>The bar's − / +: one error or skipped row more or fewer.</summary>
    Private Sub Bump(n As NumericUpDown, delta As Integer)
        n.Value = Math.Max(n.Minimum, Math.Min(n.Maximum, n.Value + delta))
        ValuesChanged()
        ActionsChanged()
    End Sub

    Public Overrides Sub Dispose()
        _clock.Dispose()
        MyBase.Dispose()
    End Sub
End Class
