Option Strict On

' ====================================================================================================
'   SessionState.vb
'   What the clinician is working on right now: the selected patient, the selected result and the
'   screener running the tests.
'
'   Screens listen to the events here instead of reaching into each other's controls (the previous
'   version shared everything through global variables).
' ====================================================================================================

Public NotInheritable Class SessionState

    Private _patient As PatientStructure?
    Private _resultId As Integer = -1

    ''' <summary>Raised when a different patient is selected (or the selection is cleared).</summary>
    Public Event PatientChanged As EventHandler
    ''' <summary>Raised when a different result is selected.</summary>
    Public Event ResultChanged As EventHandler
    ''' <summary>Raised when a different screener is chosen.</summary>
    Public Event ScreenerChanged As EventHandler
    ''' <summary>Raised after a result has been saved (so lists can refresh).</summary>
    Public Event ResultSaved(result As ReadEZResult)

#Region "Patient"

    Public ReadOnly Property HasPatient As Boolean
        Get
            Return _patient.HasValue
        End Get
    End Property

    Public ReadOnly Property Patient As PatientStructure
        Get
            Return _patient.GetValueOrDefault()
        End Get
    End Property

    Public ReadOnly Property PatientId As Integer
        Get
            Return If(_patient.HasValue, _patient.Value.ID, -1)
        End Get
    End Property

    Public ReadOnly Property PatientName As String
        Get
            If Not _patient.HasValue Then Return ""
            Dim p = _patient.Value
            Return $"{p.Title} {p.Firstname} {p.Surname}".Trim().Replace("  ", " ")
        End Get
    End Property

    Public Sub SelectPatient(p As PatientStructure?)
        Dim changed = PatientId <> If(p.HasValue, p.Value.ID, -1)
        _patient = p
        If changed Then
            _resultId = -1
            RaiseEvent PatientChanged(Me, EventArgs.Empty)
            RaiseEvent ResultChanged(Me, EventArgs.Empty)
        Else
            RaiseEvent PatientChanged(Me, EventArgs.Empty)
        End If
        If p.HasValue Then AppHost.Settings.LastPatientId = p.Value.ID
    End Sub

    ''' <summary>Re-reads the current patient from the database (after editing).</summary>
    Public Sub ReloadPatient()
        If Not _patient.HasValue Then Return
        _patient = AppHost.Db.GetPatientById(_patient.Value.ID)
        RaiseEvent PatientChanged(Me, EventArgs.Empty)
    End Sub

#End Region

#Region "Result"

    ''' <summary>ID of the selected test (-1 = none).</summary>
    Public ReadOnly Property ResultId As Integer
        Get
            Return _resultId
        End Get
    End Property

    Public ReadOnly Property HasResult As Boolean
        Get
            Return _resultId > 0
        End Get
    End Property

    Public Sub SelectResult(id As Integer)
        If id = _resultId Then Return
        _resultId = id
        RaiseEvent ResultChanged(Me, EventArgs.Empty)
    End Sub

    Public Sub NotifyResultSaved(r As ReadEZResult)
        RaiseEvent ResultSaved(r)
    End Sub

#End Region

#Region "Screener"

    ''' <summary>The person running the tests (table Tester; 0 = not chosen).</summary>
    Public ReadOnly Property ScreenerId As Integer
        Get
            Return AppHost.Settings.ScreenerID
        End Get
    End Property

    Public ReadOnly Property ScreenerName As String
        Get
            Return AppHost.Db.ScreenerName(ScreenerId)
        End Get
    End Property

    Public Sub SelectScreener(id As Integer)
        AppHost.Settings.ScreenerID = id
        AppHost.SaveSettings()
        RaiseEvent ScreenerChanged(Me, EventArgs.Empty)
    End Sub

#End Region

End Class
