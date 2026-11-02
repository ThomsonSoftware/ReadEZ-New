Option Strict On

' ====================================================================================================
'   Dpi.vb - one rule for display scaling
'
'   All sizes in this program are written in 96-dpi ("100 %") pixels. Each window and each screen
'   is scaled ONCE, after it has been built, with Dpi.ScaleOnce. Custom controls draw with
'   LogicalToDeviceUnits, which already uses the real DPI.
'
'   Anything created after its parent has been scaled (for example tiles built when a screen is
'   opened) is scaled with Dpi.ScaleNew.
'
'   Forms use AutoScaleMode.None so WinForms does not scale a second time.
' ====================================================================================================

Public NotInheritable Class Dpi
    Private Sub New()
    End Sub

    Private Shared ReadOnly _scaled As New Runtime.CompilerServices.ConditionalWeakTable(Of Control, Object)

    ''' <summary>Scale factor of the screen the control is on (1.0 at 100 %).</summary>
    Public Shared Function Factor(c As Control) As Single
        Return c.DeviceDpi / 96.0F
    End Function

    ''' <summary>True once ScaleOnce has been applied to the control (or one of its parents).</summary>
    Public Shared Function IsScaled(c As Control) As Boolean
        Dim x = c
        While x IsNot Nothing
            Dim o As Object = Nothing
            If _scaled.TryGetValue(x, o) Then Return True
            x = x.Parent
        End While
        Return False
    End Function

    ''' <summary>Scales a window or screen from 96-dpi sizes to the real DPI. Does nothing the second time.</summary>
    Public Shared Sub ScaleOnce(c As Control)
        Dim o As Object = Nothing
        If _scaled.TryGetValue(c, o) Then Return
        _scaled.Add(c, True)
        Dim f = Factor(c)
        If Math.Abs(f - 1.0F) > 0.01F Then c.Scale(New SizeF(f, f))
    End Sub

    ''' <summary>Scales a control added to an already-scaled parent.</summary>
    Public Shared Sub ScaleNew(child As Control, parent As Control)
        If Not IsScaled(parent) Then Return
        Dim f = Factor(parent)
        If Math.Abs(f - 1.0F) > 0.01F Then child.Scale(New SizeF(f, f))
    End Sub

    ''' <summary>Converts a 96-dpi length for a control that may or may not have been scaled yet.</summary>
    Public Shared Function Px(c As Control, logical As Integer) As Integer
        Return If(IsScaled(c), CInt(Math.Round(logical * Factor(c))), logical)
    End Function
End Class
