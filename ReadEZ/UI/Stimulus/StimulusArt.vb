Option Strict On

Imports System.Drawing.Drawing2D
Imports System.Drawing.Text

' ====================================================================================================
'   StimulusArt.vb - still pictures for the patient's screen
'
'   Pictures are drawn at the size of the patient's screen (Settings > Screens) and shown with
'   Show(). The reading text itself is a scene (ReadingScene), so it is drawn sharp at any size.
' ====================================================================================================

Public NotInheritable Class StimulusArt
    Private Sub New()
    End Sub

    ''' <summary>Size of the patient's screen in pixels, from Settings > Screens.</summary>
    Public Shared ReadOnly Property ScreenSize As Size
        Get
            If TabletScreen.IsActive Then Return TabletScreen.ScreenSize     ' drawn at the tablet's size
            Dim s = AppHost.Settings
            Return New Size(Math.Max(1, s.Stimulus_Width), Math.Max(1, s.Stimulus_Height))
        End Get
    End Property

    ''' <summary>Pixels per millimetre on the patient's screen (Settings > Screens > screen width).</summary>
    Public Shared ReadOnly Property PxPerMm As Single
        Get
            Dim s = AppHost.Settings
            Return If(s.Screen_Width > 0 AndAlso s.Stimulus_Width > 0, s.Stimulus_Width / CSng(s.Screen_Width), 96 / 25.4F)
        End Get
    End Property

    ''' <summary>Draws a full-screen picture and shows it on the patient's screen.</summary>
    Public Shared Sub Show(background As Color, draw As Action(Of Graphics, Size))
        Dim sz = ScreenSize
        Dim bmp As New Bitmap(sz.Width, sz.Height)
        Using g = Graphics.FromImage(bmp)
            g.Clear(background)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
            draw?.Invoke(g, sz)
        End Using
        Dim w = Stimulus.Window
        w.Scene = Nothing
        w.BackColor = background
        w.SetPicture(bmp, owns:=True)
    End Sub

    ''' <summary>Plain screen of one colour.</summary>
    Public Shared Sub Blank(background As Color)
        Show(background, Nothing)
    End Sub
End Class
