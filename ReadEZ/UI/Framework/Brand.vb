Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   Brand.vb - Thomson Software Solutions colours and the "key result" panel
'
'   The colours are taken from the Thomson Software Solutions logo (Assets\ThomsonLogo.png, on
'   white): blue #2969B3, light blue #4F8CCB, charcoal #414143 and grey #707173. They are used for
'   the accents of the analysis windows (card titles) and the printed reports (header rule,
'   section headings, table headers), with green / amber / red for results that are good, worth a
'   look or a concern.
'
'   DrawKeyResult draws the main result of a test as a tinted panel with a coloured edge: a small
'   label in capitals, the result large, and a line of explanation - on screen and in reports.
' ====================================================================================================

''' <summary>A screen or window whose cards all take the brand accent (Card.Accent) without each being set.</summary>
Public Interface IBrandedHost
    ReadOnly Property Branded As Boolean
End Interface

Public NotInheritable Class Brand
    Private Sub New()
    End Sub

    Public Shared ReadOnly Blue As Color = Color.FromArgb(41, 105, 179)          ' #2969B3
    Public Shared ReadOnly LightBlue As Color = Color.FromArgb(79, 140, 203)     ' #4F8CCB
    Public Shared ReadOnly Charcoal As Color = Color.FromArgb(65, 65, 67)        ' #414143
    Public Shared ReadOnly Grey As Color = Color.FromArgb(112, 113, 115)         ' #707173
    ''' <summary>Dark brand blue for the report header band.</summary>
    Public Shared ReadOnly Navy As Color = Color.FromArgb(27, 72, 132)           ' #1B4884
    ''' <summary>Very light blue for shaded headers and panels, and its border.</summary>
    Public Shared ReadOnly Tint As Color = Color.FromArgb(235, 242, 250)
    Public Shared ReadOnly TintBorder As Color = Color.FromArgb(196, 214, 236)

    ''' <summary>How a result reads: good (green), worth a look (amber), a concern (red), or just information (blue).</summary>
    Public Enum Level
        Info
        Good
        Caution
        Concern
    End Enum

    Public Shared Function LevelColour(l As Level) As Color
        Select Case l
            Case Level.Good : Return Color.FromArgb(46, 140, 70)
            Case Level.Caution : Return Color.FromArgb(222, 136, 14)
            Case Level.Concern : Return Color.FromArgb(198, 48, 44)
            Case Else : Return Blue
        End Select
    End Function

    ''' <summary>The fill of a level's panel: white throughout (the level shows in the coloured edge and label).</summary>
    Public Shared Function LevelTint(l As Level) As Color
        Return Color.White
    End Function

    ''' <summary>The accent for titles on screen: brand blue, lighter on a dark theme.</summary>
    Public Shared Function Accent() As Color
        Return If(IsDark(), Color.FromArgb(128, 176, 230), Blue)
    End Function

    Public Shared Function IsDark() As Boolean
        Return ThemeManager.Current.Surface.GetBrightness() < 0.4F
    End Function

    ''' <summary>
    ''' The key result panel: label (small capitals), value (large), note (wraps). surface = the colour
    ''' behind it (the tint is mixed with it on dark themes). Returns the height used.
    ''' </summary>
    Public Shared Function DrawKeyResult(g As Graphics, r As RectangleF, label As String, value As String, note As String,
                                        lvl As Level, ink As Color, muted As Color, surface As Color,
                                        Optional valueSize As Single = 20.0F) As Single
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit
        Dim edge = LevelColour(lvl)
        Dim dark = surface.GetBrightness() < 0.4F
        Dim fill = If(dark, Mix(surface, edge, 0.12F), LevelTint(lvl))
        Using lf As New Font("Segoe UI Semibold", 8.75F), vf As New Font("Segoe UI Semibold", valueSize), nf As New Font("Segoe UI", 10.0F)
            Dim pad = 14.0F
            Dim textW = r.Width - 2 * pad - 6
            Dim lh = g.MeasureString(label.ToUpperInvariant(), lf, CInt(textW)).Height
            Dim vh = g.MeasureString(value, vf, CInt(textW)).Height
            Dim nh = If(String.IsNullOrEmpty(note), 0, g.MeasureString(note, nf, CInt(textW)).Height)
            Dim h = Math.Max(r.Height, pad + lh + 2 + vh + If(nh > 0, 4 + nh, 0) + pad - 4)
            Dim box As New RectangleF(r.X, r.Y, r.Width, h)
            Using path = Global.ReadEZ.Draw.RoundedRect(box, 8), fb As New SolidBrush(fill), pen As New Pen(Color.FromArgb(150, edge), 1),
                  bar As New SolidBrush(edge)
                g.FillPath(fb, path)
                Dim state = g.Save()
                g.SetClip(path, CombineMode.Intersect)
                g.FillRectangle(bar, box.X, box.Y, 6, box.Height)
                g.Restore(state)
                g.DrawPath(pen, path)
            End Using
            Using lb As New SolidBrush(If(dark, Mix(edge, Color.White, 0.35F), edge)), vb As New SolidBrush(ink), nb As New SolidBrush(muted)
                Dim x = box.X + pad + 6, y = box.Y + pad - 2
                g.DrawString(label.ToUpperInvariant(), lf, lb, New RectangleF(x, y, textW, lh + 2))
                y += lh + 2
                g.DrawString(value, vf, vb, New RectangleF(x, y, textW, vh + 2))
                y += vh + 4
                If nh > 0 Then g.DrawString(note, nf, nb, New RectangleF(x, y, textW, nh + 2))
            End Using
            Return h
        End Using
    End Function

    ''' <summary>a, moved the fraction t towards b.</summary>
    Public Shared Function Mix(a As Color, b As Color, t As Single) As Color
        Return Color.FromArgb(CInt(a.R + (CInt(b.R) - a.R) * t), CInt(a.G + (CInt(b.G) - a.G) * t), CInt(a.B + (CInt(b.B) - a.B) * t))
    End Function
End Class
