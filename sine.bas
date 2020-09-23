Attribute VB_Name = "Sine_etc"
'
' Original ideas were found at:
' Sine wave using Bezier - http://www.tinaja.com/glib/bezsine.pdf
' Drawing rotated ellipses using Bezier - http://www.codeguru.com/gdi/ellipse.shtml
'
'Copyright © 2002, Jan Tošovský
'

'Sine constants
Const PI = 3.14159265359
Const PI12 = PI / 12
Const k = 0.2020305089104 'Sqr(2) / 7
Const k1 = 1 / 7
Const k2 = 2 / 7

'Ellipse constants
Const E_factor = 0.2761423749154

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Function EllipsePts(x0, y0, Width, Height, Angle) As POINTAPI()
Dim offsetX As Single, offsetY As Single, r As Single, alfa As Single
Dim X(12) As Single, Y(12) As Single, p(12) As POINTAPI

    offsetX = Width * E_factor
    offsetY = Height * E_factor
    
    X(0) = -Width / 2
    X(1) = X(0)
    X(11) = X(0)
    X(12) = X(0)
    
    X(5) = Width / 2
    X(6) = X(5)
    X(7) = X(5)
    
    X(2) = -offsetX
    X(10) = X(2)
    
    X(4) = offsetX
    X(8) = X(4)

    X(3) = 0: X(9) = 0
    
    Y(2) = -Height / 2
    Y(3) = Y(2)
    Y(4) = Y(2)
    
    Y(8) = Height / 2
    Y(9) = Y(8)
    Y(10) = Y(8)
    
    Y(7) = offsetY
    Y(11) = Y(7)
    
    Y(1) = -offsetY
    Y(5) = Y(1)
    
    Y(0) = 0: Y(12) = 0: Y(6) = 0
    
    For i = 0 To 12
        r = Sqr(X(i) ^ 2 + Y(i) ^ 2)
        alfa = Atn2(Y(i), X(i)) + Angle * PI / 180
        p(i).X = x0 + r * Cos(alfa)
        p(i).Y = y0 - r * Sin(alfa)
    Next i

    EllipsePts = p
End Function

Public Function SineWavePts(x0 As Long, y0 As Long, scaleX As Long, scaleY As Long, Angle As Long) As POINTAPI()
    Dim p(24) As POINTAPI
    Dim Y(24) As Single, r As Single, alfa As Single
    Dim x1 As Single, y1 As Single
    
    'y(0) =0: y(12) = 0: y(24) = 0
    Y(1) = 2 * k - k1
    Y(2) = 4 * k - k2
    Y(3) = Sqr(2) / 2
    Y(4) = 3 * k + k2
    Y(5) = 1
    Y(6) = 1
    Y(7) = 1
    Y(8) = Y(4)
    Y(9) = Y(3)
    Y(10) = Y(2)
    Y(11) = Y(1)
    Y(13) = -Y(1)
    Y(14) = -Y(2)
    Y(15) = -Y(3)
    Y(16) = -Y(4)
    Y(17) = -1
    Y(18) = -1
    Y(19) = -1
    Y(20) = Y(16)
    Y(21) = Y(15)
    Y(22) = Y(14)
    Y(23) = Y(13)
    
    For i = 0 To 24
        x1 = scaleX * i * PI12
        y1 = scaleY * Y(i)
        r = Sqr(x1 ^ 2 + y1 ^ 2)
        alfa = Atn2(y1, x1) + Angle * PI / 180
        p(i).X = x0 + r * Cos(alfa)
        p(i).Y = y0 - r * Sin(alfa)
    Next i
    
    SineWavePts = p

End Function

Private Function Atn2(Y As Single, X As Single) As Single
    If X = 0 Then
        Atn2 = IIf(Y = 0, PI / 4, Sgn(Y) * PI / 2)
    Else
        Atn2 = Atn(Y / X) + (1 - Sgn(X)) * PI / 2
    End If
End Function

