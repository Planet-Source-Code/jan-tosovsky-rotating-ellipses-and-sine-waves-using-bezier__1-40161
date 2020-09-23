VERSION 5.00
Begin VB.Form TestForm 
   BackColor       =   &H00000000&
   Caption         =   "TestForm"
   ClientHeight    =   3930
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5895
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuRndEllipses 
         Caption         =   "Random ellipses"
      End
      Begin VB.Menu mnuFillEllipses 
         Caption         =   "Filled ellipses"
      End
      Begin VB.Menu mnuSine 
         Caption         =   "Sine wave"
      End
      Begin VB.Menu mnuRndWaves 
         Caption         =   "Random Sine wave"
      End
      Begin VB.Menu mnuDiv 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Dim Ex As Boolean, W As Long, H As Long

Private Sub Form_Resize()
    Cls
    W = ScaleWidth
    H = ScaleHeight
End Sub

Private Sub mnuFillEllipses_Click()
    Dim Pts() As POINTAPI
    Ex = True
    Cls
    FillStyle = vbSolid
    Do While Ex
        X = Rnd * W
        Y = Rnd * H
        w1 = Rnd * 100
        h1 = Rnd * 100
        a = Rnd * 180
        Pts = EllipsePts(X, Y, w1, h1, a)
        ForeColor = QBColor(15 * Rnd)
        FillColor = QBColor(15 * Rnd)
        BeginPath Me.hdc
        PolyBezier Me.hdc, Pts(0), UBound(Pts) + 1
        EndPath Me.hdc
        StrokeAndFillPath Me.hdc
        DoEvents
    Loop
End Sub

Private Sub mnuRndEllipses_Click()
    Dim Pts() As POINTAPI
    Ex = True
    Cls
    Do While Ex
        X = Rnd * W
        Y = Rnd * H
        w1 = Rnd * 100
        h1 = Rnd * 100
        a = Rnd * 180
        Pts = EllipsePts(X, Y, w1, h1, a)
        ForeColor = QBColor(15 * Rnd)
        PolyBezier Me.hdc, Pts(0), UBound(Pts) + 1
        DoEvents
    Loop
End Sub

Private Sub mnuRndWaves_Click()
    Dim Pts() As POINTAPI, X As Long, Y As Long
    Dim sx As Long, sy As Long, a As Long
    Ex = True
    Cls
    Do While Ex
        X = Rnd * W
        Y = Rnd * H * 1.5
        sx = Rnd * 20
        sy = Rnd * 20
        a = Rnd * 180
        ForeColor = QBColor(15 * Rnd)
        Pts = SineWavePts(X, Y, sx, sy, a)
        PolyBezier Me.hdc, Pts(0), UBound(Pts) + 1
        DoEvents
    Loop
End Sub

Private Sub mnuSine_Click()
    Dim Pts() As POINTAPI
    Dim i As Long
    Cls
    Ex = False
    For i = 0 To 360 Step 4
        Pts = SineWavePts(W / 2, H / 2, 20, 20, i)
        ForeColor = RGB(240, 130, 0)
        PolyBezier Me.hdc, Pts(0), UBound(Pts) + 1
    Next i
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuEnd_Click()
    End
End Sub
