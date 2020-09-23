VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Classic Mixer"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   285
      Left            =   390
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   90
      Top             =   120
   End
   Begin VB.PictureBox PicSide 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3480
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3450
      ScaleWidth      =   330
      TabIndex        =   0
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox PicSide 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3480
      Index           =   1
      Left            =   9105
      Picture         =   "Form1.frx":1588
      ScaleHeight     =   3450
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox PicBar 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":2AB6
      ScaleHeight     =   2535
      ScaleWidth      =   405
      TabIndex        =   2
      Top             =   0
      Width           =   405
      Begin VB.PictureBox PicSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   0
         Left            =   90
         Picture         =   "Form1.frx":D8F2
         ScaleHeight     =   510
         ScaleWidth      =   225
         TabIndex        =   3
         Top             =   1980
         Width           =   255
      End
      Begin VB.Label Lblval 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   90
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   210
         X2              =   210
         Y1              =   705
         Y2              =   2255
      End
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   3
      Left            =   1410
      Picture         =   "Form1.frx":E20E
      Top             =   3510
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   1
      Left            =   810
      Picture         =   "Form1.frx":EB3C
      Top             =   3510
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   2
      Left            =   1110
      Picture         =   "Form1.frx":F41F
      Top             =   3510
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   0
      Left            =   510
      Picture         =   "Form1.frx":FD35
      Top             =   3510
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Slider #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   390
      TabIndex        =   5
      Top             =   2640
      Width           =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MX As Integer
Dim MY As Integer
Dim MS As Integer
Dim SelSlider As Integer
Dim SliderValue(255) As Integer
Private Sub SetRandomValues()
Dim SetValue As Integer
For k = PicBar.LBound To PicBar.UBound
    Randomize
    SetValue = Int((100 * Rnd) + 1)
    Call PicBar_MouseMove(CInt(k), 1, 0, 0, Line1(k).Y1 + ((100 - SetValue) * ((Line1(k).Y2 - Line1(k).Y1) / 100)))
Next k
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case vbKeyLeft
    Lblval(SelSlider).ForeColor = vbWhite
    PicSlider(SelSlider).Picture = Image1(0).Picture
    If SelSlider = 0 Then SelSlider = PicBar.UBound + 1
    SelSlider = SelSlider - 1
    Label1 = "Slider" + Str(SelSlider + 1)
    Lblval(SelSlider).ForeColor = vbRed
    PicSlider(SelSlider).Picture = Image1(2).Picture
Case vbKeyRight
    Lblval(SelSlider).ForeColor = vbWhite
    PicSlider(SelSlider).Picture = Image1(0).Picture
    If SelSlider = PicBar.UBound Then SelSlider = -1
    SelSlider = SelSlider + 1
    Label1 = "Slider" + Str(SelSlider + 1)
    Lblval(SelSlider).ForeColor = vbRed
    PicSlider(SelSlider).Picture = Image1(2).Picture
Case vbKeyUp
    Call PicBar_MouseMove(SelSlider, 1, 0, 0, Line1(SelSlider).Y1 + ((100 - SliderValue(SelSlider) - 2) * ((Line1(SelSlider).Y2 - Line1(SelSlider).Y1) / 100)))
Case vbKeyDown
    Call PicBar_MouseMove(SelSlider, 1, 0, 0, Line1(SelSlider).Y1 + ((100 - SliderValue(SelSlider) + 2) * ((Line1(SelSlider).Y2 - Line1(SelSlider).Y1) / 100)))
Case vbKeyPageUp
    Call PicBar_MouseMove(SelSlider, 1, 0, 0, Line1(SelSlider).Y1 + ((100 - SliderValue(SelSlider) - 8) * ((Line1(SelSlider).Y2 - Line1(SelSlider).Y1) / 100)))
Case vbKeyPageDown
    Call PicBar_MouseMove(SelSlider, 1, 0, 0, Line1(SelSlider).Y1 + ((100 - SliderValue(SelSlider) + 8) * ((Line1(SelSlider).Y2 - Line1(SelSlider).Y1) / 100)))
Case vbKeySpace
    Call SetRandomValues
Case vbKeyEscape
    Unload Me
Case vbKey0
    For k = PicBar.LBound To PicBar.UBound
        Call PicBar_MouseMove(CInt(k), 1, 0, 0, Line1(k).Y1 + ((100) * ((Line1(k).Y2 - Line1(k).Y1) / 100)))
    Next k
Case vbKeyM
    For k = PicBar.LBound To PicBar.UBound
        Call PicBar_MouseMove(CInt(k), 1, 0, 0, Line1(k).Y1)
    Next k

End Select
End Sub

Private Sub Form_Load()
Dim MaxBars As Integer
MaxBars = Fix(((Screen.Width / 2) - (2 * PicSide(0).Width)) / (PicBar(0).Width + 15))

Line1(0).Y2 = PicBar(0).Height - 370

For k = 1 To MaxBars - 1
    Load PicBar(k)
    Load PicSlider(k)
    Load Line1(k)
    Load Lblval(k)
    PicBar(k).Left = PicBar(k - 1).Left + PicBar(k - 1).Width + 15
    Set PicSlider(k).Container = PicBar(k)
    Set Line1(k).Container = PicBar(k)
    Set Lblval(k).Container = PicBar(k)
    PicBar(k).Visible = True
    PicSlider(k).Visible = True
    Line1(k).Y2 = PicBar(k).Height - 370
    Line1(k).Visible = True
    Lblval(k).Visible = True
Next k
Me.Width = (2 * PicSide(0).Width) + (k * (PicBar(0).Width + 15))
Call SetRandomValues
End Sub


Private Sub Lblval_DblClick(Index As Integer)
Call PicBar_MouseMove(Index, 1, 0, 0, Line1(Index).Y1 + ((100 - 0) * ((Line1(Index).Y2 - Line1(k).Y1) / 100)))
End Sub


Private Sub Lblval_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MS = Lblval(Index)
End Sub


Private Sub Lblval_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Lblval(Index) = MS - Fix(Y / 20)
If Lblval(Index) < 0 Then Lblval(Index) = 0
If Lblval(Index) > 100 Then Lblval(Index) = 100

End Sub


Private Sub Lblval_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicBar_MouseMove(Index, 1, 0, 0, Line1(Index).Y1 + ((100 - Lblval(Index)) * ((Line1(Index).Y2 - Line1(k).Y1) / 100)))
End Sub

Private Sub PicBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicBar(Index).MousePointer = 7
If Button = 1 Then Call PicBar_MouseMove(Index, Button, Shift, X, Y)
If Button = 2 Then Call PicBar_MouseMove(Index, 1, Shift, X, Line1(Index).Y1 + ((100 - 0) * ((Line1(Index).Y2 - Line1(k).Y1) / 100)))
End Sub


Private Sub PicBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
If SelSlider <> Index And Button = 1 Then
    Lblval(SelSlider).ForeColor = vbWhite
    PicSlider(SelSlider).Picture = Image1(0).Picture
    SelSlider = Index
    Label1 = "Slider" + Str(SelSlider + 1)
    Lblval(SelSlider).ForeColor = vbRed
    PicSlider(SelSlider).Picture = Image1(2).Picture
End If
If Shift = 1 Then
    Lblval(Index) = 100
    PicSlider(Index).Top = Line1(Index).Y1 - PicSlider(Index).Height / 2
End If

If Shift = 2 Then
    Lblval(Index) = 0
    PicSlider(Index).Top = Line1(Index).Y2 - PicSlider(Index).Height / 2
End If

If Button = 1 Or Shift = 4 Then
    Select Case Y
    Case Is < Line1(Index).Y1
        PicSlider(Index).Top = Line1(Index).Y1 - PicSlider(Index).Height / 2
        Lblval(Index) = 100
    Case Is > Line1(Index).Y2
        PicSlider(Index).Top = Line1(Index).Y2 - PicSlider(Index).Height / 2
        Lblval(Index) = 0
    Case Else
        PicSlider(Index).Top = Y - PicSlider(Index).Height / 2
        Lblval(Index) = 100 - (Fix(100 / (Line1(Index).Y2 - Line1(Index).Y1) * (Y - Line1(Index).Y1)))
    End Select
    SliderValue(Index) = Lblval(Index)
End If

SliderValue(Index) = Lblval(Index)

End Sub







Private Sub PicBar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicBar(Index).MousePointer = 0
End Sub

Private Sub PicSide_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MX = X
MY = Y
End Sub

Private Sub PicSide_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then Me.Move Me.Left + X - MX, Me.Top + Y - MY
End Sub


