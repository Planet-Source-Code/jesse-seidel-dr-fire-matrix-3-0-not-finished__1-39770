VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -195
   ClientWidth     =   15360
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   13560
      Top             =   11040
   End
   Begin VB.TextBox Time1 
      BackColor       =   &H00808080&
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
      Height          =   285
      Left            =   14040
      TabIndex        =   0
      Text            =   "THE MATRIX"
      Top             =   11160
      Width           =   1215
   End
   Begin VB.Image Image9 
      Height          =   1185
      Left            =   1200
      Picture         =   "frmMain.frx":0045
      Top             =   120
      Width           =   885
   End
   Begin VB.Image Image8 
      Height          =   1380
      Left            =   240
      Picture         =   "frmMain.frx":021E
      Top             =   5160
      Width           =   840
   End
   Begin VB.Image Image7 
      Height          =   750
      Left            =   240
      Picture         =   "frmMain.frx":0458
      Top             =   4320
      Width           =   675
   End
   Begin VB.Image Image6 
      Height          =   870
      Left            =   120
      Picture         =   "frmMain.frx":0621
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image Image5 
      Height          =   870
      Left            =   120
      Picture         =   "frmMain.frx":0771
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   1020
      Left            =   120
      Picture         =   "frmMain.frx":0914
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   1125
      Left            =   120
      Picture         =   "frmMain.frx":0A9A
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Lyric Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   240
      X2              =   2160
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "MORE OPTIONS COMING SOON"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   16
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   2160
      TabIndex        =   15
      Top             =   10080
      Width           =   4335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   0
      TabIndex        =   14
      Top             =   6600
      Width           =   5535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Source Getter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   9480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "CD Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Paint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   8760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Website Hitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Pyro Pad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Matrix Internet Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   2160
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   9480
      Y2              =   9960
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   2160
      X2              =   2160
      Y1              =   9600
      Y2              =   9960
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   1920
      Picture         =   "frmMain.frx":0C99
      Top             =   9600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   " My &Programs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " &About The Matrix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Matrix Live &Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   10320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " &Return To Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   10680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   0
      Picture         =   "frmMain.frx":0CE3
      Top             =   7920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   11040
      Y2              =   7800
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &MATRIX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   11160
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   0
      Top             =   11040
      Width           =   15375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   4455
      Left            =   4320
      TabIndex        =   13
      Top             =   6600
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Const SPI_SCREENSAVERRUNNING = 97

Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByVal lpvParam As Any, _
    ByVal fuWinIni As Long) As Long
Dim StartX, StartY
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Click()
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    Source.Move x - StartX, y - StartY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 0
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 0
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Image2_Click()
Label5(1).BackStyle = 1
Shape3.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label17.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 1
Shape3.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label17.Visible = True
End Sub



Private Sub Image3_DblClick()
frmBrowser.Show
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image3.Drag vbBeginDrag
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image3.Drag vbEndDrag
End Sub

Private Sub Image4_DblClick()
frmPad.Show
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image4.Drag vbBeginDrag
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image4.Drag vbEndDrag
End Sub

Private Sub Image5_DblClick()
frmHitter.Show
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image5.Drag vbBeginDrag
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image5.Drag vbEndDrag
End Sub

Private Sub Image6_DblClick()
frmCalc.Show
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image6.Drag vbBeginDrag
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image6.Drag vbEndDrag
End Sub

Private Sub Image7_DblClick()
frmpaint.Show
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image7.Drag vbBeginDrag
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image7.Drag vbEndDrag
End Sub

Private Sub Image8_DblClick()
frmHTML.Show
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image8.Drag vbBeginDrag
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image8.Drag vbEndDrag
End Sub

Private Sub Image9_DblClick()
frmLyrics.Show
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    StartX = x
    StartY = y
    Image9.Drag vbBeginDrag
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image9.Drag vbEndDrag
End Sub

Private Sub Label1_Click()
Shape2.Visible = True
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Image1.Visible = True
Label2.Visible = True
Label3(0).Visible = True
Label4(1).Visible = True
Label5(1).Visible = True
Image2.Visible = True
Label16.Visible = True
Line4.Visible = True
End Sub



Private Sub Label10_Click()
frmpaint.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label12_Click()
frmHTML.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label13_Click()
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Label17.Visible = False
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 0
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label14_Click()
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Label17.Visible = False
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 0
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label15_Click()
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Label17.Visible = False
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 0
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label17_Click()
frmLyrics.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label4_Click(Index As Integer)
frmAbout.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 0
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Label5(1).BackStyle = 1
Shape3.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label17.Visible = True
End Sub

Private Sub Label6_Click()
frmBrowser.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label7_Click()
frmPad.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label8_Click()
frmHitter.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Label9_Click()
frmCalc.Show
Shape2.Visible = False
Line1.Visible = False
Image1.Visible = False
Label2.Visible = False
Label3(0).Visible = False
Label4(1).Visible = False
Label5(1).Visible = False
Image2.Visible = False
Line2.Visible = False
Line3.Visible = False
Label16.Visible = False
Line4.Visible = False
Shape3.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label17.Visible = False
End Sub

Private Sub Timer1_Timer()
Time1.Text = Time
End Sub

Private Sub DisableCtrlAltDel()

RegisterServiceProcess GetCurrentProcessId, 1

End Sub
Private Sub EnableCtrlAltDel()

RegisterServiceProcess GetCurrentProcessId, 0

End Sub


Private Sub frmMain_Hide()
    Call EnableCtrlAltDel

    
    Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, "1", 0)
End Sub
