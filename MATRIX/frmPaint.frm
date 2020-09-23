VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmpaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   " I tech SketchPad"
   ClientHeight    =   7215
   ClientLeft      =   495
   ClientTop       =   165
   ClientWidth     =   10455
   DrawStyle       =   5  'Transparent
   Icon            =   "frmPaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPaint.frx":030A
   Palette         =   "frmPaint.frx":045C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   7215
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstTools 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   8160
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   6465
      ScaleWidth      =   7785
      TabIndex        =   9
      Top             =   360
      Width           =   7815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   3
      Left            =   8160
      Max             =   25
      Min             =   2
      TabIndex        =   7
      Top             =   840
      Value           =   3
      Width           =   1935
   End
   Begin VB.Timer tmrCursor 
      Interval        =   1
      Left            =   480
      Top             =   5520
   End
   Begin VB.PictureBox pCol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   8160
      MouseIcon       =   "frmPaint.frx":2132
      MousePointer    =   99  'Custom
      Picture         =   "frmPaint.frx":243C
      ScaleHeight     =   1110
      ScaleWidth      =   2145
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X"
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
      Left            =   10200
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Open 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Open"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Save 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Save"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   12
      Top             =   5280
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   7200
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   10440
      X2              =   0
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line1 
      X1              =   10440
      X2              =   10440
      Y1              =   240
      Y2              =   7200
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Paint"
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
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10455
   End
   Begin VB.Image target 
      Height          =   480
      Left            =   2280
      Picture         =   "frmPaint.frx":A15E
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Color"
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
      Left            =   8640
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to choose pen/fill color"
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
      Left            =   7920
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image bucket 
      Height          =   480
      Left            =   1440
      Picture         =   "frmPaint.frx":A468
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image curpencil 
      Height          =   480
      Left            =   960
      Picture         =   "frmPaint.frx":A5BA
      Top             =   4440
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   8160
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pen Size"
      Height          =   255
      Left            =   8160
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblPenSize 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   9120
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmpaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type pointapi
   x As Double
   y As Double
End Type

Option Explicit

Dim pressed As Boolean
Dim colpressed As Boolean
Dim filltool As Boolean
Dim drawtool As Boolean
Dim circletool As Boolean
Dim whatradius As Variant
Dim eyedroptool As Boolean
Dim circgradient As Boolean
Dim radius As Integer
Dim about
Dim onlynumbers
Dim point1 As pointapi
Dim point2 As pointapi
Dim g1
Dim g2
Dim g3
Dim cgformat
Dim gformat
Dim gdirection
Dim lf1
Dim lr2
Dim lr3
Dim lr4
Dim ud1
Dim ud2
Dim ud3
Dim ud4
Dim cg1
Dim cg2
Dim cg3
Dim cg4
Dim Index
Dim index2
Dim index3
Dim index4
Dim i As Integer
Dim a As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub cmdExit_Click()
Unload Me
End Sub



Private Sub Command1_Click()
picBoard.BackColor = &H80000009
End Sub


Private Sub Form_Load()

HScroll1.Value = 2
picBoard.DrawWidth = 2
picBoard.MouseIcon = curpencil
pCol.MouseIcon = target
drawtool = True
lstTools.AddItem ("Pen")
lstTools.AddItem ("Circle")
lstTools.AddItem ("Paint Bucket")
lstTools.AddItem ("Eye Dropper")
lstTools.AddItem ("Gradient")
lstTools.AddItem ("Circular Gradient")
lstTools.AddItem ("Clear")
picBoard.ScaleHeight = 255
picBoard.ScaleWidth = 255
End Sub



Private Sub HScroll1_Change()
lblPenSize.Caption = HScroll1.Value
picBoard.DrawWidth = HScroll1.Value
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos frmBrowser.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Label7_Click()
frmpaint.Hide
frmunload.Visible = True
End Sub

Private Sub lstTools_Click()
If lstTools.Text = "Pen" Then
drawtool = True
filltool = False
circletool = False
eyedroptool = False
circgradient = False
End If

If lstTools.Text = "Circle" Then

On Error Resume Next
drawtool = False
filltool = False
circletool = True
eyedroptool = False
circgradient = False
GetRadius:
whatradius = InputBox("Enter the radius for the circle in pixels:", "Paint")
If IsNumeric(whatradius) Or radius = "" Then
radius = Val(whatradius)
Else
onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo GetRadius
End If

End If
If lstTools.Text = "Paint Bucket" Then
filltool = True
drawtool = False
circletool = False
eyedroptool = False
circgradient = False
End If
If lstTools.Text = "Clear" Then
picBoard.BackColor = &H80000009
End If
If lstTools.Text = "Eye Dropper" Then
eyedroptool = True
drawtool = False
filltool = False
circletool = False
circgradient = False
picBoard.MouseIcon = target
End If
If lstTools.Text = "Gradient" Then
Call gradientmaker
End If
If lstTools.Text = "Circular Gradient" Then
filltool = False
drawtool = False
circletool = False
eyedroptool = False
circgradient = True
HScroll1.Value = 11
End If
End Sub



Private Sub New_Click()
picBoard.BackColor = &H80000009




End Sub

Private Sub Open_Click()
frmopen.Visible = True








 









End Sub

Private Sub pCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
colpressed = True
Shape1.FillColor = pCol.Point(x, y)
picBoard.ForeColor = pCol.Point(x, y)
End Sub


Private Sub pCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If colpressed Then
Shape1.FillColor = pCol.Point(x, y)
picBoard.ForeColor = pCol.Point(x, y)
End If
End Sub

Private Sub pCol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
colpressed = False
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

pressed = True
point1.x = x
point1.y = y
If filltool = True Then
picBoard.BackColor = Shape1.FillColor
End If
If drawtool = True Then
picBoard.Line (x, y)-(x, y)
End If
If circletool = True Then
picBoard.Circle (point1.x, point1.y), radius
End If
If eyedroptool = True Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)
picBoard.ForeColor = picBoard.Point(x, y)
End If


If pressed And circgradient Then
On Error GoTo errhandler
cgformat = InputBox("How do you want to format your gradient? (RGB),  1: ##I, 2: #I#, 3: I##, 4:III (black to white)", "Paint", "1,2,3 or 4")
If cgformat = "1" Then GoTo cg1
If cgformat = "2" Then GoTo cg2
If cgformat = "3" Then GoTo cg3
If cgformat = "4" Then GoTo cg4
cg1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")

 For Index = 1 To 400 Step 1

picBoard.Circle (x, y), Index, RGB(g1, g2, Index)
                                                                                   
pressed = False
                                                                                    
Exit Sub

cg2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")


For index2 = 1 To 400 Step 1

picBoard.Circle (x, y), index2, RGB(g1, index2, g2)
Next index2

pressed = False
Exit Sub
cg3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ")

For index3 = 1 To 400 Step 2

picBoard.Circle (x, y), index3, RGB(index3, g1, g2)
Next index3
pressed = False
Exit Sub
cg4:
For index4 = 1 To 400 Step 2

picBoard.Circle (x, y), index4, RGB(index4, index4, index4)
Next index4
pressed = False
Exit Sub
errhandler:
    MsgBox ("An error has occured")
    Exit Sub
Next
End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If pressed And drawtool Then
point2 = point1
point1.x = x
point1.y = y
picBoard.Line (point1.x, point1.y)-(point2.x, point2.y)
End If
If pressed And eyedroptool Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)
picBoard.ForeColor = picBoard.Point(x, y)
End If
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = False
End Sub

Private Sub Save_Click()
Dim Save As Image
CommonDialog1.Filter = "Bitmap Files .bmp|*.bmp"
CommonDialog1.ShowSave
SavePicture picBoard.Image, CommonDialog1.FileName




















End Sub

Private Sub Timer1_Timer()


End Sub

Private Sub tmrCursor_Timer()
If drawtool = True Then
picBoard.MouseIcon = curpencil
End If
If filltool = True Then
picBoard.MouseIcon = bucket
End If
If circletool = True Then
picBoard.MouseIcon = target
End If
End Sub

Private Sub gradientmaker()
 On Error GoTo errhandler
gdirection = InputBox("What direction do you want the gradient to fade?", "Paint", "LEFT-RIGHT or UP-DOWN")
gformat = InputBox("How do you want to format your gradient? (RGB),  1: ##I, 2: #I#, 3: I##, 4:III (black to white)", "Paint", "1,2,3 or 4")
If gdirection = "LEFT-RIGHT" And gformat = "1" Then GoTo lr1
If gdirection = "LEFT-RIGHT" And gformat = "2" Then GoTo lr2
If gdirection = "LEFT-RIGHT" And gformat = "3" Then GoTo lr3
If gdirection = "LEFT-RIGHT" And gformat = "4" Then GoTo lr4

If gdirection = "UP-DOWN" And gformat = "1" Then GoTo ud1
If gdirection = "UP-DOWN" And gformat = "2" Then GoTo ud2
If gdirection = "UP-DOWN" And gformat = "3" Then GoTo ud3
If gdirection = "UP-DOWN" And gformat = "4" Then GoTo ud4
lr1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")

 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(g1, g2, i)
    Next i
Exit Sub
lr2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")


 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(g1, i, g2)
    Next i
Exit Sub
lr3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ")

 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(i, g1, g2)
    Next i
Exit Sub
lr4:
 For i = 1 To 255
    
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(i, i, i)
    Next i
Exit Sub
ud1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")


 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(g1, g2, i)
    Next i
Exit Sub
ud2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(g1, i, g2)
    Next i
Exit Sub
ud3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(i, g1, g2)
    Next i
Exit Sub
ud4:
 For i = 1 To 255
    
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(i, i, i)
    Next i
Exit Sub
errhandler:
    MsgBox ("An error has occured")
End Sub
