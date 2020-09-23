VERSION 5.00
Begin VB.Form frmunload 
   BorderStyle     =   0  'None
   Caption         =   "Itech PaintPro "
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &No"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  &Yes"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label2 
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
      TabIndex        =   1
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you wish to save the changes you've made to your drawing?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmunload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Command1_Click()
Dim Save As Image
frmpaint.CommonDialog1.Filter = "Bitmap Files .bmp|*.bmp"
frmpaint.CommonDialog1.ShowSave
SavePicture frmpaint.picBoard.Image, frmpaint.CommonDialog1.FileName
End

End Sub

Private Sub Command2_Click()
frmunload.Hide

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos frmBrowser.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub
