VERSION 5.00
Begin VB.Form frmopen 
   BorderStyle     =   0  'None
   Caption         =   "Itech PaintPro Open     -Doubleclick to open-"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3240
      Pattern         =   "*.bmp*;*.gif*;*.jpg*"
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6615
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   6600
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   6840
      X2              =   0
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      X1              =   6840
      X2              =   6840
      Y1              =   240
      Y2              =   6600
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Open"
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
      TabIndex        =   3
      Top             =   0
      Width           =   6855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   600
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   5535
   End
End
Attribute VB_Name = "frmopen"
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
Private Sub Dir1_Change()
File1.Path = Dir1.Path


End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
 SelectedFile = File1.Path & "\" & File1.FileName
   Image1.Picture = LoadPicture(SelectedFile)
End Sub

Private Sub File1_DblClick()
    SelectedFile = File1.Path & "\" & File1.FileName
    frmpaint.picBoard.Picture = LoadPicture(SelectedFile)
    frmopen.Visible = False
    













End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos frmBrowser.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub
