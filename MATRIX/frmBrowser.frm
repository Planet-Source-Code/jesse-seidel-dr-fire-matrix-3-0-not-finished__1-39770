VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "PyroNet Explorer"
   ClientHeight    =   10695
   ClientLeft      =   180
   ClientTop       =   195
   ClientWidth     =   15255
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   615
      Left            =   480
      Picture         =   "frmBrowser.frx":0442
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "&Forward"
      Height          =   615
      Left            =   1440
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   615
      Left            =   2640
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Home"
      Height          =   615
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search"
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   255
      Left            =   14520
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8775
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   15015
      ExtentX         =   26485
      ExtentY         =   15478
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "about:blank"
      Top             =   1080
      Width           =   14055
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   2520
      Y1              =   960
      Y2              =   360
   End
   Begin VB.Label Label3 
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
      Left            =   15000
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   15240
      X2              =   15240
      Y1              =   10680
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   15240
      Y1              =   10680
      Y2              =   10680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   10680
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Matrix Internet Browser"
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
      TabIndex        =   5
      Top             =   0
      Width           =   15255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "PyroNet Explorer -=- Ready"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   10320
      Width           =   15015
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AllowPopup As Boolean 'This is for Pop-up windows
Option Explicit

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Close_Click()
End
End Sub

Private Sub cmdBack_Click()
WebBrowser1.GoBack
End Sub

Private Sub cmdForward_Click()
WebBrowser1.GoForward
End Sub

Private Sub cmdGo_Click()
WebBrowser1.Navigate txtAddress.Text
lblStatus.Caption = "Going to: " & txtAddress.Text
End Sub

Private Sub cmdRefresh_Click()

WebBrowser1.Refresh
End Sub

Private Sub cmdStop_Click()

WebBrowser1.Stop
End Sub

Private Sub mnuAbout_Click()

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOptionsAllow_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
WebBrowser1.GoHome
End Sub

Private Sub Command2_Click()
WebBrowser1.GoSearch
End Sub

Private Sub Form_Load()
WebBrowser1.GoHome
End Sub

Private Sub List1_Click()

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos frmBrowser.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Label3_Click()
frmBrowser.Hide
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

lblStatus.Caption = "Done Loading"
Label2.Caption = " Matrix Internet Browser -=- " & WebBrowser1.LocationName
txtAddress.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_DownloadBegin()

lblStatus.Caption = "Loading..."
txtAddress.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_DownloadComplete()

lblStatus.Caption = "Download Done!"
txtAddress.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

lblStatus.Caption = "Done Loading!"
frmBrowser.Caption = "Matrix Internet Browser -=- " & WebBrowser1.LocationName  'Shows webpage in title bar
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)

If AllowPopup = True Then
    Cancel = False
    DoEvents
ElseIf AllowPopup = False Then
    Cancel = True
End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

lblStatus.Caption = Text
End Sub

Function FileExist(vFile As String) As Boolean
    On Error Resume Next
    FileExist = False
    If Dir$(vFile) <> "" Then: FileExist = True
End Function
