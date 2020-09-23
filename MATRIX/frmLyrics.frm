VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLyrics 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Lyrics Finder v2.0"
   ClientHeight    =   7575
   ClientLeft      =   3225
   ClientTop       =   1935
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLyrics.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTest 
      Height          =   330
      Left            =   7170
      TabIndex        =   10
      Text            =   "<PRE style=""font:12px arial"">"
      Top             =   8040
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1170
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   5475
      Begin VB.TextBox txtArtist 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   4590
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Default         =   -1  'True
         Height          =   405
         Left            =   3765
         Picture         =   "frmLyrics.frx":628A
         TabIndex        =   7
         Top             =   645
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Artist:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   5370
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   5445
      Begin MSComDlg.CommonDialog cd 
         Left            =   3840
         Top             =   4110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Save Lyrics to File"
      End
      Begin RichTextLib.RichTextBox txtLyrics 
         Height          =   5100
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   8996
         _Version        =   393217
         ScrollBars      =   2
         MousePointer    =   1
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmLyrics.frx":D77C
      End
   End
   Begin VB.TextBox txtTemp2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   10290
      Visible         =   0   'False
      Width           =   2700
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7305
      Top             =   9390
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtNoResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   855
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmLyrics.frx":D864
      Top             =   8340
      Width           =   1965
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   7350
      Top             =   8625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.TextBox txtTemp 
      Height          =   2130
      Left            =   3090
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   7980
      Width           =   3795
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6465
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   4005
      ExtentX         =   7064
      ExtentY         =   11404
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: You MUST be online to retrieve lyrics! This is because this connects to a database"
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
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   7695
   End
   Begin VB.Label Label5 
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
      Left            =   9480
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   7560
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   9720
      X2              =   0
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line1 
      X1              =   9720
      X2              =   9720
      Y1              =   240
      Y2              =   7560
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Lyrics Search"
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
      Width           =   9735
   End
End
Attribute VB_Name = "frmLyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tUrl As String
Dim SearchState As String
Option Explicit
Dim a As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Sub Pause(duration)

Dim Current As Long
Current = Timer
Do Until Timer - Current >= duration
    DoEvents
Loop
End Sub
Private Sub Command1_Click()
Dim WebHost As String

SearchState = ""

On Error Resume Next
WebHost = "www.letssingit.com"


If txtTemp.Text <> "" Then txtTemp.Text = ""

Winsock.Close


Winsock.RemoteHost = WebHost
Winsock.RemotePort = 80

Winsock.Connect
End Sub

Private Sub Command2_Click()
MsgBox SearchState
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "about:<font face=arial size=2>Please Enter An Artist.</font>"
End Sub


Private Sub Save_Click()
cd.Filter = "All Files(*.*)|*.*|Rich Text(*.rtf)|*.rtf"
cd.FilterIndex = 2
cd.ShowSave
If cd.FileName <> "" Then
txtLyrics.SaveFile cd.FileName
End If
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

Private Sub Label5_Click()
frmLyrics.Hide
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim FinalURL
Dim theStart
Dim theStart2
Dim b() As Byte
Dim txt As String
Dim t As Integer


 If SearchState = "ArtistPage" Then

  FinalURL = Right(URL, Len(URL) - 3)
  

   If URL <> "" Then
   On Error Resume Next
    If InStr(1, URL, "templyrics", vbTextCompare) = 0 Then
    Cancel = True
    b() = Inet1.OpenURL("http://www.letssingit.com/" & FinalURL, 1)
    txt = ""
    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    txtTemp.Text = txt
    theStart = InStr(1, txtTemp.Text, "<TABLE><TR><TD><PRE", vbTextCompare)
    txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
    theStart = InStr(1, txtTemp.Text, "</PRE></TD></TR></TABLE>", vbTextCompare)
    txtTemp.Text = Left(txtTemp.Text, theStart - 1)
    txtTemp.Text = Replace(txtTemp.Text, "<TABLE><TR><TD>", "")
    txtTemp.Text = Replace(txtTemp.Text, txtTest, "")
    txtLyrics.TextRTF = txtTemp.Text
    End If
   End If
    
 ElseIf SearchState = "Results" Then
 FinalURL = Right(URL, Len(URL) - 3)
 Cancel = True


     b() = Inet1.OpenURL("http://www.letssingit.com/" & FinalURL, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    txtTemp.Text = txt
    

    On Error Resume Next

theStart = InStr(1, txtTemp.Text, "</TR>" & vbCrLf & "</TABLE>", vbTextCompare)
txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
theStart = InStr(1, txtTemp.Text, "<TABLE><TR><TD>", vbTextCompare)
theStart2 = InStr(1, txtTemp.Text, "</TR></TABLE>", vbTextCompare)
txtTemp.Text = Left(txtTemp.Text, theStart2 - 1)
txtTemp.Text = Replace(txtTemp.Text, vbCrLf, "")


txtTemp.Text = Replace(txtTemp.Text, "</TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "13de", "")
txtTemp.Text = Replace(txtTemp.Text, "<TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "</TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TD>", "")
txtTemp.Text = Replace(txtTemp.Text, "11fc", "")
txtTemp.Text = Replace(txtTemp.Text, "7f5", "")

txtTemp.Text = "<font face=arial size=2>" & txtTemp.Text & "</font>"


Dim f As Integer
f = FreeFile
Kill "C:\templyrics000.html"
Open "C:\templyrics000.html" For Binary As #f
Put #f, , txtTemp.Text
Close #f

SearchState = ""

WebBrowser1.Navigate "C:\templyrics000.html"

SearchState = "ArtistPage"

 End If
 Exit Sub
End Sub

Private Sub WebBrowser2_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Winsock_Close()
Dim theStart
Dim theStart2

SearchState = ""


Pause 0.5


Winsock.Close: Winsock.Tag = "CLOSED"


On Error Resume Next


 
If InStr(1, txtTemp.Text, "no search results", vbTextCompare) <> 0 Then
 Dim f2 As Integer
 f2 = FreeFile
 Kill "C:\templyrics000.html"
 Open "C:\templyrics000.html" For Binary As #f2
 Put #f2, , txtNoResults.Text
 Close #f2
 WebBrowser1.Navigate "C:\templyrics000.html"
 SearchState = "None"
Exit Sub
ElseIf InStr(1, txtTemp.Text, "<TABLE><TR><TD>Showing", vbTextCompare) <> 0 Then
' show search results (parsing)
 theStart = InStr(1, txtTemp.Text, "</TD></TR></TABLE><BR>", vbTextCompare)
 txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
 theStart = InStr(1, txtTemp.Text, "<BR><BR>Select page", vbTextCompare)
 txtTemp.Text = Left(txtTemp.Text, theStart - 1)
 txtTemp.Text = "<font face=arial size=2>" & txtTemp.Text & "</font>"
 Dim f3 As Integer
 f3 = FreeFile
 Kill "C:\templyrics000.html"
 Open "C:\templyrics000.html" For Binary As #f3
 Put #f3, , txtTemp.Text
 Close #f3
 SearchState = "Results"
 WebBrowser1.Navigate "C:\templyrics000.html"
Exit Sub
End If

theStart = InStr(1, txtTemp.Text, "</TR>" & vbCrLf & "</TABLE>", vbTextCompare)
txtTemp.Text = Mid(txtTemp.Text, theStart, Len(txtTemp.Text) - theStart)
theStart = InStr(1, txtTemp.Text, "<TABLE><TR><TD>", vbTextCompare)
theStart2 = InStr(1, txtTemp.Text, "</TR></TABLE>", vbTextCompare)
txtTemp.Text = Left(txtTemp.Text, theStart2 - 1)
txtTemp.Text = Replace(txtTemp.Text, vbCrLf, "")

txtTemp.Text = Replace(txtTemp.Text, "</TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "13de", "")
txtTemp.Text = Replace(txtTemp.Text, "<TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "</TABLE>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TR>", "")
txtTemp.Text = Replace(txtTemp.Text, "<TD>", "")
txtTemp.Text = Replace(txtTemp.Text, "<SCRIPT>", "")
txtTemp.Text = Replace(txtTemp.Text, "</SCRIPT>", "")
txtTemp.Text = Replace(txtTemp.Text, "11fc", "")
txtTemp.Text = Replace(txtTemp.Text, "7f5", "")

txtTemp.Text = "<font face=arial size=2>" & txtTemp.Text & "</font>"

Dim f As Integer
f = FreeFile
Kill "C:\templyrics000.html"
Open "C:\templyrics000.html" For Binary As #f
Put #f, , txtTemp.Text
Close #f

SearchState = ""

WebBrowser1.Navigate "C:\templyrics000.html"

SearchState = "ArtistPage"
End Sub

Private Sub Winsock_Connect()
Dim getString As String, ShortWebSite As String
Winsock.Tag = "OPEN"
On Error Resume Next

Winsock.SendData FindArtist(txtArtist.Text)
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
On Error Resume Next

If Winsock.Tag = "OPEN" Then Winsock.GetData Buffer

txtTemp.Text = txtTemp.Text & Buffer
End Sub

