VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Cirus Pad"
   ClientHeight    =   7155
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8775
   Icon            =   "frmPad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text1 
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10610
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmPad.frx":0442
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   6870
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sb2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   6900
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
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
      Left            =   8520
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.Label open 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Open"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   495
   End
   Begin VB.Label print 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P&rint"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label selall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Select &All"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Paste 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  &Paste"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Copy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Copy"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Save 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  &Save"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   6840
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   8760
      X2              =   0
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      X1              =   8760
      X2              =   8760
      Y1              =   240
      Y2              =   6840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pyro Pad"
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
      Width           =   8775
   End
End
Attribute VB_Name = "frmPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CharCount As Boolean
Option Explicit
Dim a As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long




Private Sub about_Click()

End Sub

Private Sub color_Click()
On Error Resume Next
cd1.ShowColor
Text1.SelColor = cd1.Color
End Sub


Private Sub Command1_Click()
MsgBox Text1.Text & vbTab & Len(Text1.Text)
End Sub

Private Sub copy_Click()
If Text1.SelLength > 0 Then SendKeys ("^c")
End Sub

Private Sub cut_Click()
If Text1.SelLength > 0 Then SendKeys ("^x")
End Sub

Private Sub dclc_Click()
If CharCount = True Then
CharCount = False
Chars_Lines
Else
CharCount = True
Chars_Lines
End If

End Sub

Private Sub decrypt_Click()
Dim AsciiOf As Integer
Dim NewText As String
Dim OldText As String
Dim X As Long
OldText = Text1.Text
Label1.Caption = "Pyro Pad - DeEncrypting..."
Text1.Text = "DeEncrypting..."

For X = 1 To Len(OldText)
    DoEvents
    AsciiOf = Asc(Mid(OldText, X, 1))
    If AsciiOf <= 25 Then AsciiOf = AsciiOf + 255
    NewText = NewText & Chr(AsciiOf - 25)
Next

Text1.Text = NewText
Label1.Caption = "Pyro Pad"
Call Chars_Lines
End Sub

Private Sub delete_Click()
If Text1.SelLength > 0 Then SendKeys "{DEL}"
End Sub

Private Sub encrypt_Click()
Dim Letter1 As String
Dim AsciiOf As Integer
Dim NewText As String
Dim MemText As String
Dim X As Long
MemText = Text1.Text
Label1.Caption = "Pyro Pad - Encrypting...."
Text1.Text = "Encrypting..."
For X = 1 To Len(MemText)
DoEvents
Letter1 = Mid(MemText, X, 1)
AsciiOf = Asc(Letter1)
AsciiOf = AsciiOf + 25
If AsciiOf > 255 Then AsciiOf = AsciiOf - 255
NewText = NewText & Chr(AsciiOf)
Next
Text1.Text = NewText
Label1.Caption = "Pyro Pad"
Call Chars_Lines
End Sub

Private Sub exit_Click()
Unload Me
End
End Sub

Private Sub ExitMe()
Dim a As Integer
If Text1.Text <> "" Then
    a = MsgBox("Would you like to save before exiting?", vbYesNoCancel, "Save?")
    If a = vbYes Then
        Call Save_Click
    End If
    If a = vbCancel Then
        Exit Sub
    End If
End If

Unload Me
End

End Sub

Private Sub font_Click()
cd1.Flags = cdlCFScreenFonts
cd1.ShowFont
Text1.SelFontName = cd1.FontName
Text1.SelBold = cd1.FontBold
Text1.SelItalic = cd1.FontItalic
Text1.SelFontSize = cd1.FontSize
Text1.SelStrikeThru = cd1.FontStrikethru
Text1.SelUnderline = cd1.FontUnderline
End Sub

Private Sub Form_Load()
Me.Show
Me.Refresh

sb2.SimpleText = "Char:0/0  Line:1/1"
If Label1.Caption = "Pyro Pad" Then
Else
Label1.Caption = "Pyro Pad"
End If
CharCount = True
sb2.Panels(1) = "hi"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ExitMe
Cancel = 1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos frmPad.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Label5_Click()
frmPad.Hide
End Sub

Private Sub Open_Click()
On Error GoTo err
Dim i As Long
Text1.Text = ""
cd1.Filter = "Txt (*.txt)|*.txt|Any File (*.*)|*.*"
cd1.ShowOpen
If cd1.FileName <> "" Then
Dim t As Long
i = FreeFile
Open cd1.FileName For Input As #i
CharCount = True
If Int(LOF(i) / 1000) > 300 Then
    If MsgBox("File is > 300kb would you like to disable Charactor coutning?", vbYesNo, "Large FIle") = vbYes Then
    CharCount = False
    Chars_Lines
    End If
End If

Text1.Text = Input(LOF(i), i)
Close #i
Else
Exit Sub
End If



Call Chars_Lines
Exit Sub



err:
Close #i
Open cd1.FileName For Binary As #i
Text1.Text = Input(LOF(i), i)
Close #i

Exit Sub




End Sub

Private Sub Label2_Click()

End Sub

Private Sub paste_Click()
SendKeys ("^v")
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub Print_Click()
Printer.Print Text1.Text
End Sub

Private Sub psetup_Click()
cd1.ShowPrinter

End Sub

Private Sub Save_Click()
'On Error GoTo err
Label1.Caption = "Pyro Pad - Saving..."
Dim a As String
cd1.Filter = "Txt (*.txt)|*.txt|Html File (*.Html)|*.Html|PHP Script (*.Php)|*.php|CGI Script (*.Cgi)|*cgi|Any File (*.*)|*.*"
cd1.ShowSave
If cd1.FileName <> "" Then
Open cd1.FileName For Output As #1
Print #1, Text1.Text
Close 1
End If
Label1.Caption = "Pyro Pad"
Exit Sub

err:
Label1.Caption = "Pyro Pad"
MsgBox "Error saving file"
End Sub

Private Sub selall_Click()
Dim a As String
Text1.SelStart = 0

Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub td_Click()
SendKeys (Now)
End Sub




Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Call Chars_Lines

End Sub

Private Sub Chars_Lines()
If CharCount = True Then
Dim Lines, Chars As String
Dim blah() As String
Dim bleh() As String
Dim Curline As String
Dim CurChar, TotalChar As String

Curline = Mid(Text1.Text, 1, Text1.SelStart)
blah() = Split(Curline, Chr$(10))
bleh() = Split(Text1.Text, Chr$(10))

If Text1.SelStart = 0 Then
CurChar = 0
Curline = 1

If Len(Text1.Text) = 0 Then
TotalChar = 0
Else
TotalChar = Len(Text1.Text) - (UBound(bleh) * 2)

End If

Else
CurChar = Text1.SelStart - (UBound(blah) * 2)
Curline = UBound(blah) + 1
TotalChar = Len(Text1.Text) - (UBound(bleh) * 2)
End If



Lines = "Line:" & Curline & "/" & SendMessage(Text1.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
Chars = "Char:" & CurChar & "/" & TotalChar

sb2.SimpleText = Chars & "  " & Lines

Else
sb2.SimpleText = "Off"
End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Chars_Lines
End Sub

Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim numFiles As Integer
    numFiles = Data.Files.Count
    If numFiles = 1 Then
    'Add all dropped files into the list

        'File or directory?
        If (GetAttr(Data.Files(1)) And vbDirectory) = vbDirectory Then
            Else



On Error GoTo err
Dim i As Long
Label1.Caption = "Pyro Pad - Opening..."
Text1.Text = ""
i = FreeFile
Open Data.Files(1) For Input As #i
CharCount = True
If Int(LOF(i) / 1000) > 300 Then
    If MsgBox("File is > 300kb would you like to disable Charactor coutning?", vbYesNo, "Large FIle") = vbYes Then CharCount = False
End If

Text1.Text = Input(LOF(i), i)
Close #i




Label1.Caption = "Pyro Pad"
Call Chars_Lines
Exit Sub



err:
Label1.Caption = "Pyro Pad - Opening in binary..."
Close #i
Open Data.Files(1) For Binary As #i
Text1.Text = Input(LOF(i), i)
Close #i
Label1.Caption = "Pyro Pad"
Exit Sub




End If

    End If

  
End Sub

Private Sub undo_Click()
SendKeys ("^z")

End Sub

Private Sub uptime_Click()
    Dim Secs, Mins, Hours, Days As Long
    Dim TotalMins, TotalHours, TotalSecs, TempSecs As Long
    Dim CaptionText As String
    TotalSecs = Int(GetTickCount / 1000)
    Days = Int(((TotalSecs / 60) / 60) / 24)
    TempSecs = Int(Days * 86400)
    TotalSecs = TotalSecs - TempSecs
    TotalHours = Int((TotalSecs / 60) / 60)
    TempSecs = Int(TotalHours * 3600)
    TotalSecs = TotalSecs - TempSecs
    TotalMins = Int(TotalSecs / 60)
    TempSecs = Int(TotalMins * 60)
    TotalSecs = (TotalSecs - TempSecs)


    If TotalHours > 23 Then
        Hours = (TotalHours - 23)
    Else
        Hours = TotalHours
    End If


    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If
    CaptionText = "Your Computer has been up: " & Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds" & vbCrLf

    MsgBox CaptionText, vbOKOnly, "Up Time"
    Clipboard.Clear
    Clipboard.SetText CaptionText
End Sub

Private Sub wordcount_Click()
Dim a() As String
Dim b() As String
Dim wordcount As Long
Dim X As Long
Label1.Caption = "Pyro Pad - Counting Words..."

a() = Split(Text1.Text, " ")
wordcount = UBound(a)
For X = 0 To UBound(a)
If a(X) = "" Then
wordcount = wordcount - 1
End If
Next

b() = Split(Text1.Text, Chr$(10))
wordcount = wordcount + UBound(b)
For X = 0 To UBound(b)
If b(X) = "" Then
wordcount = wordcount - 1
End If
Next
If wordcount = -2 Then wordcount = -1
Label1.Caption = "Pyro Pad"
MsgBox "There are: " & wordcount + 1 & " Words", vbOKOnly, "Word"


End Sub
