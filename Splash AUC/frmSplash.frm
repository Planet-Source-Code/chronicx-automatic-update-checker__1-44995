VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmSplash 
   BackColor       =   &H00CEB0A6&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3750
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMSPL~1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton UpdateButton 
      Caption         =   "Update Me!"
      Height          =   435
      Left            =   3195
      TabIndex        =   6
      Top             =   2550
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   435
      Left            =   3195
      TabIndex        =   5
      Top             =   1845
      Width           =   1260
   End
   Begin VB.TextBox Text2 
      Height          =   2085
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "FRMSPL~1.frx":000C
      Top             =   1335
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   285
      Top             =   450
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   3780
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2625
      Width           =   420
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   120
      Left            =   3330
      TabIndex        =   7
      Top             =   1950
      Width           =   495
      ExtentX         =   873
      ExtentY         =   212
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
   Begin VB.Label Labelurl 
      Caption         =   "http://www.infiniteimpossibilities.com/products/jade.php"
      Height          =   210
      Left            =   3480
      TabIndex        =   3
      Top             =   2655
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0096B8BA&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Checking For Updates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -15
      TabIndex        =   2
      Top             =   3510
      Width           =   4710
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0096B8BA&
      X1              =   195
      X2              =   4275
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   4290
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AUC"
      BeginProperty Font 
         Name            =   "wargames"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1185
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4545
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public VVersion As Integer
Private Sub Form_Load()
Timer1.Enabled = False
UpdateButton.Enabled = False
'set an invisible web browser to goto the site to look for the page giving the version (see the included php file to see what it does
WebBrowser1.Navigate ("infiniteimpossibilities.com/JMPversion.php")
'set the current version for comparison
VVersion = App.Major + App.Minor
End Sub

Private Sub Timer1_Timer()
If WebBrowser1.Busy Then
'display what its doing
Label3.Caption = "Checking For Updates"
Else
Label3.Caption = "Done Checking"
Timer1.Interval = 10000
Dim numeral As Integer
numeral = Right(Text1, 1)
'compare the versions to test the update button and see that it works ok use the following ;)
'If (numeral < VVersion) Then
If (numeral > VVersion) Then
'this isnt shown anywhere, its behind a button.  this is basicly for your personal info as the coder
Text1.Text = Right(Text1, 1) & ":" & Str(VVersion)
Beep
UpdateButton.Enabled = True
'tell them incase they miss the button lighting
Label3.Caption = "New Updates Available!"
Else
Text1.Text = Right(Text1, 1) & ":" & Str(VVersion)
Label3.Caption = "No New Updates Are Available"
End If
End If
End Sub
Private Sub command1_Click()
'this simply bypasses this and loads the featured project
Form1.Show
Unload Me
End Sub

Private Sub updatebutton_Click()
'by El Mariachi (http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=9931&lngWId=1) this opens the internet based on another label hidden under a button for you to change
    Dim ws As String
    Dim opn As Boolean
    ws = Labelurl.Caption
        opn = True
    Call OpenLink(ws, opn)
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'now lets let the timer do its stuff, now that we have established that there is an internet connection and weve begun to load
Text1.Text = URL
Timer1.Enabled = True
End Sub
