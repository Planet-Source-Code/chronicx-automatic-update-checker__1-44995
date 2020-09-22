Attribute VB_Name = "Module1"
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
         
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
         "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
         String, ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Option Explicit
'this is only thing by me, the rest is by El Mariachi (http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=9931&lngWId=1) for the internet link and sadly it only opens the splash form
Private Sub main()
Dim splash As New frmSplash
Load splash
splash.Refresh
splash.Show
End Sub
'this is only thing by me, the rest is by El Mariachi (http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=9931&lngWId=1) for the internet link and sadly it only opens the splash form

Public Sub OpenLink(WebSiteStr As String, OpenNew As Boolean)
    Dim TempFileName As String
    Dim BrowserProg As String * 255
    Dim val As Long
    Dim TempFileNum As Integer
    Dim var1 As String, var2 As String

    ' Create a temporary .htm file
    BrowserProg = Space(255)
    TempFileName = "C:\we.htm"
    TempFileNum = FreeFile
    Open TempFileName For Output As TempFileNum
        Print #TempFileNum, "<HTML> <\HTML>"
    Close TempFileNum

    ' Find the default browser program
    val = FindExecutable(TempFileName, vbNullString, BrowserProg)
    BrowserProg = Trim(BrowserProg)
    
    ' If the browser program is found then execute it
    If val <= 32 Or IsEmpty(BrowserProg) Then
        MsgBox "Could not open your browser!", vbExclamation, "Browser Not Found..."
    Else
        If OpenNew = True Then 'open webpage in a new browser window
            var1 = BrowserProg
            var2 = WebSiteStr
        Else                   'open webpage in existing browser window
            var1 = WebSiteStr
            var2 = vbNullString
        End If
        val = ShellExecute(&O0, "Open", var1, var2, vbNullString, SW_SHOWNORMAL)
        If val <= 32 Then
            MsgBox "Web page fialed to be opened.", vbExclamation, "Website URL Failed..."
        End If
    End If
    Kill TempFileName
End Sub

