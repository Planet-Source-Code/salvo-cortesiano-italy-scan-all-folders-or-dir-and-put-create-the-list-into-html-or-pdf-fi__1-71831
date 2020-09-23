VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MakeHTML-List v1.0.0"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotPapage 
      Height          =   315
      Left            =   3555
      TabIndex        =   8
      Text            =   "59"
      Top             =   750
      Width           =   420
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Create PDF File"
      Height          =   270
      Left            =   90
      TabIndex        =   6
      ToolTipText     =   "Only display the work"
      Top             =   765
      Value           =   1  'Checked
      Width           =   2025
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dispaly the work (Not recommended)..."
      Height          =   270
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Only display the work"
      Top             =   495
      Value           =   1  'Checked
      Width           =   4515
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   525
      TabIndex        =   1
      Text            =   "F:\Temp\Zero"
      Top             =   105
      Width           =   6225
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan Directory"
      Height          =   360
      Left            =   4830
      TabIndex        =   0
      Top             =   450
      Width           =   1920
   End
   Begin VB.Label Label2 
      Caption         =   "Line x Page:"
      Height          =   225
      Left            =   2235
      TabIndex        =   7
      Top             =   795
      Width           =   1320
   End
   Begin VB.Label lblFile 
      Caption         =   "#"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   1380
      Width           =   6690
   End
   Begin VB.Label lblPath 
      Caption         =   "#"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   1080
      Width           =   6690
   End
   Begin VB.Label Label1 
      Caption         =   "Dir:"
      Height          =   225
      Left            =   45
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' .... Init PDF Class
Private myPDF As New clsPDF
Private totPDFPage As Integer

' .... Limit Line for PDF Page
Private strLine As Integer      ' .... Default = 59

' .... STOP the recursive Scan Dir
Private STOP_PRESSED As Boolean

' .... to Strip a Dir
Private Const gstrNULL$ = ""
Private Const gstrSpace$ = " "
Private Const gstrSEP_DIR$ = "\"
Private Const gstrSEP_DIRALT$ = "/"
Private Const gstrSEP_EXT$ = "."
Private Const gstrCOLON$ = ":"
Private Const gstrSwitchPrefix1 = "-"
Private Const gstrSwitchPrefix2 = "/"
Private Const gstrCOMMA$ = ","

' .... Function Shell files
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' .... Constant
Private Const SW_SHOWNORMAL = 1
Private Sub AddItem2Array1D(ByRef VarArray As Variant, ByVal VarValue As Variant)
  Dim i  As Long
  Dim iVarType As Integer
  On Error Resume Next
  DoEvents
  iVarType = VarType(VarArray) - 8192
  i = UBound(VarArray)
  Select Case iVarType
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
      If VarArray(0) = 0 Then
        i = 0
      Else
        i = i + 1
      End If
    Case vbDate
      If VarArray(0) = "00:00:00" Then
        i = 0
      Else
        i = i + 1
      End If
    Case vbString
      If VarArray(0) = vbNullString Then
        i = 0
      Else
        i = i + 1
      End If
    Case vbBoolean
      If VarArray(0) = False Then
        i = 0
      Else
        i = i + 1
      End If
    Case Else
  End Select
  ReDim Preserve VarArray(i)
  VarArray(i) = VarValue
  DoEvents
End Sub

Private Function AllFilesInFolders(ByVal sFolderPath As String, Optional bWithSubFolders As _
                                            Boolean = True, Optional strFlag As String = "*.*") As String()
    Dim sTemp As String
    Dim sDirIn As String
    ReDim sFilelist(0) As String
    ReDim sSubFolderList(0) As String
    ReDim sToProcessFolderList(0) As String
    Dim i As Integer, j As Integer
    sDirIn = sFolderPath
    If Not (Right$(sDirIn, 1) = "\") Then sDirIn = sDirIn & "\"
    On Local Error Resume Next
    sTemp = Dir$(sDirIn & strFlag)
    Do While sTemp <> ""
      AddItem2Array1D sFilelist(), sDirIn & sTemp
      sTemp = Dir
      DoEvents
    Loop
    If bWithSubFolders Then
      sTemp = Dir$(sDirIn & strFlag, vbDirectory)
      Do While sTemp <> ""
      DoEvents
         If sTemp <> "." And sTemp <> ".." Then
            If (GetAttr(sDirIn & sTemp) And vbDirectory) = vbDirectory Then
              AddItem2Array1D sToProcessFolderList, sDirIn & sTemp
            End If
         End If
         sTemp = Dir
         DoEvents
         If STOP_PRESSED = True Then Exit Do
      Loop
      If UBound(sToProcessFolderList) > 0 Or UBound(sToProcessFolderList) = 0 And sToProcessFolderList(0) <> "" Then
        For i = 0 To UBound(sToProcessFolderList)
          DoEvents
          sSubFolderList = AllFilesInFolders(sToProcessFolderList(i), bWithSubFolders)
          If UBound(sSubFolderList) > 0 Or UBound(sSubFolderList) = 0 And sSubFolderList(0) <> "" Then
            For j = 0 To UBound(sSubFolderList)
              AddItem2Array1D sFilelist(), sSubFolderList(j)
              DoEvents
              If STOP_PRESSED = True Then Exit For
            Next
          End If
          DoEvents
          If STOP_PRESSED = True Then Exit For
        Next
      End If
    End If
    AllFilesInFolders = sFilelist
    DoEvents
Exit Function
End Function

Private Sub ScanFolder(strPath As String, includeSubDirectory As Boolean, HTMLTitle As String, _
                Optional sTitleOfFile As String = "Dir_Scan", Optional fFlag As String = "*.*", Optional MakeAsPDFfile As Boolean = False)
    ' .... Declarations Dirs and Files
    Dim lFLCount As Long
    Dim j As Integer
    Dim fOlder As Integer
    Dim totF As Integer
    Dim sDirOld As String
    Dim ix As Integer
    
    ' .... Declarations HTML
    Dim FL As Integer
    Dim i As Long
    Dim X As Long
    Dim strHead As String
    Dim strList As String
    Dim strBody As String
    Dim strEnd As String
    Dim strEndData As String
    
    Dim arrID() As Variant
    Dim StrKey As String
    
    Dim strTip As String
    Dim strTipFormatted As String
    
    Dim htmlFile As String
    Dim PDFileName As String
    
    On Local Error GoTo ErrorHandler
    
    ReDim FileList(0) As String
    
    ix = 0
    strLine = 0
    
    htmlFile = App.Path + "\" + App.Title + "_tmp.html"
    PDFileName = App.Path + "\" + App.Title + ".pdf"
    
    ' .... Dispaly Info?
    If Check1.Value = 1 Then
        lblFile.Caption = "pre-analysis... please wait!"
        lblPath.Caption = strPath
    End If
    
    '.... Make the HTML Strings
    strHead = "<html>"
    strHead = strHead & "<head>"
    strHead = strHead & "<meta http-equiv=""Content-Language"" content=""it"">"
    strHead = strHead & "<meta http-equiv=""Content-Type"" content=""text/html; charset=""windows-1252"">"
    strHead = strHead & "<meta name=""GENERATOR"" content="" & App.EXEName  & "">"
    strHead = strHead & "<meta name=""ProgId"" content="" & App.EXEName & "".Editor.Document>"
    
    strHead = strHead & "<html><head><title>" & HTMLTitle & "</title></head>"
    strHead = strHead & "<body link=#000080 vlink=#000080 alink=#000080>"
    strHead = strHead & "<p align=center><b><font face=courier new size=2 color=#000080><u>"
    strHead = strHead & "<a name=top></a>" & HTMLTitle & "</u></font><font color=#000080 face=courier new size=2><br></font></b>"
    strHead = strHead & "<font color=#000080 face=courier new size=2>Result of Scan Folder: </font><b><font color=#FF0000 face=courier new size=2>"
    strHead = strHead & strPath & "</font></b><br><br>"
    strHead = strHead & "<div><font face=courier new size=2 color=#000080><table border=0 width=100% cellspacing=1 cellpadding=1>"
    ' .... END ONE
    
    ' .... Start recursive Dirs and Files
    FileList = AllFilesInFolders(strPath, includeSubDirectory, fFlag)
    lFLCount = UBound(FileList)
    
    If MakeAsPDFfile Then
        ' .... Make OutPut FileName as PDF File
        myPDF.PaperSize = pdfA4
        myPDF.FileName = PDFileName
        myPDF.StartPDF
        ' ....
    End If
        
        For j = 0 To UBound(FileList)
            If FileList(j) <> "" Then
                DoEvents
                ' .... Count Line for New PDF Page
                strLine = strLine + 1
                ' .... Write only one Dir
                If StripChar(FileList(j), "\") <> sDirOld Then
                    fOlder = fOlder + 1
                    ' .... Make the link of HTML
                    strList = strList & "<tr><td bgcolor=#FFFFCC><b><a href=#" & fOlder & ">[#" & fOlder & "] " & StripChar(FileList(j), "\") & "</a></b></td></tr>"
                    strBody = strBody & "<tr><td width=100% bgcolor=#000080><b><font color=#FFFFFF><a name=" & fOlder & "></a>" & fOlder & "-" & StripChar(FileList(j), "\") & "</font></b></td></tr>"
                    strBody = strBody & "<tr><td width=100%><p align=right><a href=#top><font face=courier new color=#000080 size=1>Back to Top</font></a></td></tr>"
                    
                    strBody = strBody & "<tr><td width=100% bgcolor=#000080><b><font color=#FFFFFF><a name=#" & "[#" & fOlder & "] " & StripChar(FileList(j), "\") & "</a></font></b></td></tr>"
                    strBody = strBody & "<tr><td width=100%><p align=right><a href=#" & "[#" & fOlder & "] " & StripChar(FileList(j), "\") & "><font face=courier new color=#000080 size=1>Back to Folder</font></a></td></tr>"
                    
                    strBody = strBody & "<tr><td width=100%><font face=courier new size=2>"
                    
                    ' .... Write PDF Line
                    If MakeAsPDFfile Then
                        myPDF.FontSize = 6
                        Call MakeToPDFile("[#" & fOlder & "] " & StripChar(FileList(j), "\"), pdfBold)
                    End If
                    ' ....
                    
                    If Check1.Value = 1 Then lblPath.Caption = StripChar(FileList(j), "\")
                End If
                ix = ix + 1
                strBody = strBody & StripDirectory(FileList(j)) & "<span style=background-color:#FFFFCC><font color=#FF0000> " & ix & "</font></span><br>"
                sDirOld = StripChar(FileList(j), "\")
                
                ' .... Write PDF Line
                If MakeAsPDFfile Then
                    myPDF.FontSize = 5
                    Call MakeToPDFile(StripDirectory(FileList(j)) & " (" & ix & ")", pdfRegular)
                End If
                ' ....
                
            End If
            DoEvents
            
            ' .... Start New PDF Page
            If strLine >= txtTotPapage.Text Then strLine = 0
            
            If Check1.Value = 1 Then lblFile.Caption = StripDirectory(FileList(j))
            If STOP_PRESSED = True Then Exit For
        Next
               
    '.... Write the List of Dir
    strList = strList & "</table></div><p>&nbsp;</p><div><font face=courier new size=1><table border=0 cellpadding=1 cellspacing=1 width=100%>"
    
    ' .... Create the Final part with my personal Data
    strEnd = "</tr></table></div><p>&nbsp;</p><p><font face=courier new size=1>E-mail: <a href=mailto:salvocortesiano@netshadows.it>Salvo Cortesiano</a><br>"
    strEnd = strEnd & "On the web: <a href=http://www.netshadows.it/>http://www.netshadows.it/</a><br><br>"
    strEnd = strEnd & "© 2008-" & Format(Now, "mmmm, dd yyyy") & " by Salvo Cortesiano All Right Reserved.</a></font></p></body></html>"
    
    strEndData = "Total Folders: <font face=courier new color=#000080 size=1>" & fOlder & "<br></font>"
    strEndData = strEndData & "Total Files: <font face=courier new color=#000080 size=1>" & ix & "<br></font>"
    
    If MakeAsPDFfile Then
        myPDF.FontSize = 5
        Call MakeToPDFile(" ", pdfRegular)
        Call MakeToPDFile("© 2008-" & Format(Now, "mmmm, dd yyyy") & " by Salvo Cortesiano All Right Reserved.", pdfBold)
        Call MakeToPDFile("http://www.netshadows.it/", pdfBold)
        Call MakeToPDFile("-------------------------", pdfRegular)
        Call MakeToPDFile("Total Folders: " & fOlder, pdfBold)
        Call MakeToPDFile("Total Files: " & ix, pdfBold)
    End If
    ' ....
    
    ' .... Write the HTML File
    FL = FreeFile
    Open App.Path + "\" + sTitleOfFile + ".html" For Output As FL
    Print #FL, strHead;
    Print #FL, strList;
    Print #FL, strBody;
    Print #FL, strEnd;
    Print #FL, strEndData;
    Close FL
    
    
    ' .... Close the PDF File
    If MakeAsPDFfile Then
        myPDF.EndPDF
        Set myPDF = Nothing
    End If
    ' .... Reset All
    cmdScan.Caption = "&Scan Directory"
    lblFile.Caption = "Total Files: " & ix
    lblPath.Caption = "Total Folders: " & fOlder
    
    ' .... Open PDF FileName?
    If MakeAsPDFfile Then
        If Dir$(PDFileName) <> "" Then
            If MsgBox("Fle PDF created! Number of Page: " & totPDFPage + 1 & vbCrLf & "Open the File PDF?", vbYesNo + vbInformation + vbDefaultButton1, "Open File") = vbYes Then
                ShellExecute 0&, vbNullString, PDFileName, vbNullString, App.Path, SW_SHOWNORMAL
            End If
        End If
    End If
    
    ' .... Open FileName?
     If Dir$(htmlFile) <> "" Then
        If MsgBox("Open the File?", vbYesNo + vbInformation + vbDefaultButton1, "Open File") = vbYes Then
                ShellExecute 0&, vbNullString, htmlFile, vbNullString, App.Path, SW_SHOWNORMAL
        End If
    End If
  Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description & vbCrLf & "Localizzato: {Sub=AllFileInFolder}", vbCritical, App.Title
    Err.Clear
End Sub

Private Function StripChar(rsFileName As String, strCaracter As String) As String
  On Error Resume Next
  Dim i As Integer
  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = strCaracter Then
      Exit For
    End If
  Next
  StripChar = Mid(rsFileName, 1, i - 1)
End Function

Private Function StripDirectory(strString As String) As String
    Dim intPos As Integer
    StripDirectory = gstrNULL
    intPos = Len(strString)
    Do While intPos > 0
        Select Case Mid$(strString, intPos, 1)
        Case gstrSEP_DIR
                StripDirectory = Mid$(strString, intPos + 1)
            Exit Do
        Case gstrSEP_DIR, gstrSEP_DIRALT
                StripDirectory = Mid$(strString, intPos + 1)
            Exit Do
        End Select
        intPos = intPos - 1
    Loop
End Function

Private Sub Check2_Click()
    If Check2.Value Then txtTotPapage.Enabled = True Else txtTotPapage.Enabled = False
End Sub

Private Sub cmdScan_Click()
    If cmdScan.Caption = "&Scan Directory" Then
        If Check2.Value And txtTotPapage.Text > 59 Then
                MsgBox "To much line: " & txtTotPapage.Text & "." & vbCrLf & "Sintax: > 1; < 60", vbExclamation, App.Title
                    txtTotPapage.SelStart = 0
                    txtTotPapage.SelLength = 2
                    txtTotPapage.SetFocus
            Exit Sub
        End If
        cmdScan.Caption = "&Stop Scan"
            STOP_PRESSED = False
            Call ScanFolder(txtDir, True, App.Title, App.Title + "_tmp", "*.*", Check2.Value)
    ElseIf cmdScan.Caption = "&Stop Scan" Then
        If MsgBox("Stop the Scan?", vbYesNo + vbInformation + vbDefaultButton1, "Stop Scan") = vbYes Then
                STOP_PRESSED = True
            cmdScan.Caption = "&Scan Directory"
        End If
    End If
End Sub

Private Sub Form_Load()
'    txtDir = App.Path
End Sub



Private Sub MakeToPDFile(strString As String, Optional strFontType As pdfFont = pdfRegular)
    On Local Error GoTo ErrorHandler
    ' .... Write the Text
    myPDF.WritePDF strString, True, strFontType
    If strLine >= txtTotPapage.Text Then
        myPDF.NewPage
        strLine = 0
        totPDFPage = totPDFPage + 1
    End If
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

