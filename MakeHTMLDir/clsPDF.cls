VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 ' File name (.PDF)
Public FileName As String

' Info - these are optional
Public Producer As String   ' Producer
Public Title As String      ' Title
Public Subject As String    ' Subject
Public Author As String     ' Author

' Paper specifications
Public PaperSize As pdfPaperSize    ' Paper Size
Public LandScape As Boolean ' If TRUE, swap height & width
Public TopMargin As Long    ' Top margin
Public LeftMargin As Long   ' Left margin
Private PaperWidth As Long  ' Paper width (computed)
Private PaperHeight As Long ' Paper height (computed)

' Font specifications
Public FontSize As Long     ' Font size
Public FontType As pdfFont  ' Type (Bold/Italic/Regular)
Public VertSpace As Long    ' Spacing between rows

' Support variables
Private nStartStream
Private nObject As Long             ' Number of Objects
Private Objects As New Collection   ' Objects collection (with offset from the start of the file)
Private Kids As New Collection      ' Pages collection

' Font types
Public Enum pdfFont
    pdfRegular = 1
    pdfItalic = 2
    pdfBold = 3
    pdfBoldItalic = 4
End Enum
    
' Paper sizes
Public Enum pdfPaperSize
    pdfA3 = 1
    pdfA4 = 2
    pdfA5 = 3
    pdfTABLOID = 4
    pdfLEDGER = 5
    pdfLEGAL = 6
    pdfSTATEMENT = 7
    pdfEXECUTIVE = 8
End Enum

' Setup default values
Private Sub Class_Initialize()
    PaperSize = pdfA4
    LandScape = False
    TopMargin = 50
    LeftMargin = 50
    FontType = pdfRegular
    FontSize = 10
    VertSpace = 12
End Sub

' Write the header
Sub StartPDF()
    Dim nLen As Long
    Dim cObj As String
    Dim nTemp As Long
    
    ' Compute width & height
    Select Case PaperSize
        Case pdfA3
            PaperWidth = 842
            PaperHeight = 1190
        Case pdfA4
            PaperWidth = 595
            PaperHeight = 842
        Case pdfA5
            PaperWidth = 421
            PaperHeight = 595
        Case pdfTABLOID
            PaperWidth = 792
            PaperHeight = 1224
        Case pdfLEDGER
            PaperWidth = 1224
            PaperHeight = 792
        Case pdfLETTER
            PaperWidth = 612
            PaperHeight = 1008
        Case pdfSTATEMENT
            PaperWidth = 396
            PaperHeight = 612
        Case pdfEXECUTIVE
            PaperWidth = 540
            PaperHeight = 720
        Case Else
            ' The default is A4
            PaperWidth = 595
            PaperHeight = 842
    End Select
    
    ' If landscape, swap width & height
    If LandScape Then
        nTemp = PaperWidth
        PaperWidth = PaperHeight
        PaperHeight = nTemp
    End If
    
    ' Create the output file
    n = FreeFile
    Open FileName For Output As #n
    
    ' Version
    Print #n, "%PDF-1.2" & vbLf & _
              "%âãÏÓ" & vbLf;
    Close #n
    
    ' Info
    nObject = 1
    cObj = "/CreationDate (D:" + Format(Now, "YYYYMMDDHHMMSS") + ")"
    cObj = cObj + vbLf + "/Creator (Visual Basic)"
    cObj = cObj + vbLf + "/Producer (vb2pdf v1.2 \251 M. Nicolato)"
    If Title <> "" Then cObj = cObj + vbLf + "/Title (" + Title + ")"
    If Subject <> "" Then cObj = cObj + vbLf + "/Subject (" + Subject + ")"
    If Author <> "" Then cObj = cObj + vbLf + "/Author (" + Author + ")"
    WriteObject cObj
    
    ' Objects 2 and 3 will be created in the trailer
    nObject = 4
    
    ' Font REGULAR
    cObj = "/Type /Font"
    cObj = cObj + vbLf + "/Subtype /Type1"
    cObj = cObj + vbLf + "/Name /F1"
    cObj = cObj + vbLf + "/Encoding 8 0 R"
    cObj = cObj + vbLf + "/BaseFont /Courier"
    WriteObject cObj
    
    ' Font ITALIC
    nObject = nObject + 1
    cObj = "/Type /Font"
    cObj = cObj + vbLf + "/Subtype /Type1"
    cObj = cObj + vbLf + "/Name /F2"
    cObj = cObj + vbLf + "/Encoding 8 0 R"
    cObj = cObj + vbLf + "/BaseFont /Courier-Oblique"
    WriteObject cObj
    
    ' Font BOLD
    nObject = nObject + 1
    cObj = "/Type /Font"
    cObj = cObj + vbLf + "/Subtype /Type1"
    cObj = cObj + vbLf + "/Name /F3"
    cObj = cObj + vbLf + "/Encoding 8 0 R"
    cObj = cObj + vbLf + "/BaseFont /Courier-Bold"
    WriteObject cObj
    
    ' Font BOLD ITALIC
    nObject = nObject + 1
    cObj = "/Type /Font"
    cObj = cObj + vbLf + "/Subtype /Type1"
    cObj = cObj + vbLf + "/Name /F4"
    cObj = cObj + vbLf + "/Encoding 8 0 R"
    cObj = cObj + vbLf + "/BaseFont /Courier-BoldOblique"
    WriteObject cObj
    
    ' Font Encoding
    nObject = nObject + 1
    cObj = "/Type /Encoding"
    cObj = cObj + vbLf + "/BaseEncoding /WinAnsiEncoding"
    WriteObject cObj
    
    ' Fonts object
    nObject = nObject + 1
    cObj = "  /Font << /F1 4 0 R /F2 5 0 R /F3 6 0 R /F4 7 0 R >>"
    cObj = cObj + vbLf + "  /ProcSet [ /PDF /Text ]"
    WriteObject cObj
    
    ' Start with a new page
    NewPage
    
End Sub

' Write the PDF terminator
Public Sub EndPDF()
    Dim cObj As String
    Dim n As Long
    Dim nOffset As Long
    
    ' Close the last page
    EndPage
    
    ' Catalog
    cObj = "2 0 obj"
    cObj = cObj + vbLf + "<<"
    cObj = cObj + vbLf + "/Type /Catalog"
    cObj = cObj + vbLf + "/Pages 3 0 R"
    cObj = cObj + vbLf + "/PageLayout /OneColumn"
    cObj = cObj + vbLf + ">>"
    cObj = cObj + vbLf + "endobj"
    nOffset = WriteText(cObj)
    Objects.Add (nOffset), CStr(2)
    
    ' Pages
    cObj = "3 0 obj"
    cObj = cObj + vbLf + "<<"
    cObj = cObj + vbLf + "/Type /Pages"
    cObj = cObj + vbLf + "/Count " + CStr(Kids.Count)
    cObj = cObj + vbLf + "/MediaBox [ 0 0 " + CStr(PaperWidth) + " " + CStr(PaperHeight) + " ]"
    cObj = cObj + vbLf + "/Kids [ "
    ' List each page object
    For n = 1 To Kids.Count
        cObj = cObj + CStr(Kids(n)) + " 0 R "
    Next
    cObj = cObj + "]"
    cObj = cObj + vbLf + ">>"
    cObj = cObj + vbLf + "endobj"
    ' Write
    nOffset = WriteText(cObj)
    Objects.Add (nOffset), CStr(3)
    
    ' CrossReference
    nObject = nObject + 1
    cObj = "xref"
    cObj = cObj + vbLf + "0 " + CStr(nObject)
    cObj = cObj + vbLf + "0000000000 65535 f "
    ' List all the object's offset, ordering by object number
    For n = 1 To Objects.Count
        cObj = cObj + vbCr + Format(Objects(CStr(n)), "0#########") + " 00000 n "
    Next
    cObj = cObj + vbCr + "trailer" ' prova
    nOffset = WriteText(cObj)
    
    ' Trailer
    cObj = "<<"
    cObj = cObj + vbLf + "/Size " + CStr(nObject)
    cObj = cObj + vbLf + "/Root 2 0 R"
    cObj = cObj + vbLf + "/Info 1 0 R"
    cObj = cObj + vbLf + ">>"
    cObj = cObj + vbLf + "startxref"
    cObj = cObj + vbLf + CStr(nOffset)
    cObj = cObj + vbLf + "%%EOF"
    WriteText cObj
End Sub

' Create a new page
Public Sub NewPage()
    Dim cObj As String
    Dim nLen As Long
    Dim nOffset As Long
    
    ' Close the previous page
    If Kids.Count > 0 Then
        EndPage
    End If
    
    ' Page resources
    nObject = nObject + 1
    cObj = "/Type /Page"
    cObj = cObj + vbLf + "/Parent 3 0 R"
    cObj = cObj + vbLf + "/Resources 9 0 R"
    cObj = cObj + vbLf + "/Contents " + CStr(nObject + 1) + " 0 R"
    WriteObject cObj
    
    ' Add this object to the pages's collection
    Kids.Add nObject
    
    ' Page length is in the next object
    nObject = nObject + 1
    cObj = CStr(nObject) + " 0 obj"
    cObj = cObj + vbLf + "<<"
    cObj = cObj + vbLf + "/Length " + CStr(nObject + 1) + " 0 R"
    cObj = cObj + vbLf + ">>"
    cObj = cObj + vbLf + "stream"
    cObj = cObj + vbLf + "BT"
    nOffset = WriteText(cObj)
    Objects.Add nOffset, CStr(nObject)
    
    ' Store the offset of the stream
    nStartStream = nOffset + Len(cObj) - 5 ' Adjust start of stream
    
    ' Start with the font, page offset and vertical spacing
    cObj = "/F" + CStr(FontType) + " " + CStr(FontSize) + " Tf"
    WriteText cObj
    
    ' Start offset and rotation at default position
    SetOrigin LeftMargin, PaperHeight - TopMargin, 0
        
End Sub

' Close the page
Private Sub EndPage()
    Dim cObj As String
    Dim nLen As Long
    Dim nOffset As Long
    cObj = "ET"
    cObj = cObj + vbLf + "endstream"
    cObj = cObj + vbLf + "endobj"
    
    ' Calculate the length of the page
    nLen = WriteText(cObj) - nStartStream
    nObject = nObject + 1
    cObj = CStr(nObject) + " 0 obj"
    cObj = cObj + vbLf + CStr(nLen)
    cObj = cObj + vbLf + "endobj"
    nOffset = WriteText(cObj)
    Objects.Add (nOffset), CStr(nObject)

End Sub

' Write the text in the PDF file
Public Sub WritePDF(ByVal cText As String, bNewRow As Boolean, Optional NewFont As pdfFont)
    Dim cObj As String
    Dim cRestoreFont As String
    
    ' Change the current font ?
    If NewFont <> 0 And NewFont <> FontType Then
        cObj = "/F" + CStr(NewFont) + " " + CStr(FontSize) + " Tf" + vbLf
        cRestoreFont = vbLf + "/F" + CStr(FontType) + " " + CStr(FontSize) + " Tf"
    End If
    
    ' New row ?
    If bNewRow Then
        cObj = cObj + "T* "
    End If
    
    ' Convert special chars
    cObj = cObj + "(" + Convert(cText) + ") Tj"
    
    ' Restore original font
    cObj = cObj + cRestoreFont
    WriteText cObj

End Sub

' Set origin & text orientation (added 2003, may)
Public Sub SetOrigin(nStartX As Long, nStartY As Long, Optional nDegree As Long = 0)
    Dim cObj As String
    Dim a As Single
    Dim b As Single
    Dim c As Single
    Dim d As Single
    
    ' calculate (Tm) matrix coefficents
    Const pi = 3.141592654
    a = Cos(pi * nDegree / 180)
    b = Sin(pi * nDegree / 180)
    c = -b
    d = a
    
    ' Tm text matrix
    WriteText Number2Str(a, 3) & " " & _
              Number2Str(b, 3) & " " & _
              Number2Str(c, 3) & " " & _
              Number2Str(d, 3) & " " & _
              Str(nStartX) & " " & _
              Str(nStartY) & _
              " Tm"
    
    ' Vertical spacing
    WriteText CStr(VertSpace) + " TL"


End Sub

' Write this text as an object using PDF syntax
Private Sub WriteObject(cObject As String)
    Dim nOffset As Long
    Dim cObj As String
    cObj = CStr(nObject) + " 0 obj"
    cObj = cObj + vbLf + "<<"
    cObj = cObj + vbLf + cObject
    cObj = cObj + vbLf + ">>"
    cObj = cObj + vbLf + "endobj"
    nOffset = WriteText(cObj)
    ' Add the offset of this objetc to the Objects collection
    Objects.Add (nOffset), CStr(nObject)
End Sub

' Write this text and return his offset from the start of the file
Private Function WriteText(cText As String) As Long
    Dim n As Long
    n = FreeFile
    Open FileName For Append As #n
    WriteText = LOF(n)
    Print #n, cText & vbLf;
    Close #n
End Function

' Convert special chars as "\", "(" or ")"
Private Function Convert(cText As String) As String
    Dim i As Long
    Dim t As String
    For i = 1 To Len(cText)
        t = Mid(cText, i, 1)
        If t = "\" Or t = "(" Or t = ")" Then
            t = "\" + t
        End If
        Convert = Convert + t
    Next
End Function

' Same as FORMAT function, but returns a DOT instead of local decimal separator
' (it may be different by country)
Private Function Number2Str(nValue As Variant, nDecimals As Long) As String
    Dim cMask As String
    ' create a mask for decimal point
    If nDecimals <> 0 Then
        cMask = "0." & String(nDecimals - 1, "#") & "0"
    Else
        cMask = "0"
    End If
    Number2Str = Format(nValue, cMask)
    ' replace local decimal separator with "."
    Dim cDecSep As String
    cDecSep = Mid(CStr(0.1), 2, 1)  ' get local decimal separator character
    Dim nDecPos As Long
    nDecPos = InStr(Number2Str, cDecSep)     ' find where used
    If nDecPos <> 0 Then
        Mid(Number2Str, nDecPos, 1) = "."    ' substitute with a DOT
    End If
End Function
