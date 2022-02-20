'
' MS PowerPoint export automator - .pdf
' J. Jang
' https://github.com/jangcom/trivia
'
' References
' [1] https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat
' [2] https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
' [3] https://stackoverflow.com/questions/47656632/using-vba-how-can-i-get-the-immediate-parent-folder-name-from-a-path-string
'


Public Sub ExportPresentationAsPDF()
    ' Export a deck of slides as a .pdf file to the CWD.
    Dim presentationOutPath, presentationBname, outPathOutBname, pdfFullName
    presentationOutPath = ActivePresentation.Path
    presentationBname = Left(ActivePresentation.Name, (InStrRev(ActivePresentation.Name, ".", -1, vbTextCompare) - 1))
    outPathOutBname = presentationOutPath & "\" & presentationBname
    pdfFullName = outPathOutBname & ".pdf"

    ActivePresentation.ExportAsFixedFormat _
        Path:=pdfFullName, _
        FixedFormatType:=ppFixedFormatTypePDF, _
        Intent:=ppFixedFormatIntentPrint

    Dim PDFMsg, Buttons, Title
    PDFMsg = _
        "PDF exported as [" & pdfFullName & "]"
    Buttons = vbOKOnly + vbInformation
    Title = "J. Jang: ExportPresentationAsPDF()"
    Response = MsgBox(PDFMsg, Buttons, Title)
End Sub


Public Sub ExportPresentationAsPDFToParent()
    ' Export a deck of slides as a .pdf file to the upper directory.
    Dim presentationOutPath, presentationBname, outPathOutBname, pdfFullName
    presentationOutPath = Left(ActivePresentation.Path, InStrRev(ActivePresentation.Path, "\") - 1)
    presentationBname = Left(ActivePresentation.Name, (InStrRev(ActivePresentation.Name, ".", -1, vbTextCompare) - 1))
    outPathOutBname = presentationOutPath & "\" & presentationBname
    pdfFullName = outPathOutBname & ".pdf"

    ActivePresentation.ExportAsFixedFormat _
        Path:=pdfFullName, _
        FixedFormatType:=ppFixedFormatTypePDF, _
        Intent:=ppFixedFormatIntentPrint

    Dim PDFMsg, Buttons, Title
    PDFMsg = _
        "PDF exported as [" & pdfFullName & "]"
    Buttons = vbOKOnly + vbInformation
    Title = "J. Jang: ExportPresentationAsPDFToParent()"
    Response = MsgBox(PDFMsg, Buttons, Title)
End Sub
