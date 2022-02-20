'
' MS Word export automator - .pdf
' J. Jang
' https://github.com/jangcom/trivia
'
' References
' [1] https://docs.microsoft.com/en-us/office/vba/api/word.document
' [2] https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
' [3] https://stackoverflow.com/questions/43345109/save-excel-as-pdf-in-current-folder-using-current-workbook-name
'


Sub ExportDocumentAsPDF()
    ' Export a Word document as a .pdf file to the CWD.
    Dim documentBname As String
    documentBname = Left(ActiveDocument.Name, (InStrRev(ActiveDocument.Name, ".", -1, vbTextCompare) - 1))

    ActiveDocument.ExportAsFixedFormat _
        OutputFileName:=ActiveDocument.Path & "\" & documentBname & ".pdf", _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, _
        IncludeDocProps:=False, _
        DocStructureTags:=False
End Sub


Sub ExportDocumentAsPDFToParent()
    ' Export a Word document as a .pdf file to the upper directory.
    Dim documentOutPath, documentBname
    documentBname = Left(ActiveDocument.Name, (InStrRev(ActiveDocument.Name, ".", -1, vbTextCompare) - 1))
    documentOutPath = Left(ActiveDocument.Path, InStrRev(ActiveDocument.Path, "\") - 1)

    outPathOutBname = documentOutPath & "\" & documentBname
    pdfFullName = outPathOutBname & ".pdf"

    ActiveDocument.ExportAsFixedFormat _
        OutputFileName:=pdfFullName, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, _
        IncludeDocProps:=False, _
        DocStructureTags:=False
End Sub
