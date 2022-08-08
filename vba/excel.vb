'
' MS Excel export automator - .pdf
' J. Jang
' https://github.com/jangcom/trivia
'
' References
' [1] https://trumpexcel.com/excel-macro-examples/#tab-con-10
' [2] https://stackoverflow.com/questions/43345109/save-excel-as-pdf-in-current-folder-using-current-workbook-name
'


Sub ExportWorkbookAsPDF()
    ' Export an Excel workbook as a .pdf file to the CWD.
    Dim workbookOutPath, workbookBname, outPathOutBname, pdfFullName
    workbookOutPath = ThisWorkbook.Path
    WorkbookBname = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
    outPathOutBname = workbookOutPath & "\" & WorkbookBname
    pdfFullName = outPathOutBname & ".pdf"

    ThisWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfFullName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub
