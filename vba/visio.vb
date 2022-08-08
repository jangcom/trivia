'
' MS Visio export automator - .pdf, .emf, .png, and .jpg, or a combination of them
' J. Jang
' https://github.com/jangcom/trivia
'
' References
' [1] https://docs.microsoft.com/en-us/office/vba/api/visio.document.exportasfixedformat
' [2] https://docs.microsoft.com/en-us/office/vba/api/visio.page.export
' [3] https://docs.microsoft.com/en-us/office/vba/api/visio.applicationsettings
' [4] https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
'


Public Function GetFilePath() As String
    ' Return the document path.
    GetFilePath = ThisDocument.Path & ".."
End Function


Public Function GetFileBname() As String
    ' Return the document basename (filename without its extension).
    Dim fileName As String
    fileName = ThisDocument.Name
    GetFileBname = Left(fileName, InStrRev(fileName, ".") - 1)
End Function


Public Function GetPageName() As String
    ' Return the page name.
    GetPageName = ActivePage.Name
End Function


Public Function GetOutfileBname() As String
    ' Return outfile basename (filename without its extension).
    Dim filePath, fileBname, pageName
    filePath = GetFilePath()
    fileBname = GetFileBname()
    pageName = GetPageName()
    GetOutfileBname = filePath & "\" & fileBname & "_" & pageName
End Function


Public Function GetUserBool(Msg) As Boolean
    Dim Buttons, Title, Response
    Buttons = vbYesNo + vbQuestion + vbDefaultButton1
    Title = "J. Jang: GetUserBool()"
    Response = MsgBox(Msg, Buttons, Title)
    If Response = vbYes Then
        GetUserBool = True
    Else
        GetUserBool = False
    End If
End Function


Public Sub ExportPageAsPDF()
    Dim outfileBname As String
    outfileBname = GetOutfileBname()

    ActiveDocument.ExportAsFixedFormat _
        FixedFormat:=visFixedFormatPDF, _
        OutputFileName:=outfileBname & ".pdf", _
        Intent:=visDocExIntentPrint, _
        PrintRange:=visPrintCurrentPage, _
        FromPage:=1, _
        ToPage:=1, _
        ColorAsBlack:=False, _
        IncludeBackground:=False, _
        IncludeDocumentProperties:=False, _
        IncludeStructureTags:=False

    PDFMsg = _
        "PDF export completed; consider running" _
        & Chr(10) & "(a) Ghostscript for .pdf font embedment, and" _
        & Chr(10) & "(b) pdfcrop.pl for .pdf margin removal."
    Buttons = vbOKOnly + vbInformation
    Title = "J. Jang: ExportPageAsPDF()"
    Response = MsgBox(PDFMsg, Buttons, Title)
End Sub


Public Sub ExportPageAsEMF()
    Dim outfileBname As String
    outfileBname = GetOutfileBname()

    ActivePage.Export outfileBname & ".emf"
End Sub


Public Sub ExportPageAsSVG()
    Dim outfileBname As String
    outfileBname = GetOutfileBname()

    Application.Settings.SVGExportFormat = visSVGExcludeVisioElements
    ActivePage.Export outfileBname & ".svg"
End Sub


Public Sub ExportPageAsPNG()
    Dim outfileBname As String
    Dim vsoApplicationSettings As Visio.ApplicationSettings
    outfileBname = GetOutfileBname()
    Set vsoApplicationSettings = Visio.Application.Settings

    vsoApplicationSettings.SetRasterExportResolution visRasterUseCustomResolution, 600#, 600#, visRasterPixelsPerInch
    vsoApplicationSettings.SetRasterExportSize visRasterFitToSourceSize
    vsoApplicationSettings.RasterExportDataFormat = visRasterInterlace
    vsoApplicationSettings.RasterExportColorFormat = visRaster24Bit
    vsoApplicationSettings.RasterExportRotation = visRasterNoRotation
    vsoApplicationSettings.RasterExportFlip = visRasterNoFlip
    vsoApplicationSettings.RasterExportBackgroundColor = 16777215
    vsoApplicationSettings.RasterExportTransparencyColor = 16777215
    vsoApplicationSettings.RasterExportUseTransparencyColor = True
    ActivePage.Export outfileBname & ".png"
End Sub


Public Sub ExportPageAsJPG()
    Dim outfileBname As String
    Dim vsoApplicationSettings As Visio.ApplicationSettings
    outfileBname = GetOutfileBname()
    Set vsoApplicationSettings = Visio.Application.Settings

    vsoApplicationSettings.SetRasterExportResolution visRasterUseCustomResolution, 600#, 600#, visRasterPixelsPerInch
    vsoApplicationSettings.SetRasterExportSize visRasterFitToSourceSize
    vsoApplicationSettings.RasterExportColorFormat = visRasterRGB
    vsoApplicationSettings.RasterExportOperation = visRasterBaseline
    vsoApplicationSettings.RasterExportRotation = visRasterNoRotation
    vsoApplicationSettings.RasterExportFlip = visRasterNoFlip
    vsoApplicationSettings.RasterExportBackgroundColor = 16777215
    vsoApplicationSettings.RasterExportQuality = 100
    ActivePage.Export outfileBname & ".jpg"
End Sub


Public Sub ExportPageAsImageFiles()
    Dim userBool As Boolean
    userBoolMsg = _
        "Export the page [" & ActivePage.Name & "] to image files?" _
        & " Output path:" & Chr(10) & "[" & GetFilePath() & "]"
    userBool = GetUserBool(userBoolMsg)

    TrueMsg = _
        ".pdf, .emf, .png, and .jpg" & _
        " exported to: [" & GetFilePath() & "]"
    Buttons = vbOKOnly + vbInformation
    Title = "J. Jang: ExportPageAsImageFiles()"

    If userBool = True Then
        Call ExportPageAsPDF()
        Call ExportPageAsEMF()
        Call ExportPageAsSVG()
        Call ExportPageAsPNG()
        Call ExportPageAsJPG()
        Response = MsgBox(TrueMsg, Buttons, Title)
    End If
End Sub
