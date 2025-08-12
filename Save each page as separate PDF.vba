Sub Save_As_Separate_PDFs()

    Dim I As Long
    Dim xDlg As FileDialog
    Dim xFolder As Variant
    Dim xStart, xEnd As Integer
    On Error GoTo lbl
    Set xDlg = Application.FileDialog(msoFileDialogFolderPicker)
    If xDlg.Show <> -1 Then Exit Sub
    xFolder = xDlg.SelectedItems(1)
    xStart = CInt(InputBox("Page number:", "Start from page number"))
    xEnd = CInt(InputBox("Page number:", "End to page number"))
    If xStart <= xEnd Then
        For I = xStart To xEnd
            ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                xFolder & "\ID_" & I & ".pdf", ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
                wdExportFromTo, From:=I, To:=I, Item:=wdExportDocumentContent, _
                IncludeDocProps:=False, KeepIRM:=False, CreateBookmarks:= _
                wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=False, UseISO19005_1:=False
        Next
    End If
    Exit Sub
lbl:
    MsgBox "Insert the correct page number", vbInformation, "Error message"
End Sub


