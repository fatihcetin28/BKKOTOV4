Attribute VB_Name = "wordPreview"
Sub WordOnizle()
Dim WordDoc As Object

Set WordDoc = CreateObject("Word.Application")
WordDoc.Visible = True
WordDoc.documents.Open ThisWorkbook.path & "\dosya ismi.docx"

WordDoc.documents("dosya ismi.docx").PrintPreview

Set WordDoc = Nothing
End Sub
