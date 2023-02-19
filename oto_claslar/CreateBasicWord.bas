Attribute VB_Name = "CreateBasicWord"
Sub CreateBasicWord()

Dim wdApp As Word.Application
Set wdApp = New Word.Application


dosyaadi = "dosya ismi5"

bolumAdi = Range("A2")
bolumadi2 = StrConv(bolumAdi, vbProperCase)
toplantisayi = Range("B2")
toplantitarih = Range("C2")
gundem1 = Range("D2")
kararbaslik1 = Range("E2")
kararicerik1 = Range("F2")
Katilan1 = Range("G2")
katilan2 = Range("G3")
katilan3 = Range("G4")
katilan4 = Range("G5")
toplantigun = Range("H2")
bolumbaskani = Katilan1


With wdApp
    .Visible = True
    .Activate

    .documents.Add ThisWorkbook.path & "\sablonlar\BKK_Sablon2.dotx"
    
    .Selection.GoTo wdGoToBookmark, , , "bolumadi"
    .Selection.TypeText bolumAdi

    .Selection.GoTo wdGoToBookmark, , , "bolumadi2"
    .Selection.TypeText bolumadi2
    
    .Selection.GoTo wdGoToBookmark, , , "toplantisayi"
    .Selection.TypeText toplantisayi

    .Selection.GoTo wdGoToBookmark, , , "toplantitarih"
    .Selection.TypeText toplantitarih
    
    .Selection.GoTo wdGoToBookmark, , , "toplantigunu"
    .Selection.TypeText toplantigun
    
    .Selection.GoTo wdGoToBookmark, , , "gundem1"
    .Selection.TypeText gundem1
    
    .Selection.GoTo wdGoToBookmark, , , "kararBaslik"
    .Selection.TypeText kararbaslik1
    
    .Selection.GoTo wdGoToBookmark, , , "karar1icerik"
    .Selection.TypeText kararicerik1
    
    .Selection.GoTo wdGoToBookmark, , , "katilan1"
    .Selection.TypeText Katilan1
    
    .Selection.GoTo wdGoToBookmark, , , "katilan2"
    .Selection.TypeText katilan2
    
    .Selection.GoTo wdGoToBookmark, , , "katilan3"
    .Selection.TypeText katilan3
    
    .Selection.GoTo wdGoToBookmark, , , "katilan4"
    .Selection.TypeText katilan4
    
    .Selection.GoTo wdGoToBookmark, , , "bolumbaskani"
    .Selection.TypeText bolumbaskani

    
    .ActiveDocument.SaveAs2 ThisWorkbook.path & "\" & dosyaadi & ".docx"
    .ActiveDocument.Close
    .Quit
    
End With

Set wdApp = Nothing

End Sub


'    With .Selection
'        .ParagraphFormat.Alignment = wdAlignParagraphCenter
'        .BoldRun
'        .Font.Size = 18
'        .TypeText "Deneme yazýsý"
'        .TypeParagraph
'        .Font.Size = 12
'        .ParagraphFormat.Alignment = wdAlignParagraphLeft
'    End With
    
'    Range("A1", Range("A2").End(xlDown).End(xlToRight)).Copy
'
'    .Selection.Paste

