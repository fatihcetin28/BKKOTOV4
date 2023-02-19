Attribute VB_Name = "Kat_Izin_Imza"
Sub katimzaizin()

    '''*****SUCCESSSSSS
    ''ÝZÝNLÝ MÝZÝNLÝ ÝMZA TABLOSU OLUÞTURUYORUZ

Dim katilanSayi As Integer
Dim ksKatilanSirali As Variant
Dim ksIzinliler As Variant

Dim satir_sayi As Integer
Dim izin_durum As String
''' katilan sýralý ve o tarihte izinli olanlar dizi olarak geliyor - daha hýzlý
Dim bolum As Integer
Dim tarih As String

bolum = 3
tarih = "#15/09/2022#"

dosyaadi = "izin_durum_dahil6"

    ksKatilanSirali = func_arr_katilanlar(bolum)
    ksIzinliler = f_izinli(tarih, bolum)

    katilanSayi = UBound(ksKatilanSirali, 1) - LBound(ksKatilanSirali, 1) + 1
    
    ReDim Preserve ksKatilanSirali(0 To katilanSayi - 1, 0 To 2)
    
    ''ksKatilanSirali arr nin 3. sütununa izin bilgisi atýyoruz
    For i = 0 To katilanSayi - 1
        For j = 0 To UBound(ksIzinliler)
            If ksKatilanSirali(i, 0) = ksIzinliler(j, 0) Then
                ksKatilanSirali(i, 2) = ksIzinliler(j, 1)
                Exit For
            Else
                ksKatilanSirali(i, 2) = 0
            End If
        Next j
    Next i
    
    
    If katilanSayi Mod 2 = 0 Then
        satir_sayi = katilanSayi / 2
    Else
        satir_sayi = (katilanSayi + 1) / 2
    End If
    
    '' word'e dokme tablo olarak
    Dim WordDoc As Object
    
    Set WordDoc = CreateObject("Word.Application")
    WordDoc.Visible = True
    WordDoc.documents.Open ThisWorkbook.path & "\bos.docx"
    WordDoc.Activate

    Set myRange = WordDoc.documents("bos.docx").Range(0, 0)
    WordDoc.documents("bos.docx").Tables.Add Range:=myRange, NumRows:=satir_sayi, NumColumns:=2
    WordDoc.documents("bos.docx").Tables.Item(1).Spacing = 16
    r = 1
    m = 1
    
    For k = 1 To satir_sayi
    
        For l = 1 To 2
            If r = 1 Then
            
        izin_durum = ""
        If ksKatilanSirali(m - 1, 2) = 1 Then
            izin_durum = " (Yýllýk Ýzinli)"
        ElseIf ksKatilanSirali(m - 1, 2) = 2 Then
            izin_durum = " (Raporlu)"
        ElseIf ksKatilanSirali(m - 1, 2) = 0 Then
            izin_durum = ""
        End If
        
                WordDoc.documents("bos.docx").Tables.Item(1).Cell(k, l).Range.Text = ksKatilanSirali(m - 1, 1) & _
                vbNewLine & "Bölüm Baþkaný" & izin_durum
                WordDoc.documents("bos.docx").Tables.Item(1).Cell(k, l).Range.ParagraphFormat.SpaceAfter = 0
            Else
                izin_durum = ""
        If ksKatilanSirali(m - 1, 2) = 1 Then
            izin_durum = " (Yýllýk Ýzinli)"
        ElseIf ksKatilanSirali(m - 1, 2) = 2 Then
            izin_durum = " (Raporlu)"
        ElseIf ksKatilanSirali(m - 1, 2) = 0 Then
            izin_durum = ""
        End If
                WordDoc.documents("bos.docx").Tables.Item(1).Cell(k, l).Range.Text = ksKatilanSirali(m - 1, 1) & _
                Chr(10) & "Üye" & izin_durum
                WordDoc.documents("bos.docx").Tables.Item(1).Cell(k, l).Range.ParagraphFormat.SpaceAfter = 0
            End If
            r = r + 1
            m = m + 1
        Next l
        
    Next k
    
    
    
    'WordDoc.documents("bos.docx").PrintPreview
    WordDoc.documents("bos.docx").SaveAs2 ThisWorkbook.path & "\" & dosyaadi & ".docx"
    WordDoc.Quit
    


Set WordDoc = Nothing
    
    '''*****SUCCESSSSSS
    
End Sub
