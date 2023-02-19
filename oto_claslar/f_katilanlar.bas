Attribute VB_Name = "f_katilanlar"
'' bolum verip sýralý katýlanlar koleksiyonu alýyoruz

Function func_arr_katilanlar(bolum As Integer) As Variant

    Dim dict_katilanlar As Scripting.Dictionary
    Set dict_katilanlar = New Scripting.Dictionary
    Dim conn As New Connection
    Dim rs As New Recordset
    Dim rs2 As New Recordset
        
    Dim masterdb_path As String
    
    Dim ksKatilan As Variant
    Dim ksBolumler As Variant
    Dim func_arr_katilanlar_g As Variant
    
    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    'S_Katilan sorgusundan bölüme göre veri çekiyoruz, ünvana göre sýralý veri ve katýlma hakký olanlar var
    qry = "select * from S_Katilan where Bolum=" & bolum
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    'Bölüm Baþkanýnýn Id sini çekiyoruz, bölüme göre,,, '' bölüm adlarýný da bu sorguyla çekiyoruz
    qry2 = "select Bolumler.Baskan, Bolumler.BolumAdi, Bolumler.BolumAdiProp from Bolumler where Id=" & bolum
    rs2.Open qry2, conn, adOpenKeyset, adLockOptimistic
    
    ksKatilan = rs.GetRows
    ksBolumler = rs2.GetRows
    
    bolBaskanId = ksBolumler(0, 0)
    bolAdi = ksBolumler(1, 0)
    bolAdiProp = ksBolumler(2, 0)
    
    'bölüm baþkanýný katlýlanlar(1) e atýyoruz
   For i = 0 To UBound(ksKatilan, 2)
        If ksKatilan(0, i) = bolBaskanId Then
            dict_katilanlar.Add ksKatilan(0, i), ksKatilan(2, i) & " " & ksKatilan(1, i)
        End If
   Next i
   
   'Azize Alaylý dekan old için 2 ye atýyoruz
   For i = 0 To UBound(ksKatilan, 2)
        If ksKatilan(1, i) = "Azize ALAYLI" Then
            dict_katilanlar.Add ksKatilan(0, i), ksKatilan(2, i) & " " & ksKatilan(1, i)
        End If
   Next i
   
   'Gerisini ünvan sýrasýyla diziyoruz
   For i = 0 To UBound(ksKatilan, 2)
        If ksKatilan(0, i) <> bolBaskanId And ksKatilan(1, i) <> "Azize ALAYLI" Then
            dict_katilanlar.Add ksKatilan(0, i), ksKatilan(2, i) & " " & ksKatilan(1, i)
        End If
   Next i
    
    rs.Close
    rs2.Close
    conn.Close

    ReDim func_arr_katilanlar_g(0 To dict_katilanlar.Count - 1, 0 To 1)
    
    For j = 0 To dict_katilanlar.Count - 1
        k = 0
        func_arr_katilanlar_g(j, k) = dict_katilanlar.Keys(j)
        func_arr_katilanlar_g(j, k + 1) = dict_katilanlar.Items(j)
        
        
    Next j

    func_arr_katilanlar = func_arr_katilanlar_g
End Function

