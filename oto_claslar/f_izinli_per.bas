Attribute VB_Name = "f_izinli_per"
Function f_izinli(ByRef tarih As String, bolum As Integer) As Variant


    Dim conn As New Connection
    Dim rs As New Recordset
    
    Dim izinliler() As Variant
    Dim izinli_id() As Variant

    Dim masterdb_path As String
    
    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    'Toplantý tarihinde izinli olanlarý çekiyoruz
    qry = "select * from S_Izinler where (Bolumler_Id=" & bolum & " and BasTar<=" & tarih & " and BitisTar>" & tarih & ")"
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    izinli_sayisi = rs.RecordCount

    If izinli_sayisi <> 0 Then
        izinliler = rs.GetRows
        
        ReDim izinli_id(0 To izinli_sayisi - 1, 0 To 1)
        
        For i = 0 To izinli_sayisi - 1
            j = 0
            izinli_id(i, j) = izinliler(1, i)
            izinli_id(i, j + 1) = izinliler(4, i)
        Next i
        
    End If

    f_izinli = izinli_id
    
    rs.Close
    conn.Close
End Function



