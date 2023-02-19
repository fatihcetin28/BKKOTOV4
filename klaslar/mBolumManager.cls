VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mBolumManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    
    Dim conn As New Connection
    Dim rs As New Recordset
    Dim masterdb_path As String
    Dim conn_yolu As String
    Dim qry As String
        

Function getBolumByIdToCBolum(bolumId As Integer) As EFBolum

    Dim geciciBolum As New EFBolum
    
    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    qry = "select * from Bolumler where Id=" & bolumId
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount <> 0 Then
        geciciBolum.id = bolumId
        geciciBolum.adi = rs.Fields(2)
        geciciBolum.adiKisa = rs.Fields(3)
        geciciBolum.baskanId = rs.Fields(1)
        geciciBolum.adiProp = rs.Fields(4)
    End If
    
    rs.Close
    conn.Close
    
    Set getBolumByIdToCBolum = geciciBolum
End Function

Function getBolumByIdToDizi(bolumId As Integer) As Variant

    Dim geciciBolum() As Variant
    Dim bolumBilgileri() As Variant
    
    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    qry = "select * from Bolumler where Id=" & bolumId
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount <> 0 Then
        ReDim geciciBolum(0 To 4)
        geciciBolum(0) = bolumId
        geciciBolum(1) = rs.Fields(1)
        geciciBolum(2) = rs.Fields(2)
        geciciBolum(3) = rs.Fields(3)
        geciciBolum(4) = rs.Fields(4)
    End If
    
    rs.Close
    conn.Close
    
    getBolumByIdToDizi = geciciBolum
End Function

Sub updateBolumAdi(bolumId As Integer, yeniAd As String)

    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path

    conn.Open conn_yolu
    qry = "select * from Bolumler where Id=" & bolumId
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    Debug.Print "Bölüm Adý " & yeniAd & " olarak deðiþti."

    rs.Close
    conn.Close
End Sub

Sub updateBolumKisaAdi(bolumId As Integer, yeniKisaAd As String)

    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path

    conn.Open conn_yolu
    qry = "select * from Bolumler where Id=" & bolumId
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic

    
    rs.Close
    conn.Close
End Sub

Sub updateBolumPropAdi(bolumId As Integer, yeniPropAd As String)

    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path

    conn.Open conn_yolu
    qry = "select * from Bolumler where Id=" & bolumId
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic


    rs.Close
    conn.Close
End Sub

Sub updateBolumBaskanId(bolumId As Integer, yeniBaskanId As Integer)

    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path

    conn.Open conn_yolu
    qry = "select * from Bolumler where Id=" & bolumId
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic

    
    rs.Close
    conn.Close
End Sub

Sub addBolum(yeniBolum As EFBolum)

    If yeniBolum.id <> "" And yeniBolum.id <> 0 Then
        MsgBox "Lütfen ID kýsmýný boþ býrakýnýz"
        Exit Sub
    End If
    
End Sub

Sub deleteBolum(silinecekBolum As EFBolum)



End Sub
