VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAOAkademikPersonel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IPersonalDal
    Dim conn As New Connection
    Dim rs As New Recordset
    Dim masterdb_path As String
    Dim conn_yolu As String
    Dim qry As String
Function IPersonalDal_getAkademikPersonelsByBolumIdToColl(bolum As EFBolum, katilanlar As Boolean) As Collection

Dim geciciDizi() As Variant
Dim geciciCol As New Collection
Dim geciciCol2 As New Collection
Dim akademikPersonel As IEFPersonel
Dim i As Integer
    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    If katilanlar = False Then
        qry = "select * from AkaPersonel where BolumId=" & bolum.id & " order by UnvanId"
    Else
        qry = "select * from AkaPersonel where BolumId=" & bolum.id & " order by UnvanId"
    End If
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    geciciDizi = rs.GetRows
    
    If rs.RecordCount <> 0 Then
        For i = 0 To rs.RecordCount - 1
            Set akademikPersonel = New EFAkademikPersonel
            akademikPersonel.id = geciciDizi(0, i)
            akademikPersonel.adi = geciciDizi(3, i)
            akademikPersonel.soyadi = geciciDizi(4, i)
            akademikPersonel.unvanId = geciciDizi(2, i)
            akademikPersonel.bolumId = geciciDizi(1, i)
            geciciCol.add akademikPersonel
        Next i
        
    End If
    
    
    Dim bolumManager As New mBolumManager
    Dim gBolum As New EFBolum
    Set gBolum = bolumManager.getBolumByIdToCBolum(bolum.id)
    
    Dim j As Integer
    For j = 1 To geciciCol.Count
        If geciciCol(j).id = gBolum.baskanId Then
            geciciCol2.add geciciCol(j)
            Exit For
        End If
    Next j
    
    Dim dekanId As Integer
    Dim DAyonetici As New DAyonetici

    dekanId = DAyonetici.getDekanIdOnly
    
    Dim k As Integer
    For k = 1 To geciciCol.Count
        If geciciCol(k).id = dekanId Then
            geciciCol2.add geciciCol(k)
            Exit For
        End If
    Next k
    Dim l As Integer
    For l = 1 To geciciCol.Count
        If geciciCol(l).id <> gBolum.baskanId And geciciCol(l).id <> dekanId Then
            geciciCol2.add geciciCol(l)
        End If
    Next l
    
    rs.Close
    conn.Close
    
    Set IPersonalDal_getAkademikPersonelsByBolumIdToColl = geciciCol2
    
End Function

Function IPersonalDal_getAkademikPersonelsByBolumIdToArr(bolum As EFBolum, katilanlar As Boolean) As Variant

Dim geciciDizi() As Variant

    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    If katilanlar = False Then
        qry = "select * from AkaPersonel where Bolum=" & bolum.id
    Else
        qry = "select * from AkaPersonel where Bolum=" & bolum.id & " and BKK_KatilH=1"
    End If
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount <> 0 Then
        geciciDizi = rs.GetRows
    End If
    
    rs.Close
    conn.Close
    
    getAkademikPersonelsByBolumIdToArr = geciciDizi
    
End Function

