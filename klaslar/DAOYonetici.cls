VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAOYonetici"
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

Function getDekanIdOnly() As Integer

    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    qry = "select PerId from Yoneticiler where Mevki='DEKAN'"
    
    rs.Open qry, conn, adOpenKeyset, adLockOptimistic
    
    getDekanIdOnly = rs.Fields(0)
    
    rs.Close
    conn.Close
End Function

Sub add(toplanti As EFToplanti)

End Sub
