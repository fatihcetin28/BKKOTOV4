Attribute VB_Name = "mdl_toplantilar_pop"
Sub toplantiekleformudolur(frm As MSForms.UserForm)

    Dim conn As New Connection
    Dim masterdb_path As String
    masterdb_path = ThisWorkbook.path & "\db\master.accdb"
    conn_yolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & masterdb_path
    conn.Open conn_yolu
    
    Dim rs2 As New Recordset
    Dim qry2 As String
    Dim ks() As Variant
    Dim kayitsayisi As Integer
    
    qry2 = "SELECT Toplantilar.Id, Toplantilar.[No], Toplantilar.Tarih, Bolumler.KisaBolumAdi, Toplantilar.bolumid FROM Bolumler INNER JOIN Toplantilar ON Bolumler.[Id] = Toplantilar.[bolumid] where bolumid=" & frm.cbx_bolumler.value & " order by Tarih desc"

    rs2.Open qry2, conn, adOpenKeyset, adLockOptimistic
    
    kayitsayisi = rs2.RecordCount
    If kayitsayisi <> 0 Then
    ks = rs2.GetRows
    
    With frm

        .lbx_Toplantilar.Column = ks
        .lbx_Toplantilar.ColumnCount = 5
        .lbx_Toplantilar.ColumnWidths = "0;35;55;40;0"
        .txt_toplantiSayi = .lbx_Toplantilar.List(0, 1) + 1
        .txt_toplantiSayi.Enabled = False
        
        .Repaint
        .frame_Toplantilar.Repaint
        .frame_toplantiEkle.Repaint
        .frame_toplantiEkle.Visible = False
        .frame_Toplantilar.Visible = False
        
        .frame_toplantiEkle.Visible = True
        .frame_Toplantilar.Visible = True
    End With
    Else
    With frm
        .lbl_toplantiTarih.Caption = ""
        .lbl_toplantiSayi.Caption = ""
        
        .lbx_Toplantilar.Clear
        .lbx_Toplantilar.ColumnCount = 5
        .txt_toplantiSayi = 1
        .txt_toplantiSayi.Enabled = False
            
        .Repaint
        .frame_Toplantilar.Repaint
        .frame_toplantiEkle.Repaint
        .frame_toplantiEkle.Visible = False
        .frame_Toplantilar.Visible = False
        
        .frame_toplantiEkle.Visible = True
        .frame_Toplantilar.Visible = True
    
        Dim ctl As Control
        For Each ctl In .Controls
            If ctl.name = "lbx_topKararlar" Then
                .lbx_topKararlar.Clear
            End If
            If ctl.name = "txt_kararIcerik" Then
                .txt_kararIcerik = ""
            End If
            If ctl.name = "lbx_Ekler" Then
                .lbx_Ekler.Clear
            End If
        Next ctl

    End With
    End If
    rs2.Close
    conn.Close
End Sub
