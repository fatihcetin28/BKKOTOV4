VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mAkademikPersonelManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    
    Implements IPersonelService
    
    
    Dim conn As New Connection
    Dim rs As New Recordset
    Dim masterdb_path As String
    Dim conn_yolu As String
    Dim qry As String
    
    
Sub add(akademikPersonel As EFAkademikPersonel)

End Sub

Sub delete(akademikPersonel As EFAkademikPersonel)

End Sub

Sub updateAkaPerUnvan(akademikPersonel As EFAkademikPersonel, yeniUnvanId As Integer)

End Sub

Sub updateAkaPerBolum(akademikPersonel As EFAkademikPersonel, yeniBolumId As Integer)

End Sub

