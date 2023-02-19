VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKarar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pBaslik As String
Private pIcerik As String
Private pTopId As String
Private bkklarPath As String

Private Sub Class_Initialize()
bkklarPath = ThisWorkbook.path & "\BKKLAR"
End Sub


Public Property Get Id() As Variant
    Id = pId
End Property
Public Property Let Id(ByVal IdArg As Variant)
    pId = IdArg
End Property

Public Property Get baslik() As Variant
    baslik = pBaslik
End Property
Public Property Let baslik(ByVal baslikArg As Variant)
    pBaslik = baslikArg
End Property

Public Property Get icerik() As Variant
    icerik = pIcerik
End Property
Public Property Let icerik(ByVal icerikArg As Variant)
    pIcerik = icerikArg
End Property

Public Property Get topId() As Variant
    topId = pTopId
End Property
Public Property Let topId(ByVal topIdArg As Variant)
    pTopId = topIdArg
End Property

