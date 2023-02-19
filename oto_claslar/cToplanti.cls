VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cToplanti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
    Id            As Long
    tarih         As Date
    bolum         As cBolum
    Kararlar      As Collection
End Type




Private pId As String
Private ptarih As String
Private pbolumId As String
Private pbolumAdi As String
Private bkklarPath As String

Private This As TState

Public Property Get bolum() As cBolum
    Set bolum = This.bolum
End Property
Public Property Set bolum(ByVal value As cBolum)
    Set This.bolum = value
End Property

Private Sub Class_Initialize()
bkklarPath = ThisWorkbook.path & "\BKKLAR"
Set This.Kararlar = New Collection
Set This.bolum = New cBolum
End Sub

Private Sub Class_Terminate()
Set This.Kararlar = Nothing
End Sub

Public Property Get Id() As Variant
    Id = This.Id
End Property
Public Property Let Id(ByVal value As Variant)
    This.Id = value
End Property

Public Property Set Kararlar(ByVal value As Collection)
    Set This.Kararlar = value
End Property
Public Property Get Kararlar() As Collection
    Set Kararlar = This.Kararlar
End Property

Public Property Get tarih() As Variant
    tarih = ptarih
End Property
Public Property Let tarih(ByVal tarihArg As Variant)
    ptarih = tarihArg
End Property

'Public Property Get bolumId() As Variant
'    bolumId = pbolumId
'End Property
'Public Property Let bolumId(ByVal bolumIdArg As Variant)
'    pbolumId = bolumIdArg
'End Property
'
'Public Property Let bolumAdi(ByVal bolumAdiArg As Variant)
'    pbolumAdi = bolumAdiArg
'End Property
'Public Property Get bolumAdi() As Variant
'    bolumAdi = pbolumAdi
'End Property
'
'
'Public Property Get path() As Variant
'    path = bkklarPath & "\" & Me.bolumAdi & "\" & Me.tarih
'End Property





