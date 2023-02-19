VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EFToplanti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
    id            As Long
    tarih         As Date
    bolum         As EFBolum
    Kararlar      As Collection
    katilanlar    As Collection
End Type

Private This As TState

Private bkklarPath As String


Private Sub Class_Initialize()
    bkklarPath = ThisWorkbook.path & "\BKKLAR"
    Set This.Kararlar = New Collection
    Set This.katilanlar = New Collection
    Set This.bolum = New EFBolum
End Sub

Private Sub Class_Terminate()
    Set This.Kararlar = Nothing
End Sub

Public Property Get id() As Variant
    id = This.id
End Property
Public Property Let id(ByVal value As Variant)
    This.id = value
End Property

Public Property Set Kararlar(ByVal value As Collection)
    Set This.Kararlar = value
End Property
Public Property Get Kararlar() As Collection
    Set Kararlar = This.Kararlar
End Property

Public Property Get tarih() As Date
    tarih = This.tarih
End Property
Public Property Let tarih(ByVal tarihArg As Date)
    This.tarih = tarihArg
End Property

Public Property Get bolum() As EFBolum
    Set bolum = This.bolum
End Property
Public Property Set bolum(ByVal value As EFBolum)
    Set This.bolum = value
End Property

Public Property Get katilanlar() As Collection

'Dim akademikPersonelManager As New mAkademikPersonelManager
Dim daAkademikPersonel As IPersonalDal
Set daAkademikPersonel = New daAkademikPersonel

'Set katilanlar = akademikPersonelManager.getAkademikPersonelsByBolumIdToColl(Me.bolum, True)
Set katilanlar = daAkademikPersonel.getAkademikPersonelsByBolumIdToColl(Me.bolum, True)


    
End Property

