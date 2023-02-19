VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EFBolum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
id       As Integer
adi      As String
adiKisa  As String
adiProp  As String
baskanId As Variant
End Type

Private gAdi As String
Private gAdiProp As String
Private gBaskanId As Variant
Private gAdiKisa As String

Dim bolumManager As mBolumManager

Private bolumBilgileri() As Variant

Private This As TState

Public Property Get id() As Variant
    id = This.id
End Property

Public Property Let id(ByVal vNewValue As Variant)
    This.id = vNewValue
    
    Set bolumManager = New mBolumManager
    bolumBilgileri = bolumManager.getBolumByIdToDizi(Me.id)
    
    gAdi = bolumBilgileri(1)
    gAdiKisa = bolumBilgileri(2)
    gAdiProp = bolumBilgileri(4)
    gBaskanId = bolumBilgileri(3)
    
End Property


Public Property Get adi() As Variant
    If gAdi = "" Then
        adi = This.adi
    Else
        adi = gAdi
    End If
End Property

Public Property Let adi(ByVal vNewValue As Variant)
    This.adi = vNewValue
End Property

Public Property Get adiKisa() As Variant
    If gAdiKisa = "" Then
        adiKisa = This.adiKisa
    Else
        adiKisa = gAdiKisa
    End If
End Property
Public Property Let adiKisa(ByVal vNewValue As Variant)
    This.adiKisa = vNewValue
End Property

Public Static Property Get adiProp() As Variant
    If gAdiProp = "" Then
        adiProp = This.adiProp
    Else
        adiProp = gAdiProp
    End If
End Property

Public Property Let adiProp(ByVal vNewValue As Variant)
    This.adiProp = vNewValue
End Property

Public Property Get baskanId() As Variant
    If gBaskanId = "" Then
        baskanId = This.baskanId
    Else
        baskanId = gBaskanId
    End If
End Property

Public Property Let baskanId(ByVal vNewValue As Variant)
    This.baskanId = vNewValue
End Property

