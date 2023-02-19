VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBolum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
Id       As Integer
adi      As String
adiKisa  As String
adiProp  As String
baskanId As Integer
End Type

Private This As TState

Public Property Get Id() As Variant
    Id = This.Id
End Property

Public Property Let Id(ByVal vNewValue As Variant)
    This.Id = vNewValue
End Property

Public Property Get adi() As Variant
    adi = This.adi
End Property

Public Property Let adi(ByVal vNewValue As Variant)
    This.adi = vNewValue
End Property


Public Property Get adiKisa() As Variant
    adiKisa = This.adiKisa
End Property

Public Property Let adiKisa(ByVal vNewValue As Variant)
    This.adiKisa = vNewValue
End Property

Public Static Property Get adiProp() As Variant
    adiProp = This.adiProp
End Property

Public Static Property Let adiProp(ByVal vNewValue As Variant)
    This.adiProp = vNewValue
End Property

Public Property Get baskanId() As Variant
    baskanId = This.baskanId
End Property

Public Property Let baskanId(ByVal vNewValue As Variant)
    This.baskanId = vNewValue
End Property
