VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EFYonetici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TState
    id            As Long
    mevki         As String
    perId         As Integer
End Type

Private This As TState


Public Property Get id() As Variant
    id = This.id
End Property

Public Property Let id(ByVal vNewValue As Variant)
    This.id = vNewValue
End Property

Public Property Get mevki() As Variant
    mevki = This.mevki
End Property

Public Property Let mevki(ByVal vNewValue As Variant)
    This.mevki = vNewValue
End Property


Public Property Get perId() As Variant
    perId = This.perId
End Property

Public Property Let perId(ByVal vNewValue As Variant)
    This.perId = vNewValue
End Property
