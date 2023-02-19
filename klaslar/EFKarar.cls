VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EFKarar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
    id       As Integer
    baslik   As String
    icerik   As String
    topId    As String
    ekler    As Collection
End Type

Private This As TState

Private bkklarPath As String

Private Sub Class_Initialize()
    bkklarPath = ThisWorkbook.path & "\BKKLAR"
    Set This.ekler = New Collection
End Sub

Public Property Get id() As Variant
    id = This.id
End Property

Public Property Let id(ByVal vNewValue As Variant)
    This.id = vNewValue
End Property

Public Property Get baslik() As Variant
    baslik = This.baslik
End Property

Public Property Let baslik(ByVal vNewValue As Variant)
    This.baslik = vNewValue
End Property

Public Property Get icerik() As Variant
    icerik = This.icerik
End Property

Public Property Let icerik(ByVal vNewValue As Variant)
    This.icerik = vNewValue
End Property

Public Property Get topId() As Variant
    topId = This.topId
End Property

Public Property Let topId(ByVal vNewValue As Variant)
    This.topId = vNewValue
End Property

Private Sub Class_Terminate()
Set This.ekler = Nothing
End Sub
Public Property Set ekler(ByVal value As Collection)
    Set This.ekler = value
End Property
Public Property Get ekler() As Collection
    Set ekler = This.ekler
End Property
