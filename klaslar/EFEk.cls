VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EFEk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
id       As Integer
kararId  As Integer
path     As String
adi      As String
End Type

Private This As TState

Public Property Get id() As Integer
    id = This.id
End Property
Public Property Let id(ByVal idArg As Integer)
    This.id = idArg
End Property

Public Property Get kararId() As Integer
    kararId = This.kararId
End Property
Public Property Let kararId(ByVal kararIdArg As Integer)
    This.kararId = kararIdArg
End Property

Public Property Get path() As String
    path = This.path
End Property
Public Property Let path(ByVal pathArg As String)
    This.path = pathArg
End Property

Public Property Get adi() As String
    adi = This.adi
End Property
Public Property Let adi(ByVal adiArg As String)
    This.adi = adiArg
End Property
