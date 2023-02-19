VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEFPersonel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
Option Explicit

Private Type TState
    id              As Integer
    adi             As String
    soyadi          As String
    adSoyad         As String
    unvanId         As Integer
    bolumId         As Integer
    bkkKatilirMi    As Boolean
End Type

Private This As TState

Public Property Get id() As Integer
End Property
Public Property Let id(ByVal idArg As Integer)
End Property

Public Property Get adi() As String
End Property
Public Property Let adi(ByVal adiArg As String)
End Property

Public Property Get soyadi() As String
End Property
Public Property Let soyadi(ByVal soyadiArg As String)
End Property

Public Property Get adSoyad() As String
End Property

Public Property Get unvanId() As Integer
End Property
Public Property Let unvanId(ByVal unvanIdArg As Integer)
End Property

Public Property Get bolumId() As Integer
End Property
Public Property Let bolumId(ByVal bolumIdArg As Integer)
End Property

Public Property Get bkkKatilirMi() As Boolean
End Property
