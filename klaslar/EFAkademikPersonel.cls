VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EFAkademikPersonel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEFPersonel

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

Public Property Get IEFPersonel_id() As Integer
    IEFPersonel_id = This.id
End Property
Public Property Let IEFPersonel_id(ByVal idArg As Integer)
   This.id = idArg
End Property

Public Property Get IEFPersonel_adi() As String
    IEFPersonel_adi = This.adi
End Property
Public Property Let IEFPersonel_adi(ByVal adiArg As String)
    This.adi = adiArg
End Property

Public Property Get IEFPersonel_soyadi() As String
    IEFPersonel_soyadi = This.soyadi
End Property
Public Property Let IEFPersonel_soyadi(ByVal soyadiArg As String)
    This.soyadi = soyadiArg
End Property

Public Property Get IEFPersonel_adSoyad() As String
    IEFPersonel_adSoyad = This.adi & " " & This.soyadi
End Property

Public Property Get IEFPersonel_unvanId() As Integer
    IEFPersonel_unvanId = This.unvanId
End Property
Public Property Let IEFPersonel_unvanId(ByVal unvanIdArg As Integer)
    This.unvanId = unvanIdArg
End Property

Public Property Get IEFPersonel_bolumId() As Integer
    IEFPersonel_bolumId = This.bolumId
End Property
Public Property Let IEFPersonel_bolumId(ByVal bolumIdArg As Integer)
   This.bolumId = bolumIdArg
End Property

Public Property Get IEFPersonel_bkkKatilirMi() As Boolean
    If Me.IEFPersonel_unvanId = 1 Or Me.IEFPersonel_unvanId = 2 Or Me.IEFPersonel_unvanId = 3 Or Me.IEFPersonel_unvanId = 4 Then
        IEFPersonel_bkkKatilirMi = True
    Else
        IEFPersonel_bkkKatilirMi = False
    End If
End Property

