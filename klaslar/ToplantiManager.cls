VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToplantiManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TToplantiManager
    toplantidao As IToplantiDao
End Type

Private this As TToplantiManager

Public Function Create(toplantidao As IToplantiDao) As ToplantiManager
    With New ToplantiManager
        Set .dao = toplantidao
        Set Create = .Self
    End With
End Function

Public Property Let dao(ByVal value As Variant)
    this.toplantidao = value
End Property
Public Property Get Self() As ToplantiManager
    Set Self = Me
End Property
Sub add(toplanti1 As clsToplanti, toplantidao As IToplantiDao)

Call toplantidao.add(toplanti1)
''toplanti al�p o toplant�y� vtye ekleyecek sub


End Sub


Sub deleteToplanti(toplanti As clsToplanti)

'toplanti al�p o toplant�y� vtden silecek sub
MsgBox "Toplant� Silindi"

End Sub
Sub numberOfToplanti(toplanti As clsToplanti)

''toplanti al�p nosunu bulacak fonk

End Sub

Sub createToplantiTutanak(toplanti As clsToplanti)

'toplanti al�p word tutanak olu�turacak
''No burada laz�m olabilir

End Sub

Sub printToplantiTutanak(toplanti As clsToplanti)

''toplanti al�p yazd�racak
''Olustur dan sonra kullanmak mant�kl�

End Sub
Sub updateToplantiDate(toplanti As clsToplanti, newDate As Date)

End Sub

Sub getTopByBolumId(bolumId As Integer)

End Sub

Function getTopByTopId(topId As Integer) As Variant

End Function
