VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToplantiManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub toplantiNo(toplanti As cToplanti)

''toplanti al?p nosunu bulacak fonk

End Sub

Sub toplantiEkle(toplanti As cToplanti)

''toplanti al?p o toplant?y? vtye ekleyecek sub
MsgBox "Toplant? Eklendi"
End Sub

Sub toplantiSil(toplanti As cToplanti)

'toplanti al?p o toplant?y? vtden silecek sub
MsgBox "Toplant? Silindi"


End Sub

Sub topTutanakOlustur(toplanti As cToplanti)

'toplanti al?p word tutanak olu?turacak
''No burada laz?m olabilir

End Sub

Sub topTutanakYazdir(toplanti As cToplanti)

''toplanti al?p yazd?racak
''Olustur dan sonra kullanmak mant?kl?

End Sub
Sub topTarihGuncelle(toplanti As cToplanti, newDate As Date)

End Sub

Sub topVeriCekBolumeGore(bolumId As Integer)

End Sub

Function topVeriCekId(topId As Integer) As Variant

End Function
