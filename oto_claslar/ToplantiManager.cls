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

''toplanti alýp nosunu bulacak fonk

End Sub

Sub toplantiEkle(toplanti As cToplanti)

''toplanti alýp o toplantýyý vtye ekleyecek sub
MsgBox "Toplantý Eklendi"
End Sub

Sub toplantiSil(toplanti As cToplanti)

'toplanti alýp o toplantýyý vtden silecek sub
MsgBox "Toplantý Silindi"


End Sub

Sub topTutanakOlustur(toplanti As cToplanti)

'toplanti alýp word tutanak oluþturacak
''No burada lazým olabilir

End Sub

Sub topTutanakYazdir(toplanti As cToplanti)

''toplanti alýp yazdýracak
''Olustur dan sonra kullanmak mantýklý

End Sub
Sub topTarihGuncelle(toplanti As cToplanti, newDate As Date)

End Sub

Sub topVeriCekBolumeGore(bolumId As Integer)

End Sub

Function topVeriCekId(topId As Integer) As Variant

End Function
