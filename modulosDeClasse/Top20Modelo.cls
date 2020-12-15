VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Top20Modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim top20 As String

' top20

Public Property Get Gettop20() As String
    Gettop20 = top20
End Property

Public Property Let Settop20(ptop20 As String)
    top20 = ptop20
End Property

