VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JurosModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private juros As Integer

' id ------------------------------
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

'juros--------------------------------
Public Property Get Getjuros() As String
    Getjuros = juros
End Property

Public Property Let Setjuros(pjuros As String)
    juros = pjuros
End Property

