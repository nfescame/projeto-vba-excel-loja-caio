VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "idOrcAlteraModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long

' id ------------------------------
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

