VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GrupoModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private grupo As String
Private subGrupo As String


' id ------------------------------
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

'grupo--------------------------------
Public Property Get Getgrupo() As String
    Getgrupo = grupo
End Property

Public Property Let Setgrupo(pgrupo As String)
    grupo = pgrupo
End Property

'subGrupo--------------------------------
Public Property Get GetSubGrupo() As String
    GetSubGrupo = subGrupo
End Property

Public Property Let SetSubGrupo(pSubGrupo As String)
    subGrupo = pSubGrupo
End Property


