VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MovEstoqueModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private tipo As String
Private destino As String
Private quantidade As String
Private descricao As String
Private data As Date

' id ------------------------------
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

'tipo--------------------------------
Public Property Get Gettipo() As String
    Gettipo = tipo
End Property

Public Property Let Settipo(ptipo As String)
    tipo = ptipo
End Property

' destino ------------------------------
Public Property Get Getdestino() As String
    Getdestino = destino
End Property

Public Property Let Setdestino(pdestino As String)
    destino = pdestino
End Property

'quantidade --------------------
Public Property Get Getquantidade() As String
    Getquantidade = quantidade
End Property

Public Property Let Setquantidade(pquantidade As String)
    quantidade = pquantidade
End Property

'descricao --------------------
Public Property Get Getdescricao() As String
    Getdescricao = descricao
End Property

Public Property Let Setdescricao(pdescricao As String)
    descricao = pdescricao
End Property

'data --------------------
Public Property Get Getdata() As Date
    Getdata = data
End Property

Public Property Let Setdata(pdata As Date)
    data = pdata
End Property



