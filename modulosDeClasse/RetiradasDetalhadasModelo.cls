VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RetiradasDetalhadasModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private identificacao As String
Private valor As Double
Private descricao As String
Private data As Date
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property



Public Property Get Getidentificacao() As String
    Getidentificacao = identificacao
End Property

Public Property Let Setidentificacao(pidentificacao As String)
    identificacao = pidentificacao
End Property



Public Property Get Getvalor() As Double
    Getvalor = valor
End Property

Public Property Let Setvalor(pvalor As Double)
    valor = pvalor
End Property



Public Property Get Getdescricao() As String
    Getdescricao = descricao
End Property

Public Property Let Setdescricao(pdescricao As String)
    descricao = pdescricao
End Property



Public Property Get Getdata() As Date
    Getdata = data
End Property

Public Property Let Setdata(pdata As Date)
    data = pdata
End Property


