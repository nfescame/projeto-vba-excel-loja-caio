VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProdutoModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private descricao As String
Private grupo As String
Private subGrupo As String
Private quantidade As Integer
Private valor As String
Private dataEntrada As Date
Private custo As String
Private dataAtualizacao As Date

' ID

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' descricao

Public Property Get Getdescricao() As String
    Getdescricao = descricao
End Property

Public Property Let Setdescricao(pdescricao As String)
    descricao = pdescricao
End Property

' grupo

Public Property Get Getgrupo() As String
    Getgrupo = grupo
End Property

Public Property Let Setgrupo(pgrupo As String)
    grupo = pgrupo
End Property

' subGrupo

Public Property Get GetSubGrupo() As String
    GetSubGrupo = subGrupo
End Property

Public Property Let SetSubGrupo(pSubGrupo As String)
    subGrupo = pSubGrupo
End Property

' quantidade

Public Property Get Getquantidade() As Long
    Getquantidade = quantidade
End Property

Public Property Let Setquantidade(pquantidade As Long)
    quantidade = pquantidade
End Property

' valor
Public Property Get Getvalor() As String
    Getvalor = valor
End Property

Public Property Let Setvalor(pvalor As String)
    valor = pvalor
End Property

' dataEntrada

Public Property Get GetdataEntrada() As Date
    GetdataEntrada = dataEntrada
End Property

Public Property Let SetdataEntrada(pdataEntrada As Date)
    dataEntrada = pdataEntrada
End Property

' custo

Public Property Get Getcusto() As String
    Getcusto = custo
End Property

Public Property Let Setcusto(pcusto As String)
    custo = pcusto
End Property

' data ultima atualizacao

Public Property Get GetdataAtualizacao() As Date
    GetdataAtualizacao = dataAtualizacao
End Property

Public Property Let SetdataAtualizacao(pdataAtualizacao As Date)
    dataAtualizacao = pdataAtualizacao
End Property

