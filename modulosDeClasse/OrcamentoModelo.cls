VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrcamentoModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private vendedor As String
Private idVendedor As String
Private data As Date
Private produto As String
Private idProduto As String
Private grupo As String
Private subGrupo As String
Private quantidade As String
Private valorUnit As String
Private status As String
Private valor As String
Private numero As String
Private cliente As String
Private idCliente As String
Private dataRegistro As Date

' ID

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' vendedor

Public Property Get Getvendedor() As String
    Getvendedor = vendedor
End Property

Public Property Let Setvendedor(pvendedor As String)
    vendedor = pvendedor
End Property

' iDvendedor

Public Property Get GetiDvendedor() As String
    GetiDvendedor = idVendedor
End Property

Public Property Let SetiDvendedor(PiDvendedor As String)
    idVendedor = PiDvendedor
End Property

' data

Public Property Get Getdata() As Date
    Getdata = data
End Property

Public Property Let Setdata(pdata As Date)
    data = pdata
End Property

'produto

Public Property Get Getproduto() As String
    Getproduto = produto
End Property

Public Property Let Setproduto(pproduto As String)
    produto = pproduto
End Property

'iDproduto

Public Property Get GetiDproduto() As String
    GetiDproduto = idProduto
End Property

Public Property Let SetiDproduto(piDproduto As String)
    idProduto = piDproduto
End Property

'grupo

Public Property Get Getgrupo() As String
    Getgrupo = grupo
End Property

Public Property Let Setgrupo(pgrupo As String)
    grupo = pgrupo
End Property

'SubGrupo

Public Property Get GetSubGrupo() As String
    GetSubGrupo = subGrupo
End Property

Public Property Let SetSubGrupo(pSubGrupo As String)
    subGrupo = pSubGrupo
End Property

'quantidade

Public Property Get Getquantidade() As String
    Getquantidade = quantidade
End Property

Public Property Let Setquantidade(pquantidade As String)
    quantidade = pquantidade
End Property

'valorUnit

Public Property Get GetvalorUnit() As String
    GetvalorUnit = valorUnit
End Property

Public Property Let SetvalorUnit(pvalorUnit As String)
    valorUnit = pvalorUnit
End Property

' status

Public Property Get GetSTATUS() As String
    GetSTATUS = status
End Property

Public Property Let SetSTATUS(pSTATUS As String)
    status = pSTATUS
End Property

' valor

Public Property Get Getvalor() As String
    Getvalor = valor
End Property

Public Property Let Setvalor(pvalor As String)
    valor = pvalor
End Property

' numero

Public Property Get Getnumero() As String
    Getnumero = numero
End Property

Public Property Let Setnumero(pnumero As String)
    numero = pnumero
End Property

' cliente

Public Property Get Getcliente() As String
    Getcliente = cliente
End Property

Public Property Let Setcliente(pcliente As String)
    cliente = pcliente
End Property


' idCliente

Public Property Get GetiDcliente() As String
    GetiDcliente = idCliente
End Property

Public Property Let SetiDcliente(PiDcliente As String)
    idCliente = PiDcliente
End Property

' dataRegistro

Public Property Get GetdataRegistro() As String
    GetdataRegistro = dataRegistro
End Property

Public Property Let SetdataRegistro(pdataRegistro As String)
    dataRegistro = pdataRegistro
End Property


