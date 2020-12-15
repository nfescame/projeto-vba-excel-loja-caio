VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PagamentoModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private cliente As String
Private idCliente As Long
Private valorPg As String
Private parcela As String
Private idDebito As Long
Private dataVencimento As String
Private dataPagamento As String
Private pgDinheiro As String
Private pgCartao As String
Private juros As String
Private desconto As String
Private Qpagos As Integer

' ID

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' cliente

Public Property Get Getcliente() As String
    Getcliente = cliente
End Property

Public Property Let Setcliente(pcliente As String)
    cliente = pcliente
End Property

' idCliente

Public Property Get GetiDcliente() As Long
    GetiDcliente = idCliente
End Property

Public Property Let SetiDcliente(PiDcliente As Long)
    idCliente = PiDcliente
End Property

'valorPg

Public Property Get GetvalorPg() As String
    GetvalorPg = valorPg
End Property

Public Property Let SetvalorPg(pvalorPg As String)
    valorPg = pvalorPg
End Property

'parcela

Public Property Get Getparcela() As String
    Getparcela = parcela
End Property

Public Property Let Setparcela(pparcela As String)
    parcela = pparcela
End Property

'idDebito

Public Property Get GetidDebito() As Long
    GetidDebito = idDebito
End Property

Public Property Let SetidDebito(pidDebito As Long)
    idDebito = pidDebito
End Property

' dataVencimento

Public Property Get GetdataVencimento() As String
    GetdataVencimento = dataVencimento
End Property

Public Property Let SetdataVencimento(pdataVencimento As String)
    dataVencimento = pdataVencimento
End Property

' dataPagamento

Public Property Get GetdataPagamento() As String
    GetdataPagamento = dataPagamento
End Property

Public Property Let SetdataPagamento(pdataPagamento As String)
    dataPagamento = pdataPagamento
End Property

' pago em cartao

Public Property Get GetpgCartao() As String
    GetpgCartao = pgCartao
End Property

Public Property Let SetpgCartao(ppgCartao As String)
    pgCartao = ppgCartao
End Property

' pago em dinheiro

Public Property Get GetpgDinheiro() As String
    GetpgDinheiro = pgDinheiro
End Property

Public Property Let SetpgDinheiro(ppgDinheiro As String)
    pgDinheiro = ppgDinheiro
End Property

' juros

Public Property Get Getjuros() As String
    Getjuros = juros
End Property

Public Property Let Setjuros(pjuros As String)
    juros = pjuros
End Property

' desconto

Public Property Get GetDESCONTO() As String
    GetDESCONTO = desconto
End Property

Public Property Let SetDESCONTO(pDESCONTO As String)
    desconto = pDESCONTO
End Property

' desconto

Public Property Get GetqPagos() As String
    GetqPagos = Qpagos
End Property

Public Property Let SetqPagos(pqPagos As String)
    Qpagos = pqPagos
End Property





