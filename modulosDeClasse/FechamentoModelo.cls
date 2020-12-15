VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FechamentoModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private dataFechamento As Date
Private vendas As String
Private troco As String
Private pagamentos As String
Private cartao As String
Private crediario As String
Private despesa As String
Private salario As String
Private comissao As String
Private conducao As String
Private retirada As String
Private proximoTroco As String
Private Qvendas As Long
Private qPagamentos As Long
Private desconto As String
Private juros As String
Private status As String

' ID

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' data Fechamento

Public Property Get GetdataFechamento() As Date
    GetdataFechamento = dataFechamento
End Property

Public Property Let SetdataFechamento(pdataFechamento As Date)
    dataFechamento = pdataFechamento
End Property

' vendas

Public Property Get Getvendas() As String
    Getvendas = vendas
End Property

Public Property Let Setvendas(pvendas As String)
    vendas = pvendas
End Property

'TROCO

Public Property Get GetTROCO() As String
    GetTROCO = troco
End Property

Public Property Let SetTROCO(PTROCO As String)
    troco = PTROCO
End Property

' pagamentos

Public Property Get Getpagamentos() As String
    Getpagamentos = pagamentos
End Property

Public Property Let Setpagamentos(ppagamentos As String)
    pagamentos = ppagamentos
End Property

'cartao

Public Property Get GetCARTAO() As String
    GetCARTAO = cartao
End Property

Public Property Let SetCARTAO(pCARTAO As String)
    cartao = pCARTAO
End Property

'CREDIARIO

Public Property Get GetCREDIARIO() As String
    GetCREDIARIO = crediario
End Property

Public Property Let SetCREDIARIO(PCREDIARIO As String)
    crediario = PCREDIARIO
End Property

'DESPESA

Public Property Get GetDESPESA() As String
    GetDESPESA = despesa
End Property

Public Property Let SetDESPESA(PDESPESA As String)
    despesa = PDESPESA
End Property

'SALARIO

Public Property Get GetSALARIO() As String
    GetSALARIO = salario
End Property

Public Property Let SetSALARIO(pSALARIO As String)
    salario = pSALARIO
End Property

'COMISSAO

Public Property Get GetCOMISSAO() As String
    GetCOMISSAO = comissao
End Property

Public Property Let SetCOMISSAO(PCOMISSAO As String)
    comissao = PCOMISSAO
End Property

'CONDUCAO

Public Property Get GetCONDUCAO() As String
    GetCONDUCAO = conducao
End Property

Public Property Let SetCONDUCAO(PCONDUCAO As String)
    conducao = PCONDUCAO
End Property

'RETIRADA

Public Property Get GetRETIRADA() As String
    GetRETIRADA = retirada
End Property

Public Property Let SetRETIRADA(PRETIRADA As String)
    retirada = PRETIRADA
End Property


'PROXIMOTROCO

Public Property Get GetPROXIMOTROCO() As String
    GetPROXIMOTROCO = proximoTroco
End Property

Public Property Let SetPROXIMOTROCO(PPROXIMOTROCO As String)
    proximoTroco = PPROXIMOTROCO
End Property

' Quantidade Vendas

Public Property Get GetQVendas() As Long
    GetQVendas = Qvendas
End Property

Public Property Let SetQVendas(pQVendas As Long)
    Qvendas = pQVendas
End Property

' Quantidade pagamentos

Public Property Get GetQpagamentos() As Long
    GetQpagamentos = qPagamentos
End Property

Public Property Let SetQpagamentos(pQpagamentos As Long)
    qPagamentos = pQpagamentos
End Property

' desconto

Public Property Get GetDESCONTO() As String
    GetDESCONTO = desconto
End Property

Public Property Let SetDESCONTO(pDESCONTO As String)
    desconto = pDESCONTO
End Property

' juros

Public Property Get Getjuros() As String
    Getjuros = juros
End Property

Public Property Let Setjuros(pjuros As String)
    juros = pjuros
End Property

' STATUS

Public Property Get GetSTATUS() As String
    GetSTATUS = status
End Property

Public Property Let SetSTATUS(pSTATUS As String)
    status = pSTATUS
End Property








