VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecebimentosModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private dataCompra As Date
Private valorTotal As String
Private devedor As String
Private idDevedor As String
Private vendedor As String
Private idVendedor As String
Private vencimento As Date
Private valorParcela As String
Private parcelas As String
Private status As String
Private numeroVenda As Integer
Private idVenda As Integer

' ID DA COMPRA

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' DATA DA COMPRA

Public Property Get GetdataCompra() As Date
    GetdataCompra = dataCompra
End Property

Public Property Let SetdataCompra(pdataCompra As Date)
    dataCompra = pdataCompra
End Property

' VALOR TOTAL

Public Property Get GetValorTotal() As String
    GetValorTotal = valorTotal
End Property

Public Property Let SetValorTotal(pValorTotal As String)
    valorTotal = pValorTotal
End Property

'DEVEDOR

Public Property Get GetDevedor() As String
    GetDevedor = devedor
End Property

Public Property Let SetDevedor(pDevedor As String)
    devedor = pDevedor
End Property

'idDevedor

Public Property Get GetidDevedor() As String
    GetidDevedor = idDevedor
End Property

Public Property Let SetidDevedor(pidDevedor As String)
    idDevedor = pidDevedor
End Property

'VENDEDOR

Public Property Get Getvendedor() As String
    Getvendedor = vendedor
End Property

Public Property Let Setvendedor(pvendedor As String)
    vendedor = pvendedor
End Property

'ID vendedor

Public Property Get GetiDvendedor() As String
    GetiDvendedor = idVendedor
End Property

Public Property Let SetiDvendedor(PiDvendedor As String)
    idVendedor = PiDvendedor
End Property

' VENCIMENTO

Public Property Get GetVencimento() As Date
    GetVencimento = vencimento
End Property

Public Property Let SetVencimento(pVencimento As Date)
    vencimento = pVencimento
End Property

'valor das parcelas

Public Property Get GetvalorParcela() As String
    GetvalorParcela = valorParcela
End Property

Public Property Let SetvalorParcela(pvalorParcela As String)
    valorParcela = pvalorParcela
End Property


'QUANTIDADE DE PARCELAS

Public Property Get Getparcelas() As String
    Getparcelas = parcelas
End Property

Public Property Let Setparcelas(pparcelas As String)
    parcelas = pparcelas
End Property


' status DA COMPRA

Public Property Get GetSTATUS() As String
    GetSTATUS = status
End Property

Public Property Let SetSTATUS(pSTATUS As String)
    status = pSTATUS
End Property

' numero da Venda

Public Property Get GetnumeroVenda() As String
    GetnumeroVenda = numeroVenda
End Property

Public Property Let SetnumeroVenda(pnumeroVenda As String)
    numeroVenda = pnumeroVenda
End Property


' idVenda

Public Property Get GetidVenda() As String
    GetidVenda = idVenda
End Property

Public Property Let SetidVenda(pidVenda As String)
    idVenda = pidVenda
End Property





