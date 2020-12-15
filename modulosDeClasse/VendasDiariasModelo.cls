VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VendasDiariasModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private desconto As String
Private cartao As String
Private dinheiro As String
Private crediario As String
Private valorCompra As String
Private vendedor As String
Private idVendedor As Integer
Private cliente As String
Private idCliente As String
Private dataCompra As Date
Private QPARCELAS As Integer
Private NORCAMENTO As Integer
Private DATAORCAMENTO As Date
Private nVendas As Integer

' id ---------------------
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' DESCONTO ---------------------
Public Property Get GetDESCONTO() As String
    GetDESCONTO = desconto
End Property

Public Property Let SetDESCONTO(pDESCONTO As String)
    desconto = pDESCONTO
End Property

' CARTAO ---------------------
Public Property Get GetCARTAO() As String
    GetCARTAO = cartao
End Property

Public Property Let SetCARTAO(pCARTAO As String)
    cartao = pCARTAO
End Property

' DINHEIRO ---------------------
Public Property Get GetDINHEIRO() As String
    GetDINHEIRO = dinheiro
End Property

Public Property Let SetDINHEIRO(pDINHEIRO As String)
    dinheiro = pDINHEIRO
End Property

' CREDIARIO ---------------------
Public Property Get GetCREDIARIO() As String
    GetCREDIARIO = crediario
End Property

Public Property Let SetCREDIARIO(PCREDIARIO As String)
    crediario = PCREDIARIO
End Property


' valor da compra -----------------------
Public Property Get GetvalorCompra() As String
    GetvalorCompra = valorCompra
End Property

Public Property Let SetvalorCompra(pvalorCompra As String)
    valorCompra = pvalorCompra
End Property

' vendedor ------------------------------
Public Property Get Getvendedor() As String
    Getvendedor = vendedor
End Property

Public Property Let Setvendedor(pvendedor As String)
    vendedor = pvendedor
End Property

' iDvendedor ------------------------------
Public Property Get GetiDvendedor() As String
    GetiDvendedor = idVendedor
End Property

Public Property Let SetiDvendedor(PiDvendedor As String)
    idVendedor = PiDvendedor
End Property

'cliente --------------------
Public Property Get Getcliente() As String
    Getcliente = cliente
End Property

Public Property Let Setcliente(pcliente As String)
    cliente = pcliente
End Property

'iDcliente --------------------
Public Property Get GetiDcliente() As Integer
    GetiDcliente = idCliente
End Property

Public Property Let SetiDcliente(PiDcliente As Integer)
    idCliente = PiDcliente
End Property


'data compra --------------------
Public Property Get GetdataCompra() As Date
    GetdataCompra = dataCompra
End Property

Public Property Let SetdataCompra(pdataCompra As Date)
    dataCompra = pdataCompra
End Property

'QPARCELAS --------------------
Public Property Get GetQPARCELAS() As Integer
    GetQPARCELAS = QPARCELAS
End Property

Public Property Let SetQPARCELAS(pQPARCELASa As Integer)
    QPARCELAS = pQPARCELAS
End Property

'NORCAMENTO --------------------
Public Property Get GetNORCAMENTO() As Integer
    GetNORCAMENTO = NORCAMENTO
End Property

Public Property Let SetNORCAMENTO(pNORCAMENTO As Integer)
    NORCAMENTO = pNORCAMENTO
End Property

'DATAORCAMENTO --------------------
Public Property Get GetDATAORCAMENTO() As Date
    GetDATAORCAMENTO = DATAORCAMENTO
End Property

Public Property Let SetDATAORCAMENTO(pDATAORCAMENTO As Date)
    DATAORCAMENTO = pDATAORCAMENTO
End Property

'DATAORCAMENTO --------------------
Public Property Get GetNVENDAS() As Integer
    GetNVENDAS = nVendas
End Property

Public Property Let SetNVENDAS(pNVENDAS As Integer)
    nVendas = pNVENDAS
End Property

