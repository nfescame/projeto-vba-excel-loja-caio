VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RetiradasModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private DATARETIRADA As Date
Private retirada As Double
Private despesas As Double
Private salario As Double
Private comissao As Double
Private conducao As Double

' ID DA COMPRA

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' DATA DA RETIRADA

Public Property Get GetDATARETIRADA() As Date
    GetDATARETIRADA = DATARETIRADA
End Property

Public Property Let SetDATARETIRADA(PDATARETIRADA As Date)
    DATARETIRADA = PDATARETIRADA
End Property

' RETIRADA

Public Property Get GetRETIRADA() As Double
    GetRETIRADA = retirada
End Property

Public Property Let SetRETIRADA(PRETIRADA As Double)
    retirada = PRETIRADA
End Property

' DESPESAS

Public Property Get GetDESPESAS() As Double
    GetDESPESAS = despesas
End Property

Public Property Let SetDESPESAS(PDESPESAS As Double)
    despesas = PDESPESAS
End Property

' SALARIO

Public Property Get GetSALARIO() As Double
    GetSALARIO = salario
End Property

Public Property Let SetSALARIO(pSALARIO As Double)
    salario = pSALARIO
End Property

' COMISSAO

Public Property Get GetCOMISSAO() As Double
    GetCOMISSAO = comissao
End Property

Public Property Let SetCOMISSAO(PCOMISSAO As Double)
    comissao = PCOMISSAO
End Property

' CONDUCAO

Public Property Get GetCONDUCAO() As Double
    GetCONDUCAO = conducao
End Property

Public Property Let SetCONDUCAO(PCONDUCAO As Double)
    conducao = PCONDUCAO
End Property

