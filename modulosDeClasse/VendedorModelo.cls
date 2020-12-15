VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VendedorModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private idv As Long
Private DataCadastroV As Date
Private NomeV As String
Public Property Get GetIdV() As Long
    GetIdV = idv
End Property

Public Property Let SetIdV(pIdV As Long)
    idv = pIdV
End Property
Public Property Get GetDataCadastroV() As Date
    GetDataCadastroV = DataCadastroV
End Property

Public Property Let SetDataCadastroV(pGetDataCadastroV As Date)
    DataCadastroV = pGetDataCadastroV
End Property
Public Property Get GetNomeV() As String
    GetNomeV = NomeV
End Property

Public Property Let SetNomeV(pNomeV As String)
    NomeV = pNomeV
End Property

