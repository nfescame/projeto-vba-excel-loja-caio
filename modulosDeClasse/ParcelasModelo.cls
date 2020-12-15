VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParcelasModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private idP As Long
Private parcela As String
Private taxa As String
Private juros As Long

Public Property Get GetIdP() As Long
    GetIdP = idP
End Property

Public Property Let SetIdP(pIdP As Long)
    idP = pIdP
End Property
Public Property Get Getparcela() As String
    Getparcela = parcela
End Property

Public Property Let Setparcela(pGetParcela As String)
    parcela = pGetParcela
End Property

Public Property Get GetTaxa() As String
    GetTaxa = taxa
End Property

Public Property Let SetTaxa(pTaxa As String)
    taxa = pTaxa
End Property


'juros --------------------
Public Property Get Getjuros() As Long
    Getjuros = juros
End Property

Public Property Let Setjuros(pjuros As Long)
    juros = pjuros
End Property




