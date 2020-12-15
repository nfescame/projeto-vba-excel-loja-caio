VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DadosDaEmpresaModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private nomeEmpresa As String
Private endereco As String
Private numero As String
Private cep As String
Private bairro As String
Private cidade As String
Private telefone As String
Private celular As String
Private email As String
Private textoOrc As String
Private textoVendas As String
Private textoCarne As String
Private caminhoBD As String

' ID

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' nomeEmpresa

Public Property Get GetnomeEmpresa() As String
    GetnomeEmpresa = nomeEmpresa
End Property

Public Property Let SetnomeEmpresa(pnomeEmpresa As String)
    nomeEmpresa = pnomeEmpresa
End Property

' endereco

Public Property Get Getendereco() As String
    Getendereco = endereco
End Property

Public Property Let Setendereco(pendereco As String)
    endereco = pendereco
End Property

' numero

Public Property Get Getnumero() As String
    Getnumero = numero
End Property

Public Property Let Setnumero(pnumero As String)
    numero = pnumero
End Property

'cep

Public Property Get Getcep() As String
    Getcep = cep
End Property

Public Property Let Setcep(pcep As String)
    cep = pcep
End Property

'bairro

Public Property Get Getbairro() As String
    Getbairro = bairro
End Property

Public Property Let Setbairro(pbairro As String)
    bairro = pbairro
End Property

'cidade

Public Property Get Getcidade() As String
    Getcidade = cidade
End Property

Public Property Let Setcidade(pcidade As String)
    cidade = pcidade
End Property

'telefone

Public Property Get Gettelefone() As String
    Gettelefone = telefone
End Property

Public Property Let Settelefone(ptelefone As String)
    telefone = ptelefone
End Property

'celular

Public Property Get Getcelular() As String
    Getcelular = celular
End Property

Public Property Let Setcelular(pcelular As String)
    celular = pcelular
End Property

'email

Public Property Get Getemail() As String
    Getemail = email
End Property

Public Property Let Setemail(pemail As String)
    email = pemail
End Property

' textoOrc

Public Property Get GettextoOrc() As String
    GettextoOrc = textoOrc
End Property

Public Property Let SettextoOrc(ptextoOrc As String)
    textoOrc = ptextoOrc
End Property

' textoVendas

Public Property Get GettextoVendas() As String
    GettextoVendas = textoVendas
End Property

Public Property Let SettextoVendas(ptextoVendas As String)
    textoVendas = ptextoVendas
End Property

' textoCarne

Public Property Get GettextoCarne() As String
    GettextoCarne = textoCarne
End Property

Public Property Let SettextoCarne(ptextoCarne As String)
    textoCarne = ptextoCarne
End Property

' textoCarne

Public Property Get GetcaminhoBD() As String
    GetcaminhoBD = caminhoBD
End Property

Public Property Let SetcaminhoBD(pcaminhoBD As String)
    caminhoBD = pcaminhoBD
End Property

