VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClienteModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private nome As String
Private cpf As String
Private DataNasc As Date
Private DataCadastro As Date
Private limite As String
Private rg As String
Private Fone As String
Private endereco As String
Private numero As String
Private complemento As String
Private cep As String
Private bairro As String
Private Celular1 As String
Private Celular2 As String
Private conjuge As String
Private NascConjuge As Date
Private empresa As String
Private salario As String
Private admissao As Date

' ID

Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

' NOME

Public Property Get GetNome() As String
    GetNome = nome
End Property

Public Property Let SetNome(pNome As String)
    nome = pNome
End Property

' CPF

Public Property Get GetCpf() As String
    GetCpf = cpf
End Property

Public Property Let SetCpf(pCpf As String)
    cpf = pCpf
End Property

'DATA NASCIMENTO

Public Property Get GetDataNascimento() As Date
    GetDataNascimento = DataNasc
End Property

Public Property Let SetDataNascimento(pDataNasc As Date)
    DataNasc = pDataNasc
End Property

'DATA CADASTRO

Public Property Get GetDataCadastro() As Date
    GetDataCadastro = DataCadastro
End Property

Public Property Let SetDataCadastro(pDataCadastro As Date)
    DataCadastro = pDataCadastro
End Property

' LIMITE DE CREDITO

Public Property Get GetLimite() As String
    GetLimite = limite
End Property

Public Property Let SetLimite(pLimite As String)
    limite = pLimite
End Property

' RG

Public Property Get GetRg() As String
    GetRg = rg
End Property

Public Property Let SetRg(pRg As String)
    rg = pRg
End Property

' Fone

Public Property Get GetFone() As String
    GetFone = Fone
End Property

Public Property Let SetFone(pFone As String)
    Fone = pFone
End Property


' ENDERECO

Public Property Get Getendereco() As String
    Getendereco = endereco
End Property

Public Property Let Setendereco(pendereco As String)
    endereco = pendereco
End Property

' NUMERO

Public Property Get Getnumero() As String
    Getnumero = numero
End Property

Public Property Let Setnumero(pnumero As String)
    numero = pnumero
End Property


' COMPLEMENTO

Public Property Get GetComplemento() As String
    GetComplemento = complemento
End Property

Public Property Let SetComplemento(pComplemento As String)
    complemento = pComplemento
End Property

' CEP

Public Property Get Getcep() As String
    Getcep = cep
End Property

Public Property Let Setcep(pcep As String)
    cep = pcep
End Property

' BAIRRO

Public Property Get Getbairro() As String
    Getbairro = bairro
End Property

Public Property Let Setbairro(pbairro As String)
    bairro = pbairro
End Property

' CELULAR 1

Public Property Get GetCelular1() As String
    GetCelular1 = Celular1
End Property

Public Property Let SetCelular1(pCelular1 As String)
    Celular1 = pCelular1
End Property

' CELULAR 2

Public Property Get GetCelular2() As String
    GetCelular2 = Celular2
End Property

Public Property Let SetCelular2(pCelular2 As String)
    Celular2 = pCelular2
End Property

' CONJUGE

Public Property Get GetConjuge() As String
    GetConjuge = conjuge
End Property

Public Property Let SetConjuge(pConjuge As String)
    conjuge = pConjuge
End Property

' NASCIMENTO CONJUGE

Public Property Get GetNascConjuge() As Date
    GetNascConjuge = NascConjuge
End Property

Public Property Let SetNascConjuge(pNascConjuge As Date)
    NascConjuge = pNascConjuge
End Property

' EMPRESA

Public Property Get GetEmpresa() As String
    GetEmpresa = empresa
End Property

Public Property Let SetEmpresa(pEmpresa As String)
    empresa = pEmpresa
End Property

' SALARIO

Public Property Get GetSALARIO() As String
    GetSALARIO = salario
End Property

Public Property Let SetSALARIO(pSALARIO As String)
    salario = pSALARIO
End Property

' ADMISSAO

Public Property Get GetAdmissao() As Date
    GetAdmissao = admissao
End Property

Public Property Let SetAdmissao(pAdmissao As Date)
    admissao = pAdmissao
End Property





