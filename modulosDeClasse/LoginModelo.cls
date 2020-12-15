VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoginModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private id As Long
Private login As String
Private senha As String
Private operador As String

' id ------------------------------
Public Property Get GetId() As Long
    GetId = id
End Property

Public Property Let SetId(pId As Long)
    id = pId
End Property

'login--------------------------------
Public Property Get Getlogin() As String
    Getlogin = login
End Property

Public Property Let Setlogin(plogin As String)
    login = plogin
End Property
' senha ------------------------------

Public Property Get Getsenha() As String
    Getsenha = senha
End Property

Public Property Let Setsenha(psenha As String)
    senha = psenha
End Property

'operador --------------------
Public Property Get Getoperador() As String
    Getoperador = operador
End Property

Public Property Let Setoperador(poperador As String)
    operador = poperador
End Property





