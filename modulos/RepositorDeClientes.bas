Attribute VB_Name = "RepositorDeClientes"
Dim rs As New ADODB.Recordset
Public Function BuscarTodos() As Collection
    
    Dim todosClientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM CLIENTES", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim cliente As ClienteModelo
        Set cliente = New ClienteModelo
        
        cliente.SetId = rs("ID")
        cliente.SetNome = rs("NOME")
        cliente.SetCpf = rs("CPF")
        cliente.SetDataNascimento = rs("DATA_NASC")
        cliente.SetDataCadastro = rs("DATA_CADASTRO")
        cliente.SetLimite = rs("LIMITE")
        cliente.SetRg = rs("RG")
        cliente.SetFone = rs("TELEFONE")
        cliente.Setendereco = rs("ENDERECO")
        cliente.Setnumero = rs("NUMERO")
        cliente.SetComplemento = rs("COMPLEMENTO")
        cliente.Setcep = rs("CEP")
        cliente.Setbairro = rs("BAIRRO")
        cliente.SetCelular1 = rs("CELULAR_1")
        cliente.SetCelular2 = rs("CELULAR_2")
        cliente.SetConjuge = rs("CONJUGE")
        cliente.SetNascConjuge = rs("NASC_CONJUGE")
        cliente.SetEmpresa = rs("EMPRESA")
        cliente.SetSALARIO = rs("SALARIO")
        cliente.SetAdmissao = rs("ADMISSAO")
        
        
        todosClientes.Add cliente
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarTodos = todosClientes
End Function

Public Function BuscarClientePorNome(nome As String) As Collection
    
    Dim clientesCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM CLIENTES c WHERE c.NOME LIKE '%" & nome & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim cliente As ClienteModelo
        Set cliente = New ClienteModelo
        
        cliente.SetId = rs("ID")
        cliente.SetNome = rs("NOME")
        cliente.SetCpf = rs("CPF")
        cliente.SetDataNascimento = rs("DATA_NASC")
        cliente.SetDataCadastro = rs("DATA_CADASTRO")
        cliente.SetLimite = rs("LIMITE")
        cliente.SetRg = rs("RG")
        cliente.SetFone = rs("TELEFONE")
        cliente.Setendereco = rs("ENDERECO")
        cliente.Setnumero = rs("NUMERO")
        cliente.SetComplemento = rs("COMPLEMENTO")
        cliente.Setcep = rs("CEP")
        cliente.Setbairro = rs("BAIRRO")
        cliente.SetCelular1 = rs("CELULAR_1")
        cliente.SetCelular2 = rs("CELULAR_2")
        cliente.SetConjuge = rs("CONJUGE")
        cliente.SetNascConjuge = rs("NASC_CONJUGE")
        cliente.SetEmpresa = rs("EMPRESA")
        cliente.SetSALARIO = rs("SALARIO")
        cliente.SetAdmissao = rs("ADMISSAO")
        
        
        clientesCoincientes.Add cliente
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarClientePorNome = clientesCoincientes
End Function

Public Function BuscarClientePorId(id As String) As Collection
    
    Dim idEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM CLIENTES c WHERE c.ID LIKE '%" & id & "%'", SQL.GetConexao
    
     Do While Not rs.EOF
        Dim cliente As ClienteModelo
        Set cliente = New ClienteModelo
        
        If rs("ID") = id Then
        
            cliente.SetId = rs("ID")
            cliente.SetNome = rs("NOME")
            cliente.SetCpf = rs("CPF")
            cliente.SetDataNascimento = rs("DATA_NASC")
            cliente.SetDataCadastro = rs("DATA_CADASTRO")
            cliente.SetLimite = rs("LIMITE")
            cliente.SetRg = rs("RG")
            cliente.SetFone = rs("TELEFONE")
            cliente.Setendereco = rs("ENDERECO")
            cliente.Setnumero = rs("NUMERO")
            cliente.SetComplemento = rs("COMPLEMENTO")
            cliente.Setcep = rs("CEP")
            cliente.Setbairro = rs("BAIRRO")
            cliente.SetCelular1 = rs("CELULAR_1")
            cliente.SetCelular2 = rs("CELULAR_2")
            cliente.SetConjuge = rs("CONJUGE")
            cliente.SetNascConjuge = rs("NASC_CONJUGE")
            cliente.SetEmpresa = rs("EMPRESA")
            cliente.SetSALARIO = rs("SALARIO")
            cliente.SetAdmissao = rs("ADMISSAO")
            
            
            
            
            idEncontrado.Add cliente
            
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarClientePorId = idEncontrado
    
End Function


Public Sub AdicionarCliente(cliente As ClienteModelo)
    Dim queryCliente As String
     
    SQL.AbrirConexao
    
    queryCliente = "INSERT INTO CLIENTES (nome,cpf,data_nasc,data_cadastro,limite,rg,telefone,endereco,numero,complemento,cep,bairro,celular_1,celular_2,conjuge,nasc_conjuge,empresa,salario,admissao)" & _
        "VALUES ( '" & cliente.GetNome & "'," & _
                 "'" & cliente.GetCpf & "'," & _
                 "'" & cliente.GetDataNascimento & "'," & _
                 "'" & cliente.GetDataCadastro & "'," & _
                 "'" & cliente.GetLimite & "'," & _
                 "'" & cliente.GetRg & "'," & _
                 "'" & cliente.GetFone & "'," & _
                 "'" & cliente.Getendereco & "'," & _
                 "'" & cliente.Getnumero & "'," & _
                 "'" & cliente.GetComplemento & "'," & _
                 "'" & cliente.Getcep & "'," & _
                 "'" & cliente.Getbairro & "'," & _
                 "'" & cliente.GetCelular1 & "'," & _
                 "'" & cliente.GetCelular2 & "'," & _
                 "'" & cliente.GetConjuge & "'," & _
                 "'" & cliente.GetNascConjuge & "'," & _
                 "'" & cliente.GetEmpresa & "'," & _
                 "'" & cliente.GetSALARIO & "'," & _
                 "'" & cliente.GetAdmissao & "')"
         
    
    SQL.Execute queryCliente
    
    SQL.FecharConexao
End Sub
Public Function alterarCliente(idClienteAtual As Integer, novoNome As String, novoCpf As String, novoDataNasci As String, _
novoLimite As String, novoRg As String, novoTel As String, novoEnd As String, novoNumero As String, _
novoCompl As String, novoCep As String, novoBairro As String, novoCelular1 As String, novoCelular2 As String, _
novoConjuge As String, novoNascConj As String, novoEmpresa As String, novoSalario As String, novoAdmissao As String) As Collection
    Dim queryCliente As String
    

    SQL.AbrirConexao
    queryCliente = "UPDATE CLIENTES " _
    & " SET NOME = '" & novoNome & " ' , CPF = '" & novoCpf & " ',DATA_NASC = '" _
    & novoDataNasci & " ' , LIMITE = '" & novoLimite & " ' ,  RG = '" & novoRg & " ' , TELEFONE = '" _
    & novoTel & " ' ,ENDERECO = '" & novoEnd & " ' , NUMERO = '" & novoNumero & " ' ,  COMPLEMENTO = '" _
    & novoCompl & " ' , CEP = '" & novoCep & " ' , BAIRRO = '" & novoBairro & " ' , CELULAR_1 = '" _
    & novoCelular1 & " ', CELULAR_2 = '" & novoCelular2 & " ',CONJUGE = '" & novoConjuge & " ', NASC_CONJUGE = '" _
    & novoNascConj & " ' , EMPRESA = '" & novoEmpresa & " ' , SALARIO = '" & novoSalario & " ', ADMISSAO = '" & novoAdmissao & " '" _
    & " WHERE ID =  " & idClienteAtual & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryCliente
    
    SQL.FecharConexao
    
End Function




