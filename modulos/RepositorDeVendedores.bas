Attribute VB_Name = "RepositorDeVendedores"
Dim rs As New ADODB.Recordset
Public Function BuscarTodosVendedores() As Collection
    
    Dim todosVendedores As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM VENDEDORES", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim vendedor As VendedorModelo
        Set vendedor = New VendedorModelo
        
        vendedor.SetIdV = rs("ID")
        vendedor.SetNomeV = rs("NOME")
        vendedor.SetDataCadastroV = rs("DATA_CADASTRO")
        
        
        todosVendedores.Add vendedor
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarTodosVendedores = todosVendedores
End Function


Public Function BuscarVendedorPorNome(nome As String) As Collection
    
    Dim vendedorCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM vendedores c WHERE c.NOME LIKE '%" & nome & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim vendedor As VendedorModelo
        Set vendedor = New VendedorModelo
        
        vendedor.SetIdV = rs("ID")
        vendedor.SetNomeV = rs("NOME")
        vendedor.SetDataCadastroV = rs("DATA_CADASTRO")
        
        
        vendedorCoincientes.Add vendedor
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendedorPorNome = vendedorCoincientes
End Function
Public Function BuscarVendedorPorId(idv As String) As Collection
    
    Dim idVEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM VENDEDORES c WHERE c.ID LIKE '%" & idv & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim idVendedores As VendedorModelo
        Set idVendedores = New VendedorModelo
        
        idVendedores.SetIdV = rs("ID")
        idVendedores.SetNomeV = rs("NOME")
        
        If rs("ID") = idv Then
        
            idVEncontrado.Add idVendedores
  
        End If
        
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendedorPorId = idVEncontrado
End Function

Public Sub AdicionarVendedores(vendedor As VendedorModelo)
    Dim queryVendedor As String
    
    
     SQL.AbrirConexao
    
    queryVendedor = "INSERT INTO VENDEDORES (nome,data_cadastro)" & _
        "VALUES ( '" & vendedor.GetNomeV & "'," & _
        "'" & vendedor.GetDataCadastroV & "')"
    
    SQL.Execute queryVendedor
    
    SQL.FecharConexao
    
End Sub
Public Function alterar(VendedorAtual As String, vendedorNovo As String) As Collection
    Dim queryVendedor As String

    SQL.AbrirConexao
    
    queryVendedor = "UPDATE VENDEDORES " _
    & " SET NOME = '" & vendedorNovo & "' " _
    & " WHERE NOME =  '" & VendedorAtual & " ' "
    SQL.GetConexao
    
    
    SQL.Execute queryVendedor
    
    SQL.FecharConexao
    
End Function


