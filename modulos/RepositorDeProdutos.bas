Attribute VB_Name = "RepositorDeProdutos"
Dim rs As New ADODB.Recordset
Public Function BuscarTodosProdutos() As Collection
    
      Dim todosProdutos As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ESTOQUE c WHERE C.ID ", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim produtos As ProdutoModelo
        Set produtos = New ProdutoModelo
        
        produtos.SetId = rs("ID")
        produtos.Setdescricao = rs("DESCRICAO")
        produtos.Setgrupo = rs("GRUPO")
        produtos.Setquantidade = rs("QUANTIDADE")
        produtos.Setvalor = rs("VALOR")
        produtos.SetdataEntrada = rs("DATA_ENTRADA")
        produtos.Setcusto = rs("CUSTO")
        
        todosProdutos.Add produtos
        rs.MoveNext
        
    Loop
    
    SQL.FecharConexao
    
    Set BuscarTodosProdutos = todosProdutos
End Function


Public Function BuscarProdutoPorDescricao(descricao As String) As Collection
    
    Dim produtoCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ESTOQUE c WHERE c.DESCRICAO LIKE '%" & descricao & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim produto As ProdutoModelo
        Set produto = New ProdutoModelo
        
            produto.SetId = rs("ID")
            produto.Setdescricao = rs("DESCRICAO")
            produto.Setgrupo = rs("GRUPO")
            produto.Setquantidade = rs("QUANTIDADE")
            produto.Setvalor = rs("VALOR")
            produto.SetdataEntrada = rs("DATA_ENTRADA")
            produto.Setcusto = rs("CUSTO")
            
            produtoCoincientes.Add produto
       
       
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarProdutoPorDescricao = produtoCoincientes
End Function
Public Function BuscarProdutoPorId(idProduto As String) As Collection
    
    Dim idEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ESTOQUE c WHERE c.ID LIKE '%" & idProduto & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim produto As ProdutoModelo
        Set produto = New ProdutoModelo
        
        If rs("ID") = idProduto Then
        
            produto.SetId = rs("ID")
            produto.Setdescricao = rs("DESCRICAO")
            produto.Setgrupo = rs("GRUPO")
            produto.SetSubGrupo = rs("SUB_GRUPO")
            produto.Setquantidade = rs("QUANTIDADE")
            produto.Setvalor = rs("VALOR")
            produto.SetdataEntrada = rs("DATA_ENTRADA")
            produto.Setcusto = rs("CUSTO")
        
            idEncontrado.Add produto
  
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarProdutoPorId = idEncontrado
End Function
Public Function BuscarQuantidadePorId(idProduto As String) As Collection
    
    Dim idEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ESTOQUE c WHERE c.ID LIKE '%" & idProduto & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim produto As ProdutoModelo
        Set produto = New ProdutoModelo
        
        If rs("ID") = idProduto Then
        
            produto.Setquantidade = rs("QUANTIDADE")
    
            idEncontrado.Add produto
  
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarQuantidadePorId = idEncontrado
End Function

Public Function buscarUltimoId() As Collection
    Dim i As Integer
    Dim todosProdutos As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ESTOQUE", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim produtos As ProdutoModelo
        Set produtos = New ProdutoModelo
        
        i = i + 1
         
        rs.MoveNext
    Loop
    produtos.SetId = i
    SQL.FecharConexao
    todosProdutos.Add produtos
    Set buscarUltimoId = todosProdutos
End Function

Public Sub AdicionarProdutos(produto As ProdutoModelo)
    Dim queryProduto As String
    
    SQL.AbrirConexao
     
     queryProduto = "INSERT INTO ESTOQUE (DESCRICAO,GRUPO,SUB_GRUPO,QUANTIDADE,VALOR,DATA_ENTRADA,CUSTO,DATA_ATUALIZACAO)" & _
        "VALUES ( '" & produto.Getdescricao & "'," & _
                    "'" & produto.Getgrupo & "'," & _
                    "'" & produto.GetSubGrupo & "'," & _
                    "'" & produto.Getquantidade & "'," & _
                    "'" & produto.Getvalor & "'," & _
                    "'" & produto.GetdataEntrada & "'," & _
                    "'" & produto.Getcusto & "'," & _
                    "'" & produto.GetdataAtualizacao & "')"
        
        
    SQL.Execute queryProduto
    
    SQL.FecharConexao
End Sub
Public Function alterarProdutosPorId(idProduto As Integer, descricaoAT As String, grupoAT As String, subGrupoAT As String, quantidadeAT As String, valorAT As String, custoAT As String, dataAT) As Collection
    Dim queryProduto As String

    SQL.AbrirConexao
    
    queryProduto = "UPDATE ESTOQUE " _
    & " SET DESCRICAO = '" & descricaoAT & "',GRUPO = '" & grupoAT & "' ,SUB_GRUPO = '" & subGrupoAT & "' ,QUANTIDADE = '" & quantidadeAT & "' ,VALOR = '" & valorAT & "' ,CUSTO = '" & custoAT & "',DATA_ATUALIZACAO = '" & dataAT & "'  " _
    & " WHERE ID =  " & idProduto & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryProduto
    
    SQL.FecharConexao
    
End Function

Public Function alterarQuantidadePorId(idProduto As Integer, qvendido As Integer, quantidadeAT As Integer) As Collection
    Dim queryProduto As String

    SQL.AbrirConexao
    quantidadeAT = quantidadeAT - qvendido
    
    queryProduto = "UPDATE ESTOQUE " _
    & " SET QUANTIDADE = " & quantidadeAT & "  " _
    & " WHERE ID =  " & idProduto & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryProduto
    
    SQL.FecharConexao
    
End Function

