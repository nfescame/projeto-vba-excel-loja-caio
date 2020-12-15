Attribute VB_Name = "RepositorDeMovEstoque"
Dim rs As New ADODB.Recordset
Public Function BuscarMovimentacao() As Collection
    
    Dim todosMovEst As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM MOVIMENTACAO_ESTOQUE", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim mov As MovEstoqueModelo
        Set mov = New MovEstoqueModelo
        
            mov.SetId = rs("ID")
            mov.Settipo = rs("TIPO")
            mov.Setdestino = rs("DESTINO")
            mov.Setquantidade = rs("QUANTIDADE")
            mov.Setdescricao = rs("DESCRICAO")
            mov.Setdata = rs("DATA")
            
            todosMovEst.Add mov
         
        
        rs.MoveNext
        
    Loop
    
    SQL.FecharConexao
    
    Set BuscarMovimentacao = todosMovEst
End Function

Public Sub AdicionarMovimentacao(mov As MovEstoqueModelo)
    Dim queryMov As String
    
    SQL.AbrirConexao
     
     queryMov = "INSERT INTO MOVIMENTACAO_ESTOQUE (TIPO,DESTINO,QUANTIDADE,DESCRICAO,DATA)" & _
        "VALUES ( '" & mov.Gettipo & "'," & _
        "'" & mov.Getdestino & "'," & _
        "'" & mov.Getquantidade & "'," & _
        "'" & mov.Getdescricao & "'," & _
        "'" & mov.Getdata & "')"
        
    
    SQL.Execute queryMov
    
    SQL.FecharConexao
End Sub
Public Function alterarLogin(logAtual As String, senhaNova As String) As Collection
    Dim queryLogin As String

    SQL.AbrirConexao
    
    queryLogin = "UPDATE LOGIN " _
    & " SET SENHA = '" & senhaNova & "'" _
    & " WHERE LOGIN =  " & logAtual & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryLogin
    
    SQL.FecharConexao
    
End Function

