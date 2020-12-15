Attribute VB_Name = "RepositorDeDadosDaEmpresa"
Dim rs As New ADODB.Recordset
Public Function BuscarTodosDados() As Collection
    
    Dim todosDados As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM DADOS_DA_EMPRESA", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim dados As DadosDaEmpresaModelo
        Set dados = New DadosDaEmpresaModelo
        
        dados.SetId = rs("ID")
        dados.SetnomeEmpresa = rs("NOME")
        dados.Setendereco = rs("ENDERECO")
        dados.Setnumero = rs("NUMERO")
        dados.Setcep = rs("CEP")
        dados.Setbairro = rs("BAIRRO")
        dados.Setcidade = rs("CIDADE")
        dados.Settelefone = rs("TELEFONE")
        dados.Setcelular = rs("CELULAR")
        dados.Setemail = rs("E_MAIL")
        dados.SettextoOrc = rs("TEXTO_ORC")
        dados.SettextoVendas = rs("TEXTO_VENDAS")
        dados.SettextoCarne = rs("TEXTO_CARNE")
        
        todosDados.Add dados
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarTodosDados = todosDados
End Function
Public Sub AdicionarDadosEmpresa(dados As DadosDaEmpresaModelo)
    Dim queryDados As String
     
    SQL.AbrirConexao
    
    queryDados = "INSERT INTO DADOS_DA_EMPRESA (NOME,ENDERECO,NUMERO,CEP,BAIRRO,CIDADE,TELEFONE,CELULAR,E_MAIL,TEXTO_ORC,TEXTO_VENDAS,TEXTO_CARNE)" & _
        "VALUES ( '" & dados.GetnomeEmpresa & "'," & _
                 "'" & dados.Getendereco & "'," & _
                 "'" & dados.Getnumero & "'," & _
                 "'" & dados.Getcep & "'," & _
                 "'" & dados.Getbairro & "'," & _
                 "'" & dados.Getcidade & "'," & _
                 "'" & dados.Gettelefone & "'," & _
                 "'" & dados.Getcelular & "'," & _
                 "'" & dados.Getemail & "'," & _
                 "'" & dados.GettextoOrc & "'," & _
                 "'" & dados.GettextoVendas & "'," & _
                 "'" & dados.GettextoCarne & "')"
         
    
    SQL.Execute queryDados
    
    SQL.FecharConexao
End Sub
Public Function alterarDadosDaEmpresa(id As Integer, novoNome As String, novoEndereco As String, novoNumero As String, _
novoCep As String, novoBairro As String, novocidade As String, novoTel As String, novoCel As String, novoEmail As String, _
novoTextoOrc As String, novoTextVendas As String, novoTextoCarne As String) As Collection
    Dim queryDados As String
    

    SQL.AbrirConexao
    
    queryDados = "UPDATE DADOS_DA_EMPRESA " _
    & " SET NOME = '" & novoNome & " ' , ENDERECO = '" & novoEndereco & " ',NUMERO = '" _
    & novoNumero & " ' , CEP = '" & novoCep & " ' ,BAIRRO = '" & novoBairro & " ' ,  CIDADE = '" & novocidade & " ' , TELEFONE = '" _
    & novoTel & " ' ,CELULAR = '" & novoCel & " ' , E_MAIL = '" & novoEmail & " ' ,  TEXTO_ORC = '" _
    & novoTextoOrc & " ' ,TEXTO_CARNE = '" & novoTextoCarne & " ', TEXTO_VENDAS = '" & novoTextVendas & " '" _
    & " WHERE ID =  " & id & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryDados
    
    SQL.FecharConexao
    
End Function
