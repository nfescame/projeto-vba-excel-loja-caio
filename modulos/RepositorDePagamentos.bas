Attribute VB_Name = "RepositorDePagamentos"
Dim rs As New ADODB.Recordset
Public Function BuscarTodosPagamentos(data As Date) As Collection
    Dim selectCmd As String
    Dim pagamentoCoincientes As New Collection
   
    selectCmd = "SELECT SUM(VALOR_PG) as VALOR_PG , SUM (PG_DINHEIRO) AS PG_DINHEIRO, SUM (PG_CARTAO) AS PG_CARTAO, SUM (JUROS) as JUROS " & _
    "FROM PAGAMENTOS c WHERE c.DATA_PG LIKE '%" & data & "%'" & _
    "GROUP BY DATA_PG"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim pagamento As PagamentoModelo
        Set pagamento = New PagamentoModelo
        
          
                
                ''pagamento.SetId = rs("ID")
                ''pagamento.Setcliente = rs("CLIENTE")
                ''pagamento.SetiDcliente = rs("ID_CLIENTE")
            pagamento.SetvalorPg = rs("VALOR_PG")
                ''pagamento.SetidDebito = rs("ID_DEBITO")
                ''pagamento.SetdataVencimento = rs("DATA_VENCIMENTO")
                ''pagamento.SetdataPagamento = rs("DATA_PG")
            pagamento.SetpgDinheiro = rs("PG_DINHEIRO")
            pagamento.SetpgCartao = rs("PG_CARTAO")
            pagamento.Setjuros = rs("JUROS")
                ''pagamento.SetDESCONTO = rs("DESCONTO")
                
         
            
            pagamentoCoincientes.Add pagamento
                
            rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set BuscarTodosPagamentos = pagamentoCoincientes
End Function

Public Function RelatorioPagamentosPrint(data As Date) As Collection
    Dim selectCmd As String
    Dim pagamentoCoincientes As New Collection
   
    selectCmd = "SELECT * FROM PAGAMENTOS c WHERE c.DATA_PG LIKE '%" & data & "%'"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim pagamento As PagamentoModelo
        Set pagamento = New PagamentoModelo
        
          
                
            pagamento.SetId = rs("ID")
            pagamento.Setcliente = rs("CLIENTE")
            pagamento.SetiDcliente = rs("ID_CLIENTE")
            pagamento.SetvalorPg = rs("VALOR_PG")
            pagamento.Setparcela = rs("PARCELA")
            pagamento.SetidDebito = rs("ID_DEBITO")
            pagamento.SetdataVencimento = rs("DATA_VENCIMENTO")
            pagamento.SetdataPagamento = rs("DATA_PG")
            pagamento.SetpgDinheiro = rs("PG_DINHEIRO")
            pagamento.SetpgCartao = rs("PG_CARTAO")
            pagamento.Setjuros = rs("JUROS")
            pagamento.SetDESCONTO = rs("DESCONTO")
                
         
            
            pagamentoCoincientes.Add pagamento
                
            rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set RelatorioPagamentosPrint = pagamentoCoincientes
End Function

Public Function BuscarPagamentosPrint(id As Integer) As Collection
    Dim selectCmd As String
    Dim pagamentoCoincientes As New Collection
   
    selectCmd = "SELECT * FROM PAGAMENTOS c WHERE c.ID_DEBITO LIKE '%" & id & "%'"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim pagamento As PagamentoModelo
        Set pagamento = New PagamentoModelo
        
          
                
            pagamento.SetId = rs("ID")
            pagamento.Setcliente = rs("CLIENTE")
            pagamento.SetiDcliente = rs("ID_CLIENTE")
            pagamento.SetvalorPg = rs("VALOR_PG")
            pagamento.Setparcela = rs("PARCELA")
            pagamento.SetidDebito = rs("ID_DEBITO")
            pagamento.SetdataVencimento = rs("DATA_VENCIMENTO")
            pagamento.SetdataPagamento = rs("DATA_PG")
            pagamento.SetpgDinheiro = rs("PG_DINHEIRO")
            pagamento.SetpgCartao = rs("PG_CARTAO")
            pagamento.Setjuros = rs("JUROS")
            pagamento.SetDESCONTO = rs("DESCONTO")
                
         
            
            pagamentoCoincientes.Add pagamento
                
            rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set BuscarPagamentosPrint = pagamentoCoincientes
End Function
Public Function BuscarPagamentosPorPeriodo(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    Dim selectCmd As String
    Dim pagamentoCoincientes As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT * FROM PAGAMENTOS c WHERE c.DATA_PG BETWEEN " & _
    "#" & di & "# AND #" & df & "# "
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim pagamento As PagamentoModelo
        Set pagamento = New PagamentoModelo
        
            pagamento.SetId = rs("ID")
            pagamento.Setcliente = rs("CLIENTE")
            pagamento.SetiDcliente = rs("ID_CLIENTE")
            pagamento.SetvalorPg = rs("VALOR_PG")
            pagamento.Setparcela = rs("PARCELA")
            pagamento.SetidDebito = rs("ID_DEBITO")
            pagamento.SetdataVencimento = rs("DATA_VENCIMENTO")
            pagamento.SetdataPagamento = rs("DATA_PG")
            pagamento.SetpgDinheiro = rs("PG_DINHEIRO")
            pagamento.SetpgCartao = rs("PG_CARTAO")
            pagamento.Setjuros = rs("JUROS")
            pagamento.SetDESCONTO = rs("DESCONTO")
                
         
            pagamentoCoincientes.Add pagamento
                
            rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set BuscarPagamentosPorPeriodo = pagamentoCoincientes
End Function
Public Function BuscarPagamentosPorPeriodoNome(dataIPesquisa As Date, datafPesquisa As Date, cliente As String) As Collection
    Dim selectCmd As String
    Dim pagamentoCoincientes As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT * FROM PAGAMENTOS c WHERE c.DATA_PG BETWEEN " & _
    "#" & di & "# AND #" & df & "# "
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim pagamento As PagamentoModelo
        Set pagamento = New PagamentoModelo
            If rs("CLIENTE") = cliente Then
               pagamento.SetId = rs("ID")
               pagamento.Setcliente = rs("CLIENTE")
               pagamento.SetiDcliente = rs("ID_CLIENTE")
               pagamento.SetvalorPg = rs("VALOR_PG")
               pagamento.Setparcela = rs("PARCELA")
               pagamento.SetidDebito = rs("ID_DEBITO")
               pagamento.SetdataVencimento = rs("DATA_VENCIMENTO")
               pagamento.SetdataPagamento = rs("DATA_PG")
               pagamento.SetpgDinheiro = rs("PG_DINHEIRO")
               pagamento.SetpgCartao = rs("PG_CARTAO")
               pagamento.Setjuros = rs("JUROS")
               pagamento.SetDESCONTO = rs("DESCONTO")
                   
            
               pagamentoCoincientes.Add pagamento
            End If
            rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set BuscarPagamentosPorPeriodoNome = pagamentoCoincientes
End Function
Public Function BuscarQuantidadeDePagamentos(data As Date) As Collection
    Dim Qpg As New Collection
    Dim n As PagamentoModelo
    Set n = New PagamentoModelo
    
    SQL.AbrirConexao
    rs.Open "SELECT * FROM PAGAMENTOS c WHERE c.DATA_PG LIKE '%" & data & "%'", SQL.GetConexao
    
    If Not rs.EOF Then
    Do While Not rs.EOF
    Dim i As Integer
    i = i + 1
    
    rs.MoveNext
    
    Loop
    n.SetqPagos = i
    End If
    
    SQL.FecharConexao
    Qpg.Add n
    Set BuscarQuantidadeDePagamentos = Qpg
End Function
Public Function BuscarUltimoPagamentos() As Collection
Dim ultNumPg As New Collection
    Dim n As PagamentoModelo
    Set n = New PagamentoModelo
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT * FROM PAGAMENTOS c WHERE c.ID "
    
    rs.CursorType = adOpenKeyset
    rs.Open selectCmd, SQL.GetConexao
    On Error Resume Next
    rs.MoveLast
    
    If rs("DATA") = data Then
    If Not rs.EOF Then
    Do While Not rs.EOF
    DoEvents
    
    n.SetId = rs("ID")
    
    rs.MoveNext
    Loop
    End If
    End If
    SQL.FecharConexao
    ultNumPg.Add n
    Set BuscarUltimoPagamentos = ultNumPg

End Function
Public Sub AdicionarPagamentos(recebidos As PagamentoModelo)

 Dim queryPagamento As String
     
    SQL.AbrirConexao
    
    queryPagamento = "INSERT INTO PAGAMENTOS (id_cliente,cliente,valor_pg,parcela,id_debito,data_vencimento,data_pg,pg_dinheiro,pg_cartao,juros,desconto)" & _
        "VALUES ( '" & recebidos.GetiDcliente & "'," & _
                 "'" & recebidos.Getcliente & "'," & _
                 "'" & recebidos.GetvalorPg & "'," & _
                 "'" & recebidos.Getparcela & "'," & _
                 "'" & recebidos.GetidDebito & "'," & _
                 "'" & recebidos.GetdataVencimento & "'," & _
                 "'" & recebidos.GetdataPagamento & "'," & _
                 "'" & recebidos.GetpgDinheiro & "'," & _
                 "'" & recebidos.GetpgCartao & "'," & _
                 "'" & recebidos.Getjuros & "'," & _
                 "'" & recebidos.GetDESCONTO & "')"
         
    
    SQL.Execute queryPagamento
    
    SQL.FecharConexao

End Sub
Public Function alterarPagamento(id As Integer, dinheiro As String, cartao As String, juros As String) As Collection
    Dim queryPagamento As String
    
    
    SQL.AbrirConexao
    
    queryPagamento = "UPDATE PAGAMENTOS " _
    & " SET pg_dinheiro = '" & dinheiro & "', pg_cartao = '" & cartao & "',juros = '" & juros & "'" _
    & " WHERE ID =  " & id & "  "
    SQL.GetConexao
   
    
    SQL.Execute queryPagamento
    
    SQL.FecharConexao
    
End Function


