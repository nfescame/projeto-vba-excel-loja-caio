Attribute VB_Name = "RepositorDeFechamento"
Dim rs As New ADODB.Recordset
Public Function BuscarFechamentosPrint(id As Integer) As Collection
    
    Dim FechamentosCoincientes As New Collection
    
    SQL.AbrirConexao
    rs.Open "SELECT * FROM FECHAMENTOS c WHERE c.ID LIKE '%" & id & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim fechamento As FechamentoModelo
        Set fechamento = New FechamentoModelo
  
                fechamento.SetId = rs("ID")
                fechamento.SetdataFechamento = rs("DATA_FECHAMENTO")
                fechamento.Setvendas = rs("VENDAS")
                fechamento.SetTROCO = rs("TROCO")
                fechamento.Setpagamentos = rs("PAGAMENTOS")
                fechamento.SetCARTAO = rs("CARTAO")
                fechamento.SetCREDIARIO = rs("CREDIARIO")
                fechamento.SetDESPESA = rs("DESPESA")
                fechamento.SetSALARIO = rs("SALARIO")
                fechamento.SetCOMISSAO = rs("COMISSAO")
                fechamento.SetCONDUCAO = rs("CONDUCAO")
                fechamento.SetRETIRADA = rs("RETIRADA")
                fechamento.SetPROXIMOTROCO = rs("PROXIMO_TROCO")
                fechamento.SetQVendas = rs("QUANTIDADE_VENDAS")
                fechamento.SetQpagamentos = rs("QUANTIDADE_PAGAMENTOS")
                fechamento.SetDESCONTO = rs("DESCONTO")
                fechamento.Setjuros = rs("JUROS")
                

                FechamentosCoincientes.Add fechamento
                
                rs.MoveNext
            Loop
        
    SQL.FecharConexao
    
    Set BuscarFechamentosPrint = FechamentosCoincientes
End Function

Public Function BuscarFechamentosPorPeriodo(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim FechamentosCoincientes As New Collection
    
    SQL.AbrirConexao
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT ID,DATA_FECHAMENTO,VENDAS,TROCO,PAGAMENTOS,CARTAO,CREDIARIO,DESPESA,SALARIO,COMISSAO,CONDUCAO,RETIRADA,PROXIMO_TROCO,QUANTIDADE_VENDAS,QUANTIDADE_PAGAMENTOS,DESCONTO,JUROS " & _
    "FROM FECHAMENTOS c WHERE c.DATA_FECHAMENTO BETWEEN " & _
    "#" & di & "# AND #" & df & "# "

    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
    
        Dim fechamento As FechamentoModelo
        Set fechamento = New FechamentoModelo
        
        fechamento.SetId = rs("ID")
        fechamento.SetdataFechamento = rs("DATA_FECHAMENTO")
        fechamento.Setvendas = rs("VENDAS")
        fechamento.SetTROCO = rs("TROCO")
        fechamento.Setpagamentos = rs("PAGAMENTOS")
        fechamento.SetCARTAO = rs("CARTAO")
        fechamento.SetCREDIARIO = rs("CREDIARIO")
        fechamento.SetDESPESA = rs("DESPESA")
        fechamento.SetSALARIO = rs("SALARIO")
        fechamento.SetCOMISSAO = rs("COMISSAO")
        fechamento.SetCONDUCAO = rs("CONDUCAO")
        fechamento.SetRETIRADA = rs("RETIRADA")
        fechamento.SetPROXIMOTROCO = rs("PROXIMO_TROCO")
        fechamento.SetQVendas = rs("QUANTIDADE_VENDAS")
        fechamento.SetQpagamentos = rs("QUANTIDADE_PAGAMENTOS")
        fechamento.SetDESCONTO = rs("DESCONTO")
        fechamento.Setjuros = rs("JUROS")
            
    
        FechamentosCoincientes.Add fechamento
                
        rs.MoveNext
    Loop
        
    SQL.FecharConexao
    
    Set BuscarFechamentosPorPeriodo = FechamentosCoincientes
End Function

Public Function BuscarTroco(id As Integer) As Collection
    Dim troco As New Collection
    Dim t As FechamentoModelo
    Set t = New FechamentoModelo
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM FECHAMENTOS c WHERE c.ID LIKE '%" & id & "%'", SQL.GetConexao
    
   
    If Not rs.EOF Then
    Do While Not rs.EOF
    DoEvents
    
    t.SetPROXIMOTROCO = rs("PROXIMO_TROCO")
    t.SetId = rs("ID")
    
    rs.MoveNext
    Loop
   
    End If
    SQL.FecharConexao
    troco.Add t
    Set BuscarTroco = troco
End Function
Public Function BuscarTrocoDoDia(id As Integer) As Collection
    Dim troco As New Collection
    Dim t As FechamentoModelo
    Set t = New FechamentoModelo
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM FECHAMENTOS c WHERE c.ID LIKE '%" & id & "%'", SQL.GetConexao
    
    
    If Not rs.EOF Then
    Do While Not rs.EOF
    DoEvents
    
    t.SetTROCO = rs("TROCO")
    
    
    rs.MoveNext
    Loop
   
    End If
    SQL.FecharConexao
    troco.Add t
    Set BuscarTrocoDoDia = troco
End Function

Public Function BuscarStatusCaixaAnterior(id As Integer) As Collection
    Dim statusAtual As New Collection
    Dim status As FechamentoModelo
    Set status = New FechamentoModelo
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM FECHAMENTOS c WHERE c.ID LIKE '%" & id & "%'", SQL.GetConexao
    
    If Not rs.EOF Then

    status.SetId = rs("ID")
    status.SetdataFechamento = rs("DATA_FECHAMENTO")
    status.Setvendas = rs("VENDAS")
    status.SetTROCO = rs("TROCO")
    status.Setpagamentos = rs("PAGAMENTOS")
    status.SetCARTAO = rs("CARTAO")
    status.SetCREDIARIO = rs("CREDIARIO")
    status.SetDESPESA = rs("DESPESA")
    status.SetSALARIO = rs("SALARIO")
    status.SetCOMISSAO = rs("COMISSAO")
    status.SetCONDUCAO = rs("CONDUCAO")
    status.SetRETIRADA = rs("RETIRADA")
    status.SetPROXIMOTROCO = rs("PROXIMO_TROCO")
    status.SetQVendas = rs("QUANTIDADE_VENDAS")
    status.SetQpagamentos = rs("QUANTIDADE_PAGAMENTOS")
    status.SetDESCONTO = rs("DESCONTO")
    status.Setjuros = rs("JUROS")
    status.SetSTATUS = rs("STATUS")

    End If
    SQL.FecharConexao
    statusAtual.Add status
    Set BuscarStatusCaixaAnterior = statusAtual
End Function

Public Function BuscarUltimoIdCaixa() As Collection
    Dim statusAtual As New Collection
    Dim status As FechamentoModelo
    Set status = New FechamentoModelo
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT MAX(ID) FROM FECHAMENTOS c WHERE c.STATUS "
    
    rs.Open selectCmd, SQL.GetConexao
    
        Do Until rs.EOF
            status.SetId = "" & rs(0)
    
            rs.MoveNext
        Loop
    

    SQL.FecharConexao
    statusAtual.Add status
    Set BuscarUltimoIdCaixa = statusAtual
End Function
Public Sub AdicionarFechamento(fechamento As FechamentoModelo)

 Dim queryFechamento As String
     
    SQL.AbrirConexao
    
    queryFechamento = "INSERT INTO FECHAMENTOS (DATA_FECHAMENTO,VENDAS,TROCO,PAGAMENTOS,CARTAO,CREDIARIO,DESPESA,SALARIO,COMISSAO,CONDUCAO,RETIRADA,PROXIMO_TROCO,QUANTIDADE_VENDAS,QUANTIDADE_PAGAMENTOS,DESCONTO,JUROS,STATUS)" & _
        "VALUES ( '" & fechamento.GetdataFechamento & "'," & _
                 "'" & fechamento.Getvendas & "'," & _
                 "'" & fechamento.GetTROCO & "'," & _
                 "'" & fechamento.Getpagamentos & "'," & _
                 "'" & fechamento.GetCARTAO & "'," & _
                 "'" & fechamento.GetCREDIARIO & "'," & _
                 "'" & fechamento.GetDESPESA & "'," & _
                 "'" & fechamento.GetSALARIO & "'," & _
                 "'" & fechamento.GetCOMISSAO & "'," & _
                 "'" & fechamento.GetCONDUCAO & "'," & _
                 "'" & fechamento.GetRETIRADA & "'," & _
                 "'" & fechamento.GetPROXIMOTROCO & "'," & _
                 "'" & fechamento.GetQVendas & "'," & _
                 "'" & fechamento.GetQpagamentos & "'," & _
                 "'" & fechamento.GetDESCONTO & "'," & _
                 "'" & fechamento.Getjuros & "'," & _
                 "'" & fechamento.GetSTATUS & "')"
         
    
    SQL.Execute queryFechamento
    
    SQL.FecharConexao

End Sub

Public Function alterarfechameno(id As Integer, data As Date, vendas As String, troco As String, pagamento As String, _
cartao As String, crediario As String, despesa As String, salario As String, comissao As String, _
conducao As String, retirada As String, proximoTroco As String, Qvendas As String, qPagamentos As String, _
desconto As String, juros As String, status As String) As Collection
    Dim queryFechamento As String
    

    SQL.AbrirConexao
    queryFechamento = "UPDATE FECHAMENTOS " _
    & " SET DATA_FECHAMENTO = '" & data & " ' , VENDAS = '" & vendas & " ',TROCO = '" _
    & troco & " ' , PAGAMENTOS = '" & pagamento & " ' ,  CARTAO = '" & cartao & " ' , CREDIARIO = '" _
    & crediario & " ' ,DESPESA = '" & despesa & " ' , SALARIO = '" & salario & " ' ,  COMISSAO = '" _
    & comissao & " ' , CONDUCAO = '" & conducao & " ' , RETIRADA = '" & retirada & " ' , PROXIMO_TROCO = '" _
    & proximoTroco & " ', QUANTIDADE_VENDAS = '" & Qvendas & " ',QUANTIDADE_PAGAMENTOS = '" & qPagamentos & " ', DESCONTO = '" _
    & desconto & " ' , JUROS = '" & juros & " ' , STATUS = '" & status & " '" _
    & " WHERE ID =  " & id & "  "
    SQL.GetConexao
    
    '' DATA_FECHAMENTO,VENDAS,TROCO,PAGAMENTOS,CARTAO,CREDIARIO,DESPESA, _
    SALARIO,COMISSAO,CONDUCAO,RETIRADA,PROXIMO_TROCO,QUANTIDADE_VENDAS, _
    QUANTIDADE_PAGAMENTOS,DESCONTO,JUROS,STATUS
    
    SQL.Execute queryFechamento
    
    SQL.FecharConexao
    
End Function




