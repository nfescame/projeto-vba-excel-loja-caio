Attribute VB_Name = "RepositorDeVendasDiarias"
Dim rs As New ADODB.Recordset
Public Function BuscarProximoNumeroVenda(dataPesquisa As Date) As Collection
    
    Dim ultNumVenda As New Collection
    Dim vendaCoincientes As New Collection
    Dim n As VendasDiariasModelo
    Set n = New VendasDiariasModelo
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT * FROM VENDAS c WHERE c.DATA_COMPRA LIKE '%" & dataPesquisa & "%'"
    
    
    rs.Open selectCmd, SQL.GetConexao
    Do While Not rs.EOF
    
        Dim venda As VendasDiariasModelo
        Set venda = New VendasDiariasModelo

        venda.SetNVENDAS = rs("NUMERO_VENDA")

        vendaCoincientes.Add venda

        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarProximoNumeroVenda = vendaCoincientes
    
End Function
Public Function BuscarVendaPorPeriodo(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim vendaCoincientes As New Collection
    Dim selectCmd As String
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT VENDEDOR, sum(VALOR_COMPRA) as VALOR_COMPRA  " & _
    "FROM VENDAS c WHERE c.DATA_COMPRA BETWEEN " & _
    "#" & di & "# AND #" & df & "# " & _
    "GROUP BY VENDEDOR "

    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim venda As VendasDiariasModelo
        Set venda = New VendasDiariasModelo

        venda.Setvendedor = rs("VENDEDOR")
        venda.SetvalorCompra = rs("VALOR_COMPRA")

        vendaCoincientes.Add venda

        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendaPorPeriodo = vendaCoincientes
End Function

Public Function BuscarTodasVendas(dataI As Date, dataF As Date) As Collection
    
    Dim todosVendas As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM VENDAS", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim vendas As VendasDiariasModelo
        Set vendas = New VendasDiariasModelo
        
        If dataI <= rs("DATA_COMPRA") And dataF >= rs("DATA_COMPRA") Then
        
            vendas.SetId = rs("ID")
            vendas.SetDESCONTO = rs("DESCONTO")
            vendas.SetCARTAO = rs("CARTAO")
            vendas.SetDINHEIRO = rs("DINHEIRO")
            vendas.SetCREDIARIO = rs("CREDIARIO")
            vendas.SetvalorCompra = rs("VALOR_COMPRA")
            vendas.Setvendedor = rs("VENDEDOR")
            vendas.SetiDvendedor = rs("ID_VENDEDOR")
            vendas.Setcliente = rs("CLIENTE")
            vendas.SetiDcliente = rs("ID_CLIENTE")
            vendas.SetdataCompra = rs("DATA_COMPRA")
            vendas.SetQPARCELAS = rs("QUANTIDADE_PARCELAS")
            vendas.SetNORCAMENTO = rs("NUMERO_ORCAMENTO")
            vendas.SetDATAORCAMENTO = rs("DATA_ORCAMENTO")
    
            todosVendas.Add vendas
            
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarTodasVendas = todosVendas
End Function
Public Function BuscarVendasPorCliente(nome As String) As Collection
    
    Dim todosVendas As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM VENDAS", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim vendas As VendasDiariasModelo
        Set vendas = New VendasDiariasModelo
        
        If nome = rs("CLIENTE") Then
        
            vendas.SetId = rs("ID")
            vendas.SetDESCONTO = rs("DESCONTO")
            vendas.SetCARTAO = rs("CARTAO")
            vendas.SetDINHEIRO = rs("DINHEIRO")
            vendas.SetCREDIARIO = rs("CREDIARIO")
            vendas.SetvalorCompra = rs("VALOR_COMPRA")
            vendas.Setvendedor = rs("VENDEDOR")
            vendas.SetiDvendedor = rs("ID_VENDEDOR")
            vendas.Setcliente = rs("CLIENTE")
            vendas.SetiDcliente = rs("ID_CLIENTE")
            vendas.SetdataCompra = rs("DATA_COMPRA")
            vendas.SetQPARCELAS = rs("QUANTIDADE_PARCELAS")
            vendas.SetNORCAMENTO = rs("NUMERO_ORCAMENTO")
            vendas.SetDATAORCAMENTO = rs("DATA_ORCAMENTO")
    
            todosVendas.Add vendas
            
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendasPorCliente = todosVendas
End Function

Public Function BuscarVendasPorVendedor(vendedor As String, dataI As Date, dataF As Date) As Collection
    
    Dim todosVendas As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM VENDAS", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim vendas As VendasDiariasModelo
        Set vendas = New VendasDiariasModelo
        
        If vendedor = rs("VENDEDOR") Then
            If dataI <= rs("DATA_COMPRA") And dataF >= rs("DATA_COMPRA") Then
            
            vendas.SetId = rs("ID")
            vendas.SetDESCONTO = rs("DESCONTO")
            vendas.SetCARTAO = rs("CARTAO")
            vendas.SetDINHEIRO = rs("DINHEIRO")
            vendas.SetCREDIARIO = rs("CREDIARIO")
            vendas.SetvalorCompra = rs("VALOR_COMPRA")
            vendas.Setvendedor = rs("VENDEDOR")
            vendas.SetiDvendedor = rs("ID_VENDEDOR")
            vendas.Setcliente = rs("CLIENTE")
            vendas.SetiDcliente = rs("ID_CLIENTE")
            vendas.SetdataCompra = rs("DATA_COMPRA")
            vendas.SetQPARCELAS = rs("QUANTIDADE_PARCELAS")
            vendas.SetNORCAMENTO = rs("NUMERO_ORCAMENTO")
            vendas.SetDATAORCAMENTO = rs("DATA_ORCAMENTO")
        
                todosVendas.Add vendas
                
            End If
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendasPorVendedor = todosVendas
End Function

Public Function BuscarVendasPorData(dataPesquisa As Date) As Collection
    
    Dim vendasCoincientes As New Collection
    Dim selectCmd As String
    
   selectCmd = "SELECT DATA_COMPRA, SUM(CARTAO) as CARTAO, sum(CREDIARIO) as CREDIARIO , SUM(VALOR_COMPRA) as VALOR_COMPRA, SUM(DESCONTO) as DESCONTO " & _
    "FROM VENDAS c WHERE c.DATA_COMPRA LIKE '%" & dataPesquisa & "%'" & _
    "GROUP BY DATA_COMPRA"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao

    Do While Not rs.EOF
        Dim vendas As VendasDiariasModelo
        Set vendas = New VendasDiariasModelo
            
            vendas.SetdataCompra = rs("DATA_COMPRA")
            vendas.SetDESCONTO = rs("DESCONTO")
            vendas.SetCARTAO = rs("CARTAO")
            vendas.SetCREDIARIO = rs("CREDIARIO")
            vendas.SetvalorCompra = rs("VALOR_COMPRA")
           
        vendasCoincientes.Add vendas
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendasPorData = vendasCoincientes
End Function
Public Function BuscarVendasPorDataHoje(dataPesquisa As Date) As Collection
    
    Dim vendasCoincientes As New Collection
    Dim selectCmd As String
    
   selectCmd = "SELECT * FROM VENDAS c WHERE c.DATA_COMPRA LIKE '%" & dataPesquisa & "%'"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao

    Do While Not rs.EOF
        Dim vendas As VendasDiariasModelo
        Set vendas = New VendasDiariasModelo
            
            vendas.SetId = rs("ID")
            vendas.SetDESCONTO = rs("DESCONTO")
            vendas.SetCARTAO = rs("CARTAO")
            vendas.SetDINHEIRO = rs("DINHEIRO")
            vendas.SetCREDIARIO = rs("CREDIARIO")
            vendas.SetvalorCompra = rs("VALOR_COMPRA")
            vendas.Setvendedor = rs("VENDEDOR")
            vendas.SetiDvendedor = rs("ID_VENDEDOR")
            vendas.Setcliente = rs("CLIENTE")
            vendas.SetiDcliente = rs("ID_CLIENTE")
            vendas.SetdataCompra = rs("DATA_COMPRA")
            vendas.SetQPARCELAS = rs("QUANTIDADE_PARCELAS")
            vendas.SetNORCAMENTO = rs("NUMERO_ORCAMENTO")
            vendas.SetDATAORCAMENTO = rs("DATA_ORCAMENTO")
            vendas.SetNVENDAS = rs("NUMERO_VENDA")
           
        vendasCoincientes.Add vendas
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarVendasPorDataHoje = vendasCoincientes
End Function
Public Function BuscarQuantidadeDeVendasPorData(dataPesquisa As Date) As Collection
    
    Dim Qvendas As New Collection
    Dim selectCmd As String
    
    selectCmd = "SELECT * FROM VENDAS c WHERE c.DATA_COMPRA LIKE '%" & dataPesquisa & "%'"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao

    Do While Not rs.EOF
        Dim Q As VendasDiariasModelo
        Set Q = New VendasDiariasModelo
            
            Q.SetNVENDAS = rs("NUMERO_VENDA")
           
        Qvendas.Add Q
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarQuantidadeDeVendasPorData = Qvendas
End Function

Public Function BuscarIdVenda(nVendas As Integer, data As Date) As Collection
    
    Dim idCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM VENDAS c WHERE c.DATA_COMPRA LIKE '%" & data & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    Dim id As RecebimentosModelo
    Set id = New RecebimentosModelo
    
    If rs("NUMERO_VENDA") = nVendas Then
        id.SetId = rs("ID")
        idCoincientes.Add id
    End If
 
    rs.MoveNext
    
    Loop
        
    SQL.FecharConexao
    
    Set BuscarIdVenda = idCoincientes
End Function

Public Sub AdicionarVendasDiarias(vendas As VendasDiariasModelo)
    Dim queryVendas As String
    
    SQL.AbrirConexao
    
    queryVendas = "INSERT INTO VENDAS (desconto,cartao,dinheiro,crediario,valor_compra,vendedor,ID_VENDEDOR,cliente,ID_CLIENTE,data_compra,quantidade_parcelas,numero_orcamento,data_orcamento,numero_venda)" & _
        "VALUES ( '" & vendas.GetDESCONTO & "'," & _
                 "'" & vendas.GetCARTAO & "'," & _
                 "'" & vendas.GetDINHEIRO & "'," & _
                 "'" & vendas.GetCREDIARIO & "'," & _
                 "'" & vendas.GetvalorCompra & "'," & _
                 "'" & vendas.Getvendedor & "'," & _
                 "'" & vendas.GetiDvendedor & "'," & _
                 "'" & vendas.Getcliente & "'," & _
                 "'" & vendas.GetiDcliente & "'," & _
                 "'" & vendas.GetdataCompra & "'," & _
                 "'" & vendas.GetQPARCELAS & "'," & _
                 "'" & vendas.GetNORCAMENTO & "'," & _
                 "'" & vendas.GetDATAORCAMENTO & "'," & _
                 "'" & vendas.GetNVENDAS & "')"
         
    
    SQL.Execute queryVendas
    
    SQL.FecharConexao
    
End Sub

Public Function alterarVendas(id As Integer, desconto As String, cartao As String, dinheiro As String, _
crediario As String, valorCompra As String, vendedor As String, idVendedor As String, cliente As String, _
idCliente As String) As Collection
    Dim queryVendas As String
    

    SQL.AbrirConexao
    queryVendas = "UPDATE VENDAS " _
    & " SET DESCONTO = '" & desconto & " ' , CARTAO = '" & cartao & " ',DINHEIRO = '" _
    & dinheiro & " ' , CREDIARIO = '" & crediario & " ' ,  VALOR_COMPRA = '" & valorCompra & " ' , VENDEDOR = '" _
    & vendedor & " ' ,ID_VENDEDOR = '" & idVendedor & " ' , CLIENTE = '" & cliente & " ', ID_CLIENTE = '" & idCliente & " '" _
    & " WHERE ID =  " & id & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryVendas
    
    SQL.FecharConexao
    
End Function
Public Function excluirVenda(id As Integer) As Collection
    Dim queryVendas As String

    SQL.AbrirConexao
    
    queryVendas = "DELETE FROM VENDAS " _
    & " WHERE ID =  " & id & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryVendas
    
    SQL.FecharConexao
    
End Function

