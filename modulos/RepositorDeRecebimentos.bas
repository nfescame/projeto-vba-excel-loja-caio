Attribute VB_Name = "RepositorDeRecebimentos"
Dim rs As New ADODB.Recordset
Public Function BuscarPorPeriodoAberto(dataI As Date, dataF As Date, status As String) As Collection
    Dim selectCmd As String
    Dim DevedorCoincientes As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataI)
    df = UtilidadesParaDatas.getDataISO(dataF)
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT ID, CLIENTE, VENCIMENTO, VALOR_PARCELA,STATUS " & _
    "FROM DEBITOS c WHERE c.VENCIMENTO BETWEEN " & "#" & di & "# AND #" & df & "# "
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim recebiveis As RecebimentosModelo
        Set recebiveis = New RecebimentosModelo
        If rs("STATUS") = status Then
        recebiveis.SetId = rs("ID")
        recebiveis.SetDevedor = rs("CLIENTE")
        recebiveis.SetVencimento = rs("VENCIMENTO")
        recebiveis.SetvalorParcela = rs("VALOR_PARCELA")
        recebiveis.SetSTATUS = rs("STATUS")
        
        DevedorCoincientes.Add recebiveis
        End If
        rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set BuscarPorPeriodoAberto = DevedorCoincientes
End Function
Public Function BuscarPorPeriodoNomeAberto(dataI As Date, dataF As Date, status As String, cliente As String) As Collection
    Dim selectCmd As String
    Dim DevedorCoincientes As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataI)
    df = UtilidadesParaDatas.getDataISO(dataF)
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT ID, CLIENTE, VENCIMENTO, VALOR_PARCELA,STATUS " & _
    "FROM DEBITOS c WHERE c.VENCIMENTO BETWEEN " & "#" & di & "# AND #" & df & "# "
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim recebiveis As RecebimentosModelo
        Set recebiveis = New RecebimentosModelo
        If rs("STATUS") = status Then
            If rs("CLIENTE") = cliente Then
                recebiveis.SetId = rs("ID")
                recebiveis.SetDevedor = rs("CLIENTE")
                recebiveis.SetVencimento = rs("VENCIMENTO")
                recebiveis.SetvalorParcela = rs("VALOR_PARCELA")
                recebiveis.SetSTATUS = rs("STATUS")
                
                DevedorCoincientes.Add recebiveis
            End If
        End If
        rs.MoveNext
        Loop
        
    SQL.FecharConexao
    
    Set BuscarPorPeriodoNomeAberto = DevedorCoincientes
End Function

Public Function BuscarDevedorPorIdDebito(idDebito As String) As Collection
  
    Dim DevedorCoincientes As New Collection
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM DEBITOS c WHERE c.ID_CLIENTE LIKE '%" & idDebito & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim recebiveis As RecebimentosModelo
        Set recebiveis = New RecebimentosModelo
        
        If rs("ID_CLIENTE") = idDebito Then
        recebiveis.SetId = rs("ID")
        recebiveis.SetdataCompra = rs("DATA_COMPRA")
        recebiveis.SetValorTotal = rs("VALOR_TOTAL")
        recebiveis.SetDevedor = rs("CLIENTE")
        recebiveis.SetidDevedor = rs("ID_CLIENTE")
        recebiveis.Setvendedor = rs("VENDEDOR")
        recebiveis.SetiDvendedor = rs("ID_VENDEDOR")
        recebiveis.SetVencimento = rs("VENCIMENTO")
        recebiveis.SetvalorParcela = rs("VALOR_PARCELA")
        recebiveis.Setparcelas = rs("PARCELAS")
        recebiveis.SetSTATUS = rs("STATUS")
        End If
        If Not rs("STATUS") = "PAGO" Then
        
            DevedorCoincientes.Add recebiveis
        
        End If
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarDevedorPorIdDebito = DevedorCoincientes
End Function

Public Function BuscarCarnePrint(id As Integer, data As Date) As Collection
    
    Dim vendasCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM DEBITOS c WHERE c.NUMERO_VENDA LIKE '%" & id & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    If rs("DATA_COMPRA") = data Then
    Dim carne As RecebimentosModelo
    Set carne = New RecebimentosModelo
    
            
        carne.SetId = rs("ID")
        carne.SetdataCompra = rs("DATA_COMPRA")
        carne.SetValorTotal = rs("VALOR_TOTAL")
        carne.SetDevedor = rs("CLIENTE")
        carne.SetidDevedor = rs("ID_CLIENTE")
        carne.Setvendedor = rs("VENDEDOR")
        carne.SetiDvendedor = rs("ID_VENDEDOR")
        carne.SetVencimento = rs("VENCIMENTO")
        carne.SetvalorParcela = rs("VALOR_PARCELA")
        carne.Setparcelas = rs("PARCELAS")
        carne.SetSTATUS = rs("STATUS")
        carne.SetnumeroVenda = rs("NUMERO_VENDA")
                
       
    vendasCoincientes.Add carne
    End If
    rs.MoveNext
    
    Loop
        
    SQL.FecharConexao
    
    Set BuscarCarnePrint = vendasCoincientes
End Function
Public Function BuscarIdVenda(nVendas As Integer, data As Date) As Collection
    
    Dim vendasCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM DEBITOS c WHERE c.DATA_COMPRA LIKE '%" & data & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    Dim carne As RecebimentosModelo
    Set carne = New RecebimentosModelo
    
    If rs("NUMERO_VENDA") = nVendas Then
        carne.SetId = rs("ID")
        vendasCoincientes.Add carne
    End If
 
    rs.MoveNext
    
    Loop
        
    SQL.FecharConexao
    
    Set BuscarIdVenda = vendasCoincientes
End Function

Public Sub AdicionarRecebiveis(recebiveis As RecebimentosModelo)
    Dim queryRecebimentos As String
       
    SQL.AbrirConexao
    
    queryRecebimentos = "INSERT INTO DEBITOS (DATA_COMPRA,VALOR_TOTAL,CLIENTE,ID_CLIENTE,VENDEDOR,ID_VENDEDOR,VENCIMENTO,VALOR_PARCELA,PARCELAS, STATUS,NUMERO_VENDA,ID_VENDA)" & _
        "VALUES ( '" & recebiveis.GetdataCompra & "'," & _
                 "'" & recebiveis.GetValorTotal & "'," & _
                 "'" & recebiveis.GetDevedor & "'," & _
                 "'" & recebiveis.GetidDevedor & "'," & _
                 "'" & recebiveis.Getvendedor & "'," & _
                 "'" & recebiveis.GetiDvendedor & "'," & _
                 "'" & recebiveis.GetVencimento & "'," & _
                 "'" & recebiveis.GetvalorParcela & "'," & _
                 "'" & recebiveis.Getparcelas & "'," & _
                 "'" & recebiveis.GetSTATUS & "'," & _
                 "'" & recebiveis.GetnumeroVenda & "'," & _
                 "'" & recebiveis.GetidVenda & "')"
         
    
    SQL.Execute queryRecebimentos
    
    SQL.FecharConexao
    
End Sub
Public Function alterarStatusParcela(idParcela As Integer, statusAlt As String) As Collection
    Dim queryParcelas As String

    SQL.AbrirConexao
    
    queryParcelas = "UPDATE DEBITOS " _
    & " SET STATUS = '" & statusAlt & "' " _
    & " WHERE ID =  " & idParcela & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryParcelas
    
    SQL.FecharConexao
    
End Function
Public Function alterarValorParcela(idParcela As Integer, valor As String) As Collection
    Dim queryParcelas As String

    SQL.AbrirConexao
    
    queryParcelas = "UPDATE DEBITOS " _
    & " SET VALOR_PARCELA = '" & valor & "' " _
    & " WHERE ID =  " & idParcela & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryParcelas
    
    SQL.FecharConexao
    
End Function
Public Function excluirParcela(idParcela As Integer) As Collection
    Dim queryParcelas As String

    SQL.AbrirConexao
    
    queryParcelas = "DELETE FROM DEBITOS " _
    & " WHERE ID =  " & idParcela & "  "
    SQL.GetConexao
    
    SQL.Execute queryParcelas
    
    SQL.FecharConexao
    
End Function



