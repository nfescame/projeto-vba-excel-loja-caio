Attribute VB_Name = "RepositorDeOrcamento"
Dim rs As New ADODB.Recordset
Public Function BuscarOrcamentoTotalItens(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
   
    selectCmd = "SELECT ID_PRODUTO, PRODUTO, SUM(QUANTIDADE) as QUANTIDADE, VALOR_UNITARIO, GRUPO ,STATUS " & _
    "FROM ORCAMENTO c WHERE c.DATA_REGISTRO BETWEEN " & _
    "#" & di & "# " & "AND #" & df & "# " & _
    "GROUP BY PRODUTO, ID_PRODUTO, VALOR_UNITARIO, GRUPO, STATUS"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
   
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        If rs("STATUS") = "FECHADO" Then
        
        orcamento.Setproduto = rs("PRODUTO")
        orcamento.SetiDproduto = rs("ID_PRODUTO")
        orcamento.Setquantidade = rs("QUANTIDADE")
        orcamento.SetvalorUnit = rs("VALOR_UNITARIO")
        orcamento.Setgrupo = rs("GRUPO")
        orcamento.SetSTATUS = rs("STATUS")
        
        OrcamentoCoincientes.Add orcamento
        End If
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarOrcamentoTotalItens = OrcamentoCoincientes
End Function
Public Function BuscarOrcamentoPorPeriodo(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
   
    selectCmd = "SELECT ID_PRODUTO, PRODUTO, SUM(QUANTIDADE) as QUANTIDADE, VALOR_UNITARIO, GRUPO ,STATUS " & _
    "FROM ORCAMENTO c WHERE c.DATA_REGISTRO BETWEEN " & _
    "#" & di & "# " & "AND #" & df & "# " & _
    "GROUP BY PRODUTO, ID_PRODUTO, VALOR_UNITARIO, GRUPO, STATUS"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
   
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        If rs("STATUS") = "FECHADO" Then
        
        orcamento.Setproduto = rs("PRODUTO")
        orcamento.SetiDproduto = rs("ID_PRODUTO")
        orcamento.Setquantidade = rs("QUANTIDADE")
        orcamento.SetvalorUnit = rs("VALOR_UNITARIO")
        orcamento.Setgrupo = rs("GRUPO")
        orcamento.SetSTATUS = rs("STATUS")
        
        OrcamentoCoincientes.Add orcamento
        End If
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarOrcamentoPorPeriodo = OrcamentoCoincientes
End Function
Public Function BuscarPorDataDoOrcamentoGrupo(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT GRUPO, SUM(QUANTIDADE) as QUANTIDADE, SUM(VALOR_UNITARIO) AS VALOR_UNITARIO,STATUS " & _
    "FROM ORCAMENTO c WHERE c.DATA_REGISTRO BETWEEN " & _
    "#" & di & "# AND #" & df & "# " & _
    "GROUP BY GRUPO,STATUS"

    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Debug.Print selectCmd
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        If rs("STATUS") = "FECHADO" Then
        orcamento.Setquantidade = rs("QUANTIDADE")
        orcamento.Setgrupo = rs("GRUPO")
        orcamento.SetvalorUnit = rs("VALOR_UNITARIO")

        OrcamentoCoincientes.Add orcamento
        End If
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarPorDataDoOrcamentoGrupo = OrcamentoCoincientes
End Function
Public Function BuscarPorDataDoOrcamentoSubGrupo(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT SUB_GRUPO, SUM(QUANTIDADE) as QUANTIDADE, SUM(VALOR_UNITARIO) AS VALOR_UNITARIO,STATUS " & _
    "FROM ORCAMENTO c WHERE c.DATA_REGISTRO BETWEEN " & _
    "#" & di & "# AND #" & df & "# " & _
    "GROUP BY SUB_GRUPO,STATUS"

    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Debug.Print selectCmd
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        If rs("STATUS") = "FECHADO" Then

        orcamento.Setquantidade = rs("QUANTIDADE")
        orcamento.SetSubGrupo = rs("SUB_GRUPO")
        orcamento.SetvalorUnit = rs("VALOR_UNITARIO")

        OrcamentoCoincientes.Add orcamento
        End If
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarPorDataDoOrcamentoSubGrupo = OrcamentoCoincientes
End Function
Public Function BuscarSubGrupoPrint(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT SUB_GRUPO, SUM(QUANTIDADE) as QUANTIDADE,STATUS " & _
    "FROM ORCAMENTO c WHERE c.DATA_REGISTRO BETWEEN " & _
    "#" & di & "# AND #" & df & "# " & _
    "GROUP BY SUB_GRUPO,STATUS"

    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Debug.Print selectCmd
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        If rs("STATUS") = "FECHADO" Then
        orcamento.Setquantidade = rs("QUANTIDADE")
        orcamento.SetSubGrupo = rs("SUB_GRUPO")

        OrcamentoCoincientes.Add orcamento
        End If
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarSubGrupoPrint = OrcamentoCoincientes
End Function

Public Function BuscarProximoNumeroOrcamento(data As Date) As Collection
    
    Dim ultNumOrc As New Collection
    Dim n As OrcamentoModelo
    Set n = New OrcamentoModelo
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT * FROM ORCAMENTO c WHERE c.NUMERO_ORCAMENTO "
    
    rs.CursorType = adOpenKeyset
    rs.Open selectCmd, SQL.GetConexao
    On Error Resume Next
    rs.MoveLast
    troco = rs.AbsolutePosition
    
    If rs("DATA") = data Then
    If Not rs.EOF Then
    Do While Not rs.EOF
    DoEvents
    
    n.Setnumero = rs("NUMERO_ORCAMENTO")
    
    rs.MoveNext
    Loop
    End If
    End If
    SQL.FecharConexao
    ultNumOrc.Add n
    Set BuscarProximoNumeroOrcamento = ultNumOrc
End Function
Public Function buscarPorIdEDataDoOrcamento(idpesquisa As Integer, data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    selectCmd = "SELECT ID_PRODUTO, PRODUTO, SUM(QUANTIDADE) as QUANTIDADE, VALOR_UNITARIO, STATUS,VENDEDOR,NUMERO_ORCAMENTO,DATA, ID_VENDEDOR,VALOR_ORCAMENTO,CLIENTE,ID_CLIENTE " & _
    "FROM ORCAMENTO c WHERE c.DATA LIKE '%" & data & "%'" & _
    "GROUP BY PRODUTO, ID_PRODUTO, VALOR_UNITARIO, STATUS,VENDEDOR,NUMERO_ORCAMENTO,DATA, ID_VENDEDOR,VALOR_ORCAMENTO,CLIENTE,ID_CLIENTE"
    
    Debug.Print selectCmd
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        
        
           
                If rs("NUMERO_ORCAMENTO") = idpesquisa Then
                    
                    
                    orcamento.Setvendedor = rs("VENDEDOR")
                    orcamento.SetiDvendedor = rs("ID_VENDEDOR")
                    orcamento.Setproduto = rs("PRODUTO")
                    orcamento.SetiDproduto = rs("ID_PRODUTO")
                    orcamento.Setquantidade = rs("QUANTIDADE")
                    orcamento.SetvalorUnit = rs("VALOR_UNITARIO")
                    orcamento.SetSTATUS = rs("STATUS")
                    orcamento.Setnumero = rs("NUMERO_ORCAMENTO")
                    orcamento.Setdata = rs("DATA")
                    orcamento.Setvalor = rs("VALOR_ORCAMENTO")
                    orcamento.Setcliente = rs("CLIENTE")
                    orcamento.SetiDcliente = rs("ID_CLIENTE")
                    orcamento.SetSTATUS = rs("STATUS")
                    
                    OrcamentoCoincientes.Add orcamento
                    
                    
                End If
           
      
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set buscarPorIdEDataDoOrcamento = OrcamentoCoincientes
End Function
Public Function buscarPorNOrcDataDoOrcamento(n As Integer, data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    selectCmd = "SELECT * FROM ORCAMENTO c WHERE c.DATA_REGISTRO LIKE '%" & data & "%'"
    
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        
            If rs("NUMERO_ORCAMENTO") = n Then
                    
                orcamento.SetId = rs("ID")
                    
                OrcamentoCoincientes.Add orcamento
      
            End If
           
      
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set buscarPorNOrcDataDoOrcamento = OrcamentoCoincientes
End Function


Public Function PesquisarPorIdEDataDoOrcamento(idpesquisa As Integer, data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    selectCmd = "SELECT ID_PRODUTO, PRODUTO, SUM(QUANTIDADE) as QUANTIDADE, VALOR_UNITARIO, STATUS,VENDEDOR,NUMERO_ORCAMENTO,DATA, ID_VENDEDOR,VALOR_ORCAMENTO,CLIENTE,ID_CLIENTE,GRUPO,SUB_GRUPO " & _
    "FROM ORCAMENTO c WHERE c.DATA LIKE '%" & data & "%'" & _
    "GROUP BY PRODUTO, ID_PRODUTO, VALOR_UNITARIO, STATUS,VENDEDOR,NUMERO_ORCAMENTO,DATA, ID_VENDEDOR,VALOR_ORCAMENTO,CLIENTE,ID_CLIENTE,GRUPO,SUB_GRUPO "

    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo
        
        
           
                If rs("NUMERO_ORCAMENTO") = idpesquisa Then
                    If rs("STATUS") = "ABERTO" Then
                    
                    
                    orcamento.Setvendedor = rs("VENDEDOR")
                    orcamento.SetiDvendedor = rs("ID_VENDEDOR")
                    orcamento.Setdata = rs("DATA")
                    orcamento.Setproduto = rs("PRODUTO")
                    orcamento.SetiDproduto = rs("ID_PRODUTO")
                    orcamento.Setgrupo = rs("GRUPO")
                    orcamento.SetSubGrupo = rs("SUB_GRUPO")
                    orcamento.Setquantidade = rs("QUANTIDADE")
                    orcamento.SetvalorUnit = rs("VALOR_UNITARIO")
                    orcamento.SetSTATUS = rs("STATUS")
                    orcamento.Setvalor = rs("VALOR_ORCAMENTO")
                    orcamento.Setnumero = rs("NUMERO_ORCAMENTO")
                    orcamento.Setcliente = rs("CLIENTE")
                    orcamento.SetiDcliente = rs("ID_CLIENTE")
                    
                    OrcamentoCoincientes.Add orcamento
                    
                    End If
                End If
           
      
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set PesquisarPorIdEDataDoOrcamento = OrcamentoCoincientes
End Function

Public Function produtosParaPrint(idOrcamento As Integer, data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    Dim selectCmd As String
    
    selectCmd = "SELECT NUMERO_ORCAMENTO, ID_PRODUTO, PRODUTO, SUM(QUANTIDADE) as QUANTIDADE,VALOR_UNITARIO " & _
    "FROM ORCAMENTO c WHERE c.DATA LIKE '%" & data & "%'" & _
    "GROUP BY NUMERO_ORCAMENTO, ID_PRODUTO, PRODUTO, QUANTIDADE,VALOR_UNITARIO"

    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim orcamento As OrcamentoModelo
        Set orcamento = New OrcamentoModelo

                If rs("NUMERO_ORCAMENTO") = idOrcamento Then
                
                    orcamento.Setproduto = rs("PRODUTO")
                    orcamento.SetiDproduto = rs("ID_PRODUTO")
                    orcamento.Setquantidade = rs("QUANTIDADE")
                    orcamento.Setvalor = rs("VALOR_UNITARIO")
                    
                    OrcamentoCoincientes.Add orcamento
                    
                End If
      
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set produtosParaPrint = OrcamentoCoincientes
End Function

Public Function BuscarTodosOrcamentos() As Collection
    
    Dim OrcamentoCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ORCAMENTO c WHERE c.STATUS LIKE '%" & "ABERTO" & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    Dim orcamento As OrcamentoModelo
    Set orcamento = New OrcamentoModelo
    
    
    
        orcamento.SetId = rs("ID")
        orcamento.Setvendedor = rs("VENDEDOR")
        orcamento.SetiDvendedor = rs("ID_VENDEDOR")
        orcamento.Setdata = rs("DATA")
        orcamento.Setproduto = rs("PRODUTO")
        orcamento.SetiDproduto = rs("ID_PRODUTO")
        orcamento.Setgrupo = rs("GRUPO")
        orcamento.SetSubGrupo = rs("SUB_GRUPO")
        orcamento.Setquantidade = rs("QUANTIDADE")
        orcamento.SetvalorUnit = rs("VALOR_UNITARIO")
        orcamento.SetSTATUS = rs("STATUS")
        orcamento.Setvalor = rs("VALOR_ORCAMENTO")
        orcamento.Setnumero = rs("NUMERO_ORCAMENTO")
        orcamento.Setcliente = rs("CLIENTE")
        orcamento.SetiDcliente = rs("ID_CLIENTE")
                
           
        OrcamentoCoincientes.Add orcamento
        
    
    rs.MoveNext
    Loop
        
    SQL.FecharConexao
    
    Set BuscarTodosOrcamentos = OrcamentoCoincientes
End Function
Public Function BuscarOrcamentosParaBaixa(n As Integer, data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ORCAMENTO c WHERE c.DATA LIKE '%" & data & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    Dim orcamento As OrcamentoModelo
    Set orcamento = New OrcamentoModelo
    
    If rs("NUMERO_ORCAMENTO") = n Then
    
        orcamento.SetId = rs("ID")
 
        OrcamentoCoincientes.Add orcamento
        
    End If
    
    rs.MoveNext
    Loop
        
    SQL.FecharConexao
    
    Set BuscarOrcamentosParaBaixa = OrcamentoCoincientes
End Function
Public Function BuscarOrcamentosGeral(data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    
    SQL.AbrirConexao
    
    selectCmd = "SELECT VENDEDOR,DATA,STATUS,VALOR_ORCAMENTO,NUMERO_ORCAMENTO " & _
    "FROM ORCAMENTO c WHERE c.DATA LIKE '%" & data & "%'" & _
    "GROUP BY VENDEDOR,DATA,STATUS,VALOR_ORCAMENTO,NUMERO_ORCAMENTO "
    
    rs.Open selectCmd, SQL.GetConexao
     
    Do While Not rs.EOF
    Dim orcamento As OrcamentoModelo
    Set orcamento = New OrcamentoModelo
    
        orcamento.Setvendedor = rs("VENDEDOR")
        orcamento.Setdata = rs("DATA")
        orcamento.SetSTATUS = rs("STATUS")
        orcamento.Setvalor = rs("VALOR_ORCAMENTO")
        orcamento.Setnumero = rs("NUMERO_ORCAMENTO")
       
        OrcamentoCoincientes.Add orcamento
           
    rs.MoveNext
    Loop
        
    SQL.FecharConexao
    
    Set BuscarOrcamentosGeral = OrcamentoCoincientes
End Function
Public Function BuscarIdOrcAlterado(id As Integer) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ORCAMENTO c WHERE c.NUMERO_ORCAMENTO LIKE '%" & id & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    Dim orcamento As OrcamentoModelo
    Set orcamento = New OrcamentoModelo
    
    '
    
        orcamento.SetId = rs("ID")
          
        OrcamentoCoincientes.Add orcamento
        
    
    rs.MoveNext
    Loop
        
    SQL.FecharConexao
    
    Set BuscarIdOrcAlterado = OrcamentoCoincientes
End Function
Public Function BuscarOrcamentosPrint(id As Integer, data As Date) As Collection
    
    Dim OrcamentoCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM ORCAMENTO c WHERE c.NUMERO_ORCAMENTO LIKE '%" & id & "%'", SQL.GetConexao
    
    
    Do While Not rs.EOF
    Dim orcamento As OrcamentoModelo
    Set orcamento = New OrcamentoModelo
    
    If rs("DATA") = data Then
        orcamento.SetId = rs("ID")
        orcamento.Setvendedor = rs("VENDEDOR")
        orcamento.SetiDvendedor = rs("ID_VENDEDOR")
        orcamento.Setdata = rs("DATA")
        orcamento.Setvalor = rs("VALOR_ORCAMENTO")
        orcamento.Setnumero = rs("NUMERO_ORCAMENTO")
        orcamento.Setcliente = rs("CLIENTE")
        orcamento.SetiDcliente = rs("ID_CLIENTE")

        OrcamentoCoincientes.Add orcamento
        
    End If
    rs.MoveNext
    Loop
        
    SQL.FecharConexao
    
    Set BuscarOrcamentosPrint = OrcamentoCoincientes
End Function

Public Sub AdicionarOrcamento(orcamento As OrcamentoModelo)
    Dim queryOrcamento As String
       
    SQL.AbrirConexao
    
    queryOrcamento = "INSERT INTO ORCAMENTO (VENDEDOR,ID_VENDEDOR,DATA,PRODUTO,ID_PRODUTO,GRUPO,SUB_GRUPO,QUANTIDADE,VALOR_UNITARIO,STATUS,VALOR_ORCAMENTO,NUMERO_ORCAMENTO,CLIENTE,ID_CLIENTE,DATA_REGISTRO)" & _
        "VALUES ( '" & orcamento.Getvendedor & "'," & _
                 "'" & orcamento.GetiDvendedor & "'," & _
                 "'" & orcamento.Getdata & "'," & _
                 "'" & orcamento.Getproduto & "'," & _
                 "'" & orcamento.GetiDproduto & "'," & _
                 "'" & orcamento.Getgrupo & "'," & _
                 "'" & orcamento.GetSubGrupo & "'," & _
                 "'" & orcamento.Getquantidade & "'," & _
                 "'" & orcamento.GetvalorUnit & "'," & _
                 "'" & orcamento.GetSTATUS & "'," & _
                 "'" & orcamento.Getvalor & "'," & _
                 "'" & orcamento.Getnumero & "'," & _
                 "'" & orcamento.Getcliente & "'," & _
                 "'" & orcamento.GetiDcliente & "'," & _
                 "'" & orcamento.GetdataRegistro & "')"
         
    
    SQL.Execute queryOrcamento
    
    SQL.FecharConexao
    
End Sub
Public Function alterarStatusOrcamento(idpesquisa As Integer, statusAlt As String, dataRegistroAtu As Date) As Collection
    Dim queryOrcamento As String
    
    
    SQL.AbrirConexao
    
    queryOrcamento = "UPDATE ORCAMENTO " _
    & " SET STATUS = '" & statusAlt & "', DATA_REGISTRO ='" & dataRegistroAtu & " '" _
    & " WHERE ID =  " & idpesquisa & "  "
    SQL.GetConexao
   
    
    SQL.Execute queryOrcamento
    
    SQL.FecharConexao
    
End Function
Public Function excluirOrcamento(idpesquisa As Integer) As Collection
    Dim queryOrcamento As String

    SQL.AbrirConexao
    
    queryOrcamento = "DELETE FROM ORCAMENTO " _
    & " WHERE ID =  " & idpesquisa & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryOrcamento
    
    SQL.FecharConexao
    
End Function




