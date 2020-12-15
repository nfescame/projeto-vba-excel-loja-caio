Attribute VB_Name = "RepositorDeRetiradasDetalhadas"
Dim rs As New ADODB.Recordset
Public Function BuscarRetiradasDetalhadasIdent(dataIPesquisa As Date, datafPesquisa As Date) As Collection
    Dim selectCmd As String
    Dim todosRetiradas As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT IDENTIFICACAO, SUM(VALOR) as VALOR " & _
    "FROM RETIRADAS_DETALHADAS c WHERE c.DATA BETWEEN " & _
    "#" & di & "# " & "AND #" & df & "# " & _
    "GROUP BY IDENTIFICACAO"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim retiradas As RetiradasDetalhadasModelo
        Set retiradas = New RetiradasDetalhadasModelo
        
        ''retiradas.SetId = rs("ID")
        retiradas.Setidentificacao = rs("IDENTIFICACAO")
        retiradas.Setvalor = rs("VALOR")
        ''retiradas.Setdescricao = rs("DESCRICAO")
        ''retiradas.Setdata = rs("DATA")
       
        todosRetiradas.Add retiradas
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarRetiradasDetalhadasIdent = todosRetiradas
End Function
Public Function BuscarRetiradasDetalhadasDescricao(ident As String, dataIPesquisa As Date, datafPesquisa As Date) As Collection
    Dim selectCmd As String
    Dim todosRetiradas As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT DESCRICAO, SUM(VALOR) as VALOR,IDENTIFICACAO " & _
    "FROM RETIRADAS_DETALHADAS c WHERE c.DATA BETWEEN " & _
    "#" & di & "# " & "AND #" & df & "# " & _
    "GROUP BY IDENTIFICACAO,DESCRICAO"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
    
        Dim retiradas As RetiradasDetalhadasModelo
        Set retiradas = New RetiradasDetalhadasModelo
        
        If rs("IDENTIFICACAO") = ident Then
             ''retiradas.SetId = rs("ID")
             retiradas.Setidentificacao = rs("IDENTIFICACAO")
             retiradas.Setvalor = rs("VALOR")
             retiradas.Setdescricao = rs("DESCRICAO")
             ''retiradas.Setdata = rs("DATA")
            
             todosRetiradas.Add retiradas
        End If
        
        rs.MoveNext
   
    Loop
    
    SQL.FecharConexao
    
    Set BuscarRetiradasDetalhadasDescricao = todosRetiradas
End Function
Public Function BuscarRetiradasDetalhadasData(identificacao As String, descricao As String, dataIPesquisa As Date, datafPesquisa As Date) As Collection
    Dim selectCmd As String
    Dim todosRetiradas As New Collection
    
    di = UtilidadesParaDatas.getDataISO(dataIPesquisa)
    df = UtilidadesParaDatas.getDataISO(datafPesquisa)
    
    selectCmd = "SELECT SUM(VALOR)AS VALOR,IDENTIFICACAO,DESCRICAO,DATA " & _
    "FROM RETIRADAS_DETALHADAS c WHERE c.DATA BETWEEN " & _
    "#" & di & "# " & "AND #" & df & "# " & _
    "GROUP BY IDENTIFICACAO,DESCRICAO,DATA"
    

    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
    
        Dim retiradas As RetiradasDetalhadasModelo
        Set retiradas = New RetiradasDetalhadasModelo
        
        If rs("IDENTIFICACAO") = identificacao Then
            If rs("DESCRICAO") = descricao Then
            
                 retiradas.Setidentificacao = rs("IDENTIFICACAO")
                 retiradas.Setvalor = rs("VALOR")
                 retiradas.Setdescricao = rs("DESCRICAO")
                 retiradas.Setdata = rs("DATA")
                
                 todosRetiradas.Add retiradas
                 
            End If
        End If
        
        rs.MoveNext
   
    Loop
    
    SQL.FecharConexao
    
    Set BuscarRetiradasDetalhadasData = todosRetiradas
End Function

Public Sub AdicionarRetiradasDetalhadas(retirada As RetiradasDetalhadasModelo)
    Dim queryRetiradaDetalhadas As String
    
    SQL.AbrirConexao
     
     queryRetiradaDetalhadas = "INSERT INTO RETIRADAS_DETALHADAS (IDENTIFICACAO,VALOR,DESCRICAO,DATA)" & _
        "VALUES ( '" & retirada.Getidentificacao & "'," & _
                    "'" & retirada.Getvalor & "'," & _
                    "'" & retirada.Getdescricao & "'," & _
                    "'" & retirada.Getdata & "')"
        
        
    SQL.Execute queryRetiradaDetalhadas
    
    SQL.FecharConexao
End Sub


