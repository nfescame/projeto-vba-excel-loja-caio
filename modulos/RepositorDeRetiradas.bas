Attribute VB_Name = "RepositorDeRetiradas"
Dim rs As New ADODB.Recordset
Public Function BuscarRetiradasPorData(data As Date) As Collection
    Dim selectCmd As String
    Dim todosRetiradas As New Collection

    selectCmd = "SELECT SUM(RETIRADA) as RETIRADA,SUM(DESPESAS) as DESPESAS, SUM(SALARIO) as SALARIO, SUM(COMISSAO) as COMISSAO, SUM(CONDUCAO) as CONDUCAO " & _
    "FROM RETIRADAS c WHERE c.DATA_RETIRADA LIKE '%" & data & "%'" & _
    "GROUP BY DATA_RETIRADA"
    
    SQL.AbrirConexao
    
    rs.Open selectCmd, SQL.GetConexao
    
    Do While Not rs.EOF
        Dim retiradas As RetiradasModelo
        Set retiradas = New RetiradasModelo
        
        ''retiradas.SetId = rs("ID")
        ''retiradas.SetDATARETIRADA = rs("DATA_RETIRADA")
        retiradas.SetRETIRADA = rs("RETIRADA")
        retiradas.SetDESPESAS = rs("DESPESAS")
        retiradas.SetSALARIO = rs("SALARIO")
        retiradas.SetCOMISSAO = rs("COMISSAO")
        retiradas.SetCONDUCAO = rs("CONDUCAO")
       
        todosRetiradas.Add retiradas
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarRetiradasPorData = todosRetiradas
End Function

Public Sub AdicionarRetiradas(retirada As RetiradasModelo)
    Dim queryRetirada As String
    
    SQL.AbrirConexao
     
     queryRetirada = "INSERT INTO RETIRADAS (DATA_RETIRADA,RETIRADA,DESPESAS,SALARIO,COMISSAO,CONDUCAO)" & _
        "VALUES ( '" & retirada.GetDATARETIRADA & "'," & _
                    "'" & retirada.GetRETIRADA & "'," & _
                    "'" & retirada.GetDESPESAS & "'," & _
                    "'" & retirada.GetSALARIO & "'," & _
                    "'" & retirada.GetCOMISSAO & "'," & _
                    "'" & retirada.GetCONDUCAO & "')"
        
        
    SQL.Execute queryRetirada
    
    SQL.FecharConexao
End Sub

