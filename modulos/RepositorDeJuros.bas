Attribute VB_Name = "RepositorDeJuros"
Dim rs As New ADODB.Recordset

Public Function BuscarJurosPorId(idJuros As Integer) As Collection
    
    Dim jurosEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM JUROS c WHERE c.ID LIKE '%" & idJuros & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim juros As JurosModelo
        Set juros = New JurosModelo
        
        juros.Setjuros = rs("JUROS")
        
        jurosEncontrado.Add juros
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    Set BuscarJurosPorId = jurosEncontrado
End Function

Public Function alterarJuros(idJuros As Integer, JurosAtuatizada As String) As Collection
    Dim queryJuros As String

    SQL.AbrirConexao
    
    queryJuros = "UPDATE JUROS " _
    & " SET JUROS = '" & JurosAtuatizada & "'" _
    & " WHERE ID =  " & idJuros & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryJuros
    
    SQL.FecharConexao
    
End Function




