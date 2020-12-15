Attribute VB_Name = "RepositorDeGrupo"
Dim rs As New ADODB.Recordset
Public Function BuscarGrupos() As Collection
    
    Dim gruposEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM GRUPO_ESTOQUE", SQL.GetConexao
    
    
    Do While Not rs.EOF
        Dim grupo As GrupoModelo
        Set grupo = New GrupoModelo
        
        grupo.SetId = rs("ID")
        grupo.Setgrupo = rs("GRUPO")
        grupo.SetSubGrupo = rs("SUB_GRUPO")
        
        gruposEncontrado.Add grupo
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    Set BuscarGrupos = gruposEncontrado
End Function
Public Function BuscarSubGrupos(grupo As String) As Collection
    
    Dim subGruposEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM GRUPO_ESTOQUE c WHERE c.GRUPO LIKE '%" & grupo & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim Sgrupo As GrupoModelo
        Set Sgrupo = New GrupoModelo
        
        Sgrupo.SetSubGrupo = rs("SUB_GRUPO")
        
        subGruposEncontrado.Add Sgrupo
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    Set BuscarSubGrupos = subGruposEncontrado
End Function

Public Sub AdicionarGrupo(grupo As GrupoModelo)
    Dim queryGrupo As String
    
    SQL.AbrirConexao
     
    queryGrupo = "INSERT INTO GRUPO_ESTOQUE (GRUPO,SUB_GRUPO)" & _
       "VALUES ( '" & grupo.Getgrupo & "'," & _
                "'" & grupo.GetSubGrupo & "')"
        
     
       
    SQL.Execute queryGrupo
    
    SQL.FecharConexao
End Sub
Public Function alterarGrupo(id As Integer, descricaoAtu As String, subGrupoAtu As String) As Collection
    Dim queryGrupo As String

    SQL.AbrirConexao
    
    queryGrupo = "UPDATE GRUPO_ESTOQUE " _
    & " SET GRUPO = '" & descricaoAtu & "',SUB_GRUPO = '" & subGrupoAtu & "'" _
    & " WHERE ID =  " & id & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryGrupo
    
    SQL.FecharConexao
    
End Function



