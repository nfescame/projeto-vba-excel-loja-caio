Attribute VB_Name = "RepositorDeLogin"
Dim rs As New ADODB.Recordset
Public Function BuscarLoginSenha() As Collection
    
    Dim todosLogins As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM LOGIN", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim log As LoginModelo
        Set log = New LoginModelo
        
            log.SetId = rs("ID")
            log.Setlogin = rs("LOGIN")
            log.Setsenha = rs("SENHA")
            log.Setoperador = rs("ACESSO")
            
            todosLogins.Add log
         
        
        rs.MoveNext
        
    Loop
    
    SQL.FecharConexao
    
    Set BuscarLoginSenha = todosLogins
End Function
Public Function buscarLogin(login As String) As Collection
    
    Dim todosLogins As New Collection
    
    SQL.AbrirConexao
    rs.Open "SELECT * FROM LOGIN c WHERE c.LOGIN LIKE '%" & login & "%'", SQL.GetConexao
   
    
    Do While Not rs.EOF
        Dim log As LoginModelo
        Set log = New LoginModelo
        
            log.SetId = rs("ID")
            log.Setlogin = rs("LOGIN")
            log.Setsenha = rs("SENHA")
            log.Setoperador = rs("ACESSO")
            
            todosLogins.Add log
         
        
        rs.MoveNext
        
    Loop
    
    SQL.FecharConexao
    
    Set buscarLogin = todosLogins
End Function


Public Sub AdicionarLogin(log As LoginModelo)
    Dim queryLogin As String
    
    SQL.AbrirConexao
     
     queryLogin = "INSERT INTO LOGIN (login,senha,acesso)" & _
        "VALUES ( '" & log.Getlogin & "'," & _
        "'" & log.Getsenha & "'," & _
        "'" & log.Getoperador & "')"
        
    
    SQL.Execute queryLogin
    
    SQL.FecharConexao
End Sub
Public Function alterarLogin(logAtual As String, senhaNova As String) As Collection
    Dim queryLogin As String

    SQL.AbrirConexao
    
    queryLogin = "UPDATE LOGIN " _
    & " SET SENHA = '" & senhaNova & "'" _
    & " WHERE LOGIN =  " & logAtual & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryLogin
    
    SQL.FecharConexao
    
End Function


Sub bloquear()
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
    , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
    :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
    AllowDeletingRows:=True
End Sub
Sub desbloquear()
    ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=False _
    , AllowFormattingCells:=False, AllowFormattingColumns:=False, _
    AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowInsertingRows _
    :=False, AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, _
    AllowDeletingRows:=False
End Sub




