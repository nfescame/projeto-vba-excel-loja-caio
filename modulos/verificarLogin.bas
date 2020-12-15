Attribute VB_Name = "verificarLogin"
Public Function buscarLog(log As String) As Boolean
   
    Dim login As New Collection
    
    Set login = RepositorDeLogin.buscarLogin(log)
   
    
End Function
