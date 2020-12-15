Attribute VB_Name = "SQL"
Dim gConexao As ADODB.Connection
Dim caminhoBancoDeDados As String

Public Function GetConexao() As ADODB.Connection
    Set GetConexao = gConexao
End Function

Public Sub Execute(query As String)
    gConexao.Execute query
End Sub

Public Sub AbrirConexao()

    Dim strConexao As String
    Set gConexao = New ADODB.Connection
    
    caminhoBancoDeDados = "C:\Users\StarShoes\OneDrive\BancoDeDados\BancoDeDadosCaio.accdb"
    
    strConexao = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & caminhoBancoDeDados & ";" & _
        "Persist Security Info=False"
    
    gConexao.Open strConexao
End Sub

Public Sub FecharConexao()
    If Not gConexao Is Nothing Then
        gConexao.Close
        Set gConexao = Nothing
    End If
End Sub


