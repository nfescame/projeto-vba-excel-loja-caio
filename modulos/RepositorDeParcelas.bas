Attribute VB_Name = "RepositorDeParcelas"
Dim rs As New ADODB.Recordset
Public Function BuscarTodosParcelas() As Collection
    
    Dim todosparcelas As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM PARCELAS", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim parcelas As ParcelasModelo
        Set parcelas = New ParcelasModelo
        
        parcelas.SetIdP = rs("ID")
        parcelas.Setparcela = rs("PARCELA")
        parcelas.SetTaxa = rs("TAXA")
       
        
        
        todosparcelas.Add parcelas
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarTodosParcelas = todosparcelas
End Function


Public Function BuscarParcelaPorDescricao(parcela As String) As Collection
    
    Dim parcelasCoincientes As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM PARCELAS c WHERE c.PARCELA LIKE '%" & parcela & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim parcelas As ParcelasModelo
        Set parcelas = New ParcelasModelo
        If parcela = rs("PARCELA") Then
        
            parcelas.SetIdP = rs("ID")
            parcelas.Setparcela = rs("PARCELA")
            parcelas.SetTaxa = rs("TAXA")
            
            parcelasCoincientes.Add parcelas
       
        End If
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarParcelaPorDescricao = parcelasCoincientes
End Function
Public Function BuscarParcelasPorId(idParcela As String) As Collection
    
    Dim idEncontrado As New Collection
    
    SQL.AbrirConexao
    
    rs.Open "SELECT * FROM PARCELAS c WHERE c.ID LIKE '%" & idParcela & "%'", SQL.GetConexao
    
    Do While Not rs.EOF
        Dim parcelas As ParcelasModelo
        Set parcelas = New ParcelasModelo
        
        parcelas.SetIdP = rs("ID")
        parcelas.Setparcela = rs("PARCELA")
        parcelas.SetTaxa = rs("TAXA")
        
        If rs("ID") = idParcela Then
        
            idEncontrado.Add parcelas
  
        End If
        
        
        rs.MoveNext
    Loop
    
    SQL.FecharConexao
    
    Set BuscarParcelasPorId = idEncontrado
End Function


Public Sub AdicionarParcelas(parcelas As ParcelasModelo)
    Dim queryParcelas As String
    
    SQL.AbrirConexao
     
     queryParcelas = "INSERT INTO PARCELAS (parcela,taxa)" & _
        "VALUES ( '" & parcelas.Getparcela & "'," & _
        "'" & parcelas.GetTaxa & "')"
        
        
    
    SQL.Execute queryParcelas
    
    SQL.FecharConexao
End Sub
Public Function alterarParcela(idParcela As Integer, parcelaAtuatizada As String, taxaAtuatizada As String) As Collection
    Dim queryParcelas As String

    SQL.AbrirConexao
    
    queryParcelas = "UPDATE PARCELAS " _
    & " SET PARCELA = '" & parcelaAtuatizada & "',TAXA = '" & taxaAtuatizada & "' " _
    & " WHERE ID =  " & idParcela & "  "
    SQL.GetConexao
    
    
    SQL.Execute queryParcelas
    
    SQL.FecharConexao
    
End Function



