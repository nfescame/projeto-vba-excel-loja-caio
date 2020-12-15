Attribute VB_Name = "VerificarData"
Public Function dataValida(data As String) As Boolean
    
    On Error GoTo dataInvalida
     Dim dataReal As Date
     
     dataReal = data
     dataValida = True
         
     Exit Function
dataInvalida:
     
     dataValida = False
    
End Function

