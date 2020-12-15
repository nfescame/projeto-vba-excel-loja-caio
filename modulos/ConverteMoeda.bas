Attribute VB_Name = "ConverteMoeda"
Public Function FormataParaMoeda(valor As String) As String

   Dim val As New Collection
    If IsNumeric(valor) Then
        If InStr(1, valor, "-") >= 1 Then valor = Replace(valor, "-", "") 'retira sinal negativo
        If InStr(1, valor, ",") >= 1 Then valor = CDbl(Replace(valor, ",", "")) 'retirar a virgula
        If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "") 'para trabalhar melhor retiramos ponto
        Select Case Len(valor) 'verifica casas para inserção de ponto
            Case 1
            numPonto = "00" & valor
            Case 2
            numPonto = "0" & valor
            Case 6 To 8
            numPonto = Left(valor, Len(valor) - 5) & "." & Right(valor, 5)
            Case 9 To 11
            numPonto = inseriPonto(8, valor)
            Case 12 To 14
            numPonto = inseriPonto(11, valor)
            Case Else
            numPonto = valor
        End Select
        numVirgula = Left(numPonto, Len(numPonto) - 2) & "," & Right(numPonto, 2)
        valor = "R$ " & numVirgula
    Else
        If valor = "" Then Exit Function
        MsgBox "Número invalido", vbCritical, "Caracter Invalido"
        Exit Function
    End If
     FormataParaMoeda = valor
End Function
Function inseriPonto(inicio, valor)
    i = Left(valor, Len(valor) - inicio)
    M1 = Left(Right(valor, inicio), 3)
    M2 = Left(Right(valor, 8), 3)
    f = Right(valor, 5)
    If (M2 = M1) And (Len(valor) < 12) Then
    inseriPonto = i & "." & M1 & "." & f
    Else
    inseriPonto = i & "." & M1 & "." & M2 & "." & f
    End If
End Function
