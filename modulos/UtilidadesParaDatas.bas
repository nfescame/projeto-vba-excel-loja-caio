Attribute VB_Name = "UtilidadesParaDatas"
Public Function getDataISO(data As Date) As String
    Dim dia, mes, ano As String
    
    dia = Day(data)
    mes = Month(data)
    ano = Year(data)
    
    getDataISO = ano & " " & mes & " " & dia
End Function
