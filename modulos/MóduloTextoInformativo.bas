Attribute VB_Name = "MóduloTextoInformativo"
Public Function TextInfo(Button As String) As String

Select Case Button
Case "CADASTRO"
TextInfo = "OPCÕES PARA CADASTRO." & vbNewLine & _
"1.Cadastro De Cliente: Cadastra dados dos clientes." & vbNewLine & _
"2.Cadastro De Produtos: Cadastra dados dos produtos." & vbNewLine & _
"3.Cadastro De Parcela: Cadastra e Altera Parcela,Taxa Da Parcela e Juros Por Atrazo." & vbNewLine & _
"4.Cadastro De Grupo: Cadastra e altera grupo e subgrupo para produtos." & vbNewLine & _
"5.Cadastro De Vendedor: Cadastra e altera vendedores." & vbNewLine & _
"6.Cadastro De Dados Da Empresa: Cadastra e altera dados da empresa para impressoes em geral." & vbNewLine & _
"7.Cadastro De Debitos Extra: Cadastra debitos sem venda relacionada." & vbNewLine & _
"8.Cadastro De Login: Cadastra login para acesso(ADM ou CAIXA)."

Case "PESQUISA"
TextInfo = "OPCÕES PARA CONSULTAS" & vbNewLine & _
"1.Pesquisa De Cliente: Pesquisa e altera dados do cliente e visualiza compras detalhadas do cliente." & vbNewLine & _
"2.Pesquisa De Produtos: Pesquisa e altera dado do produto." & vbNewLine & _
"3.Relatorio De Vendas: Pesquisa vendas detalhadas por vendedor,descricao,grupo e subgrupo." & vbNewLine & _
"4.Relatorio De Debitos: Pesquisa de debitos(PAGOS ou VENCIDOS)de clientes por periodo de data." & vbNewLine & _
"5.Pesquisa De Retiradas: Pesquisa de retiradas,salarios,comissoes,conducoes e despesas em geral." & vbNewLine & _
"6.Pesquisa De Orcamentos: Pesquisa,cancela e reimprime orcamentos do dia." & vbNewLine & _
"7.Pesquisa De Vendas: Pesquisa e cancelamento de vendas do dia." & vbNewLine & _
"8.Pesquisa De Pagamentos: Pesquisa e alteracao de pagamentos do dia."

Case "MOVIMENTACAO"
TextInfo = "OPCÕES PARA MOVIMENTACÃO" & vbNewLine & _
"1.Orcamento: Associa vendedor,cliente e produtos em uma pre venda." & vbNewLine & _
"2.Venda: Seleciona um orcamento pre definido e associa uma forma de pagamento para registro de venda." & vbNewLine & _
"3.Pagamento: Seleciona debitos de clientes e associa uma forma de pagamento para registro de pagamento." & vbNewLine & _
"4.Fechamento: Seleciona todas as movimentacões do dia e calcula automaticamente saldo de caixa(TROCO FINAL)." & vbNewLine & _
"5.Retiradas: Registra retiradas,salarios,comissoes,conducoes e despesas em geral." & vbNewLine & _
"6.Abertura: Recebe troco final da ultima movimentacão para abertura de caixa atual,permite alteracão para ajuste de valor."
End Select
End Function
