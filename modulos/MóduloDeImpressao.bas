Attribute VB_Name = "MóduloDeImpressao"
Dim porta As String

'variaveis dados do carne ---------------------------------
Dim quantidade, valor, cliente, idCliente, parcelas As String
Dim produtos As New Collection
Dim data As String
Dim vencimento As Date

'cariaveis para fechamento------------------------------------
Dim vendas, troco, cartao, crediario, despesa, conducao, comissao, salario, retirada, proximoTroco, Qvendas, Qpg, desconto, juros As String
Dim idFechamento As Integer
Dim pg As String

'variaveis para pagamento
Dim id As String
Dim dinheiro As String
Dim idPagador As String
Dim parcela As String
Dim totalPg As String
Dim Djuros As Double
Dim Dvalor As Double


'variaveis dados da empresa ------------------------------------
Dim nomeDaEmpresa, endereco, numero, bairro, cep, cidade, telefone, celular, email, textoOrc, textoCarne As String

'variaveis dados do orcamento ---------------------------------
Dim total As String, nOrc, vendedor, idVendedor, produto, idProduto, valorOrc, valorUnit As String

'*********************************************************************
'                   RELATORIO DE VENDAS POR VENDEDOR
'*********************************************************************
Public Function relatorioDeVendasPrint(data As Date)
porta = "\\DESKTOP-AV8P0BQ\EPSON"
Dim todasVendas As New Collection
        
Set todasVendas = RepositorDeVendasDiarias.BuscarVendaPorPeriodo(data, data)
        
imprimirRelatorioVendas todasVendas
End Function

Sub imprimirRelatorioVendas(lista As Collection)
Dim vendas As New VendasDiariasModelo

Open porta For Output Access Write As #1

Print #1, Tab(16); "CAIO CALCADOS"
Print #1, Tab(0); " "
Print #1, Tab(0); "RELATORIO DE VENDAS POR VENDEDORES"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); Date; " Hora: " & Time;

For Each vendas In lista
    Print #1, Tab(0); "------------------------------------------------";
    Print #1, Tab(2); vendas.Getvendedor; Tab(15); Format(vendas.GetvalorCompra, "currency");
Next vendas

Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Close #1
End Sub

'*********************************************************************
'                   RELATORIO DE PAGAMENTO
'*********************************************************************
Public Function relatorioDePagamentosPrint(data As Date)
porta = "\\DESKTOP-AV8P0BQ\EPSON"
Dim dados As New Collection
    
Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
preencherDadosDaEmpresa dados

Dim dadosRelatorioPagamento As New Collection
    
Set dadosRelatorioPagamento = RepositorDePagamentos.RelatorioPagamentosPrint(data)
    
imprimirRelatorioPg dadosRelatorioPagamento
End Function

Sub preencherDadosDaEmpresa(lista As Collection)

Dim dados As DadosDaEmpresaModelo
For Each dados In lista
    
    nomeDaEmpresa = dados.GetnomeEmpresa
    endereco = dados.Getendereco
    numero = dados.Getnumero
    cep = dados.Getcep
    bairro = dados.Getbairro
    cidade = dados.Getcidade
    telefone = dados.Gettelefone
    celular = dados.Getcelular
    email = dados.Getemail
    textoOrc = dados.GettextoOrc
    textoCarne = dados.GettextoCarne
        
Next dados

End Sub

Sub imprimirRelatorioPg(lista As Collection)
Dim relatPg As New PagamentoModelo

Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(0); "RELATORIO DE PAGAMENTO DE CREDIARIO"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); Date; " Hora: " & Time;
Print #1, Tab(0); "------------------------------------------------";

For Each relatPg In lista
    Print #1, Tab(0); "------------------------------------------------";
    Print #1, Tab(2); relatPg.Getcliente;
    Print #1, Tab(2); relatPg.GetdataVencimento; Tab(15); Format(relatPg.GetvalorPg, "currency");
    Print #1, Tab(0); "------------------------------------------------";
Next relatPg

Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Close #1
End Sub

'*********************************************************************
'                   RECIBO DE PAGAMENTO
'*********************************************************************
Public Function ReciboDePagamentoPrint(id As Integer)
porta = "\\DESKTOP-AV8P0BQ\EPSON"
Dim dados As New Collection
    
Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
preencherDadosDaEmpresa dados

Dim dadosPagamento As New Collection
    
Set dadosPagamento = RepositorDePagamentos.BuscarPagamentosPrint(id)
    
preencherDadosPagamentos dadosPagamento

imprimirPagamanto
End Function

Sub preencherDadosPagamentos(lista As Collection)
Dim dados As PagamentoModelo
For Each dados In lista

    id = dados.GetId
    cliente = dados.Getcliente
    idPagador = dados.GetiDcliente
    valor = dados.GetvalorPg
    parcela = dados.Getparcela
    data = dados.GetdataPagamento
    vencimento = dados.GetdataVencimento
    cartao = dados.GetpgCartao
    dinheiro = dados.GetpgDinheiro
    juros = dados.Getjuros
    desconto = dados.GetDESCONTO
    
    Djuros = dados.Getjuros
    Dvalor = dados.GetvalorPg
    totalPg = Djuros + Dvalor
Next dados
End Sub

Sub imprimirPagamanto()


Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(2); "Rua: " + endereco + numero;
Print #1, Tab(2); "Bairro: " + bairro;
Print #1, Tab(2); "CEP: " + cep;
Print #1, Tab(2); "Cidade: " + cidade;
Print #1, Tab(2); "Tel: " + telefone;
Print #1, Tab(2); "Wathsapp: " + celular;
Print #1, Tab(2); "E-Mail: " + email
Print #1, Tab(0); " "
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(10); "PAGAMENTO DE CREDIARIO"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " " + data; " Hora: " & Time;
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "CLIENTE: " + cliente;
Print #1, Tab(2); "ID: " + idPagador
Print #1, Tab(2); "ID PAGAMENTO: " + id
Print #1, Tab(0); " "
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "VALOR:          VENCIMENTO:   PARCELA: ";
Print #1, Tab(2); Format(valor, "CURRENCY"); Tab(18); vencimento; Tab(32); parcela
Print #1, Tab(0); " "
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "JUROS: " + Format(juros, "CURRENCY")
Print #1, Tab(0); " "
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(23); "VALOR RECEBIDO: " + Format(totalPg, "CURRENCY")
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Chr$(&H1D); "V"; Chr$(66); Chr$(0);
Close #1

End Sub

'*********************************************************************
'                   RELATORIO DE FECHAMENTO PARA TEXT
'*********************************************************************
Public Function fechamentoPrint(id As Integer)
porta = "C:\Users\StarShoes\Desktop\fechamento.txt"
Dim dados As New Collection
    
Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
preencherDadosDaEmpresa dados

Dim dadosFechamento As New Collection
    
Set dadosFechamento = RepositorDeFechamento.BuscarFechamentosPrint(id)
    
preencherDadosFechamento dadosFechamento

Dim todasVendas As New Collection
        
Set todasVendas = RepositorDeVendasDiarias.BuscarVendaPorPeriodo(Date, Date)

Dim SGQuan As New Collection
        
Set SGQuan = RepositorDeOrcamento.BuscarSubGrupoPrint(Date, Date)

Call imprimirFechamento(todasVendas, SGQuan)
End Function

Sub preencherDadosFechamento(lista As Collection)
Dim dadosFechamento As New Collection
Dim dados As FechamentoModelo

For Each dados In lista
    data = dados.GetdataFechamento
    vendas = dados.Getvendas
    troco = dados.GetTROCO
    pg = dados.Getpagamentos
    cartao = dados.GetCARTAO
    crediario = dados.GetCREDIARIO
    despesa = dados.GetDESPESA
    comissao = dados.GetCOMISSAO
    conducao = dados.GetCONDUCAO
    salario = dados.GetSALARIO
    retirada = dados.GetRETIRADA
    proximoTroco = dados.GetPROXIMOTROCO
    Qvendas = dados.GetQVendas
    Qpg = dados.GetQpagamentos
    desconto = dados.GetDESCONTO
    juros = dados.Getjuros
Next dados

End Sub

Sub imprimirFechamento(lista1 As Collection, lista2 As Collection)

Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(0); "FECHAMENTO"
Print #1, Tab(0); " " + data; " Hora: " & Time;
Print #1, Tab(0); " "
Print #1, Tab(2); "       VENDAS: " + Format(vendas, "CURRENCY")
Print #1, Tab(2); "        TROCO: " + Format(troco, "CURRENCY")
Print #1, Tab(2); "    PAGAMENTO: " + Format(pg, "CURRENCY")
Print #1, Tab(2); "       CARTAO: " + Format(cartao, "CURRENCY")
Print #1, Tab(2); "    CREDIARIO: " + Format(crediario, "CURRENCY")
Print #1, Tab(2); "     DESPESAS: " + Format(despesa, "CURRENCY")
Print #1, Tab(2); "     COMISSAO: " + Format(comissao, "CURRENCY")
Print #1, Tab(2); "     CONDUCAO: " + Format(conducao, "CURRENCY")
Print #1, Tab(2); "      SALARIO: " + Format(salario, "CURRENCY")
Print #1, Tab(2); "     RETIRADA: " + Format(retirada, "CURRENCY")
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "PROXIMO TROCO: " + Format(proximoTroco, "CURRENCY")
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); "************************************************"
Print #1, Tab(0); " "
'vendas por vendedor-------------------------------------------------------------
Print #1, Tab(2); "RELATORIO DE VENDAS POR VENDEDORES"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "VENDEDOR.              VENDAS.";
Print #1, Tab(0); "------------------------------------------------";

For Each v In lista1
    Print #1, Tab(2); v.Getvendedor; Tab(25); Format(v.GetvalorCompra, "currency");
    Print #1, Tab(0); "------------------------------------------------";
Next v
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); "************************************************";
Print #1, Tab(0); " "
'sub grupo quantidade------------------------------------------------------------
Print #1, Tab(2); "RELATORIO DE QUANTIDADE DE ITENS POR SUBGRUPO"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "SUB GRUPO.             QUANTIDADE.";

For Each sg In lista2
    Print #1, Tab(0); "------------------------------------------------";
    Print #1, Tab(2); sg.GetSubGrupo; Tab(25); sg.Getquantidade;
Next sg

Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Close #1
End Sub



'*********************************************************************
'                   ORCAMENTO INTERNO
'*********************************************************************
Public Function OcamentoPrint(idOrc As Integer, dataOrc As Date)
porta = "\\DESKTOP-AV8P0BQ\EPSON"

    Dim dados As New Collection
    
    Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
    preencherDadosDaEmpresa dados
    
    Dim dadosOrc As New Collection
    
    Set dadosOrc = RepositorDeOrcamento.BuscarOrcamentosPrint(idOrc, dataOrc)
    
    preencherDadosOrc dadosOrc
    
    Dim dadosProdOrc As New Collection
    
    Set dadosProdOrc = RepositorDeOrcamento.produtosParaPrint(idOrc, dataOrc)
    
    imprimirOrcamento dadosProdOrc
 
End Function

Sub preencherDadosOrc(lista As Collection)

    Dim dados As OrcamentoModelo
    For Each dados In lista
    
        vendedor = dados.Getvendedor
        idVendedor = dados.GetiDvendedor
        data = dados.Getdata
        valorOrc = dados.Getvalor
        cliente = dados.Getcliente
        idCliente = dados.GetiDcliente
        nOrc = dados.Getnumero
    
    Next dados

End Sub

Sub imprimirOrcamento(produtos As Collection)

Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(29); "N Orcamento: " + nOrc
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "Data: " + data; " Hora: " & Time;
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(15); "CONTROLE INTERNO"
Print #1, Tab(2); "Cliente: " + cliente;
Print #1, Tab(2); "Codigo Cliente: " + idCliente
Print #1, Tab(2); "Vendedor: " + vendedor; Tab(25); "Id Vendedor: " + idVendedor

Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); " Id    Descr.Produto           Quant. SubTotal ";
Print #1, Tab(0); "------------------------------------------------";

For Each p In produtos
    total = p.Getquantidade * p.Getvalor
    Print #1, Tab(2); p.GetiDproduto; Tab(9); p.Getproduto; Tab(33); p.Getquantidade; Tab(40); Format(total, "currency")
Next p

Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "VALOR TOTAL"; Tab(40); valorOrc;
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(12); textoOrc;
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Chr$(&H1D); "V"; Chr$(66); Chr$(0);
Close #1

End Sub

Sub teste()
Dim id As Integer
Dim data As Date
id = 1
data = Date
Call carnePrint(id, data)
End Sub
'*********************************************************************
'                   CARNE DE CREDIARIO
'*********************************************************************
Public Function carnePrint(idOrc As Integer, dataOrc As Date)
porta = "\\DESKTOP-AV8P0BQ\EPSON"

    Dim dados As New Collection
    
    Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
    preencherDadosDaEmpresa dados
    
    Dim dadosVenda As New Collection
    
    Set dadosVenda = RepositorDeRecebimentos.BuscarCarnePrint(idOrc, dataOrc)
    
    preencherDadosCarne dadosVenda
    
    Dim dadosCarne As New Collection
    
    Set dadosCarne = RepositorDeRecebimentos.BuscarCarnePrint(idOrc, dataOrc)
    
    imprimirCarne dadosCarne
 
End Function

Sub preencherDadosCarne(lista As Collection)
    Dim dadosParcela As New Collection
    Dim dados As RecebimentosModelo
    For Each dados In lista
    
        cliente = dados.GetDevedor
        idCliente = dados.GetidDevedor
        data = dados.GetdataCompra
        valor = dados.GetValorTotal
        
    Next dados
    
End Sub

Sub imprimirCarne(lista As Collection)
Dim c As New RecebimentosModelo

Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(2); "Rua: " + endereco + numero;
Print #1, Tab(2); "Bairro: " + bairro;
Print #1, Tab(2); "CEP: " + cep;
Print #1, Tab(2); "Cidade: " + cidade;
Print #1, Tab(2); "Tel: " + telefone;
Print #1, Tab(2); "Wathsapp: " + celular;
Print #1, Tab(2); "E-Mail: " + email
Print #1, Tab(0); " "
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " Cliente: " + cliente;
Print #1, Tab(0); " Codigo Do Cliente: " + idCliente
Print #1, Tab(0); " Data Da Compra: " + data;
Print #1, Tab(0); " Valor Da Compra: " + valor;
Print #1, Tab(0); " "
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " vencimento.  valor.         parcela.";
Print #1, Tab(0); "------------------------------------------------";

For Each c In lista
If Not c.GetvalorParcela = Empty Then
    Print #1, Tab(2); c.GetVencimento; Tab(15); Format(c.GetvalorParcela, "currency"); Tab(30); c.Getparcelas;
End If
Next c
    
Print #1, Tab(0); " ";
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); textoCarne
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Chr$(&H1D); "V"; Chr$(66); Chr$(0);
Close #1

End Sub

'*********************************************************************
'                   RELATORIO DE QUANTIDADE DE ITENS POR SUBGRUPO
'*********************************************************************
Public Function relatorioDeQuantidadeSubGrupoPrint(data As Date)
porta = "\\DESKTOP-AV8P0BQ\EPSON"

Dim dados As New Collection
    
Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
preencherDadosDaEmpresa dados
    
Dim SGQuan As New Collection
        
Set SGQuan = RepositorDeOrcamento.BuscarSubGrupoPrint(data, data)
        
imprimirRelatorioSGQuant SGQuan

End Function

Sub imprimirRelatorioSGQuant(lista As Collection)
Dim qu As New OrcamentoModelo

Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(0); "RELATORIO DE QUANTIDADE DE ITENS POR SUBGRUPO"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); Date; " Hora: " & Time;
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); "  SUB GRUPO               QUANTIDADE";

For Each qu In lista
    Print #1, Tab(0); "------------------------------------------------";
    Print #1, Tab(2); qu.GetSubGrupo; Tab(25); qu.Getquantidade;
Next qu

Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Chr$(&H1D); "V"; Chr$(66); Chr$(0);
Close #1
End Sub

Sub cortaPapel()

porta = "\\DESKTOP-AV8P0BQ\EPSON"
Open porta For Output Access Write As #1
Print #1, Chr$(&H1D); "V"; Chr$(66); Chr$(0);
Close #1

End Sub

'*********************************************************************
'                   RELATORIO DE GERAL DIARIO
'*********************************************************************
Public Function relatorioGeralPrint(data As Date, idFechamento As Integer)
porta = "\\DESKTOP-AV8P0BQ\EPSON"

'DADOS DA EMPRESA ----------------------------------------------------
Dim dados As New Collection
    
Set dados = RepositorDeDadosDaEmpresa.BuscarTodosDados
        
preencherDadosDaEmpresa dados

'DADOS DO FECHAMENTO--------------------------------------------------
Dim dadosFechamento As New Collection
    
Set dadosFechamento = RepositorDeFechamento.BuscarFechamentosPrint(idFechamento)
   
preencherDadosFechamento dadosFechamento

'DADOS DOS PAGAMENTOS -------------------------------------------------
Dim dadosRelatorioPagamento As New Collection
    
Set dadosRelatorioPagamento = RepositorDePagamentos.RelatorioPagamentosPrint(data)

'DADOS VENDAS POR VENDEDOR---------------------------------------------
Dim todasVendas As New Collection
        
Set todasVendas = RepositorDeVendasDiarias.BuscarVendaPorPeriodo(data, data)
           
'DADOS QUANTIDADE SUB GRUPO--------------------------------------------
Dim SGQuan As New Collection
        
Set SGQuan = RepositorDeOrcamento.BuscarSubGrupoPrint(data, data)
        
imprimirRelatorioGeral dadosRelatorioPagamento, todasVendas, SGQuan

End Function
Sub imprimirRelatorioGeral(lista1 As Collection, lista2 As Collection, lista3 As Collection)
Dim pagos As PagamentoModelo
Dim v As VendasDiariasModelo
Dim sg As OrcamentoModelo

Open porta For Output Access Write As #1

Print #1, Tab(16); nomeDaEmpresa
Print #1, Tab(0); " "
Print #1, Tab(0); "FECHAMENTO"
Print #1, Tab(0); " " + data; " Hora: " & Time;
Print #1, Tab(0); " "
Print #1, Tab(2); "       VENDAS: " + Format(vendas, "CURRENCY")
Print #1, Tab(2); "        TROCO: " + Format(troco, "CURRENCY")
Print #1, Tab(2); "    PAGAMENTO: " + Format(pg, "CURRENCY")
Print #1, Tab(2); "       CARTAO: " + Format(cartao, "CURRENCY")
Print #1, Tab(2); "    CREDIARIO: " + Format(crediario, "CURRENCY")
Print #1, Tab(2); "     DESPESAS: " + Format(despesa, "CURRENCY")
Print #1, Tab(2); "     COMISSAO: " + Format(comissao, "CURRENCY")
Print #1, Tab(2); "     CONDUCAO: " + Format(conducao, "CURRENCY")
Print #1, Tab(2); "      SALARIO: " + Format(salario, "CURRENCY")
Print #1, Tab(2); "     RETIRADA: " + Format(retirada, "CURRENCY")
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "PROXIMO TROCO: " + Format(proximoTroco, "CURRENCY")
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); "************************************************";
'pagamento----------------------------------------------------------------------
Print #1, Tab(0); "RELATORIO DE PAGAMENTO DE CREDIARIO"
Print #1, Tab(0); "------------------------------------------------";
Dim contenPG As Boolean

contenPG = False

For Each pagos In lista1

    Djuros = pagos.Getjuros
    Dvalor = pagos.GetvalorPg
    totalPg = Djuros + Dvalor

    contenPG = True
    Print #1, Tab(2); pagos.Getcliente;
    Print #1, Tab(2); "VENCIMENTO.  VALOR.       PARCELA:  C/JUROS:";
    Print #1, Tab(2); pagos.GetdataVencimento; Tab(15); Format(pagos.GetvalorPg, "currency"); Tab(28); pagos.Getparcela; Tab(38); Format(totalPg, "currency")
    Print #1, Tab(0); "------------------------------------------------";
Next pagos
If contenPG = False Then
    Print #1, Tab(2); "NAO HA PAGAMENTOS NO DIA "
    Print #1, Tab(0); "------------------------------------------------";
End If
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); "************************************************";
'vendas por vendedor-------------------------------------------------------------
Print #1, Tab(2); "RELATORIO DE VENDAS POR VENDEDORES"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "VENDEDOR.              VENDAS.";
Print #1, Tab(0); "------------------------------------------------";

For Each v In lista2
    Print #1, Tab(2); v.Getvendedor; Tab(25); Format(v.GetvalorCompra, "currency");
    Print #1, Tab(0); "------------------------------------------------";
Next v
Print #1, Tab(0); " "
Print #1, Tab(0); " "
Print #1, Tab(0); "************************************************";
'sub grupo quantidade------------------------------------------------------------
Print #1, Tab(2); "RELATORIO DE QUANTIDADE DE ITENS POR SUBGRUPO"
Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(2); "SUB GRUPO.             QUANTIDADE.";

For Each sg In lista3
    Print #1, Tab(0); "------------------------------------------------";
    Print #1, Tab(2); sg.GetSubGrupo; Tab(25); sg.Getquantidade;
Next sg

Print #1, Tab(0); "------------------------------------------------";
Print #1, Tab(0); " "
Close #1
End Sub








