using System.Text.Json;
using COTACAO_INSUMO;
using ClosedXML.Excel;

void Check(bool ok, string mensagem)
{
    if (!ok) throw new Exception(mensagem);
}
void Rejeitar(Action acao)
{
    try { acao(); }
    catch (Exception ex) when (ex is JsonException or RespostaIAException) { return; }
    throw new Exception("Resposta inválida foi aceita.");
}

const string json = """
    {"fornecedor":"Fornecedor Á","Insumos":[
      {"Nome":"AMORA","Preco":119.99,"Unidade":"Kg"},
      {"Nome":"AMORA","Preco":0.0001,"Unidade":"G"},
      {"Nome":"ITEM SEM BASE","Preco":null,"Unidade":null},
      {"Nome":"FRASCO","Preco":24.7,"Unidade":"Milheiro"}]}
    """;
var resultado = LeituraInsumos.Interpretar(json);
Check(resultado.Insumos.Count == 4 && resultado.Insumos[1].Preco == .0001m,
    "Preservar todas as linhas, repetições, preços pequenos e itens indeterminados.");
foreach (var unidade in new[] { "Kg", "G", "ML", "L", "Unidade", "Milheiro" })
    Check(LeituraInsumos.Interpretar(json.Replace("Milheiro", unidade)).Insumos[3].Unidade == unidade,
        "Aceitar unidade " + unidade);
Rejeitar(() => LeituraInsumos.Interpretar(json.Replace("119.99", "\"119,99\"")));
Rejeitar(() => LeituraInsumos.Interpretar(json.Replace("Milheiro", "EMBALAGEM")));
Rejeitar(() => LeituraInsumos.Interpretar(json.Replace("\"Unidade\":\"Kg\"", "\"OutraUnidade\":\"Kg\"")));
Rejeitar(() => LeituraInsumos.Interpretar(json.Replace("119.99", "-1")));
Rejeitar(() => LeituraInsumos.Interpretar("{\"fornecedor\":\"X\",\"Insumos\":[]}"));

foreach (var (cabecalho, quantidade, informada, esperada) in new[] {
    ("Preço kg/l", "50 g", "G", "Kg"),
    ("Preço kg/l", "250 g", "G", "Kg"),
    ("Preço kg/l", "800 ml", "ML", "Kg"),
    ("Preço/L", "800 ml", "ML", "Kg"),
    ("R$/KG", "50 g", "G", "Kg"),
    ("R$/G", "1 kg", "Kg", "G"),
    ("Preço kg/l", "1.000 Cap", "Unidade", "Kg") })
{
    string resposta = JsonSerializer.Serialize(new { fornecedor = "Nutrifarm", Insumos = new[] {
        new { Nome = "ACIDO TRANEXAMICO", Preco = 3299.99m, Unidade = informada,
            CabecalhoPreco = cabecalho, QuantidadeCotada = quantidade, BasePrecoLinha = (string?)null } } });
    var corrigido = LeituraInsumos.InterpretarRespostaPdf(resposta);
    Check(corrigido.Insumos[0].Unidade == esperada && corrigido.Insumos[0].Preco == 3299.99m,
        "Conferir unidade pelo cabeçalho e dimensão da quantidade, sem alterar preço.");
    if (esperada == "Kg")
        Check(ConversorPrecoUnitario.Converter(corrigido.Insumos[0].Preco, esperada) == 3.3m,
            "Ácido tranexâmico: 3299,99/kg deve produzir 3,3000/g.");
    Check(!JsonSerializer.Serialize(corrigido).Contains("CabecalhoPreco"), "Não exportar evidências internas.");
}
Rejeitar(() => LeituraInsumos.InterpretarRespostaPdf(json));

var fornecedorPadronizado = LeituraInsumos.InterpretarRespostaPdf("""
    {"fornecedor":"Fornecedor","Insumos":[
      {"Nome":"A","Preco":100,"Unidade":"L","CabecalhoPreco":"Preço/L","QuantidadeCotada":"1 L","BasePrecoLinha":null},
      {"Nome":"B","Preco":200,"Unidade":"Unidade","CabecalhoPreco":null,"QuantidadeCotada":"100 Cap","BasePrecoLinha":"Unidade"},
      {"Nome":"C","Preco":300,"Unidade":"G","CabecalhoPreco":"Preço/G","QuantidadeCotada":"50 g","BasePrecoLinha":null}]}
    """);
Check(fornecedorPadronizado.Insumos.All(item => item.Unidade == "Kg"),
    "Cabeçalho KG ou L deve padronizar todos os insumos do fornecedor como Kg.");
Check(fornecedorPadronizado.Insumos.Select(item => item.Preco).SequenceEqual(new decimal?[] {100m, 200m, 300m}),
    "Padronização do fornecedor deve preservar os preços originais.");

foreach (var (provedor, resposta) in new[] {
    ("OpenAI", """{"status":"incomplete","output":[]}"""),
    ("Gemini", """{"candidates":[{"finishReason":"MAX_TOKENS"}]}"""),
    ("Claude", """{"stop_reason":"max_tokens","content":[]}""") })
{
    using var doc = JsonDocument.Parse(resposta);
    Rejeitar(() => ExtracaoPdfIA.ExtrairTexto(doc.RootElement, provedor));
}

string pasta = Path.Combine(AppContext.BaseDirectory, "teste-" + Guid.NewGuid().ToString("N"));
try
{
    var configuracoesJson = new ConfiguracaoIARepository(Path.Combine(pasta, "configuracoes.db"));
    Check(!configuracoesJson.CarregarGravacaoJson(), "Gravação JSON deve iniciar desligada.");
    configuracoesJson.SalvarGravacaoJson(true);
    Check(new ConfiguracaoIARepository(Path.Combine(pasta, "configuracoes.db")).CarregarGravacaoJson(),
        "Opção JSON deve persistir entre sessões.");
    configuracoesJson.SalvarGravacaoJson(false);
    Check(!configuracoesJson.CarregarGravacaoJson(), "Deve permitir desligar a gravação JSON.");
    string pastaRetorno = Path.Combine(pasta, "retorno");
    string retorno = await ExportadorInsumos.SalvarCamposAsync(
        ComparadorInsumos.Comparar(resultado, new[] { "AMORA" }), "orcamento.pdf", pastaRetorno);
    Check(Directory.GetFiles(pastaRetorno).Length == 1 && retorno.EndsWith("_campos.json"),
        "Exportação após comparação deve criar apenas o JSON _campos.");
    File.Delete(retorno);
    Directory.Delete(pastaRetorno);
    string primeiro = await ExportadorInsumos.SalvarAsync(resultado, "orcamento.pdf", pasta);
    string segundo = await ExportadorInsumos.SalvarAsync(resultado, "orcamento.pdf", pasta);
    Check(primeiro != segundo && Directory.GetFiles(pasta, "*.json").Length == 2, "Não sobrescrever extrações anteriores.");
    using var doc = JsonDocument.Parse(await File.ReadAllTextAsync(primeiro));
    Check(doc.RootElement.EnumerateObject().Select(x => x.Name).SequenceEqual(new[] { "fornecedor", "Insumos" }),
        "JSON deve ter apenas os campos solicitados.");
    Check(doc.RootElement.GetProperty("Insumos")[0].EnumerateObject().Select(x => x.Name)
        .SequenceEqual(new[] { "Nome", "Preco", "Unidade" }), "Formato exato do insumo.");
    Check(LeituraInsumos.Interpretar(await File.ReadAllTextAsync(primeiro)).Insumos.Count == 4,
        "Persistência deve manter todos os itens.");

    Check(ComparadorInsumos.MatchInsumo("ácido   cítrico", [" ÁCIDO CÍTRICO "]), "Normalizar nomes.");
    Check(ComparadorInsumos.MatchInsumo("AMORA", ["AMORA EXTRA"], 100), "Referência pela menor string.");
    Check(ComparadorInsumos.MatchInsumo("AMORA EXTRA", ["AMORA"], 100), "Comparação simétrica por tamanho.");
    Check(ComparadorInsumos.MatchInsumo("ABCDEFGHIJKLMNOPQRST", ["ABCDEFGHIJKLMNOPQXYZ"], 85), "Limite exato de 85%.");
    Check(!ComparadorInsumos.MatchInsumo("ABCDEFGHIJKLMNOPQRST", ["ABCDEFGHIJKLMNOPQXYZ"], 85.01m), "Abaixo do limite.");
    Check(ComparadorInsumos.MatchInsumo("ABCDEFGHIJKLMNOPQRST", ["ABCDEFGHIJKLMNOPQXYZ"], 84), "Acima do limite.");
    Check(!ComparadorInsumos.MatchInsumo("ABCD", ["AXBXCXD"], 85), "Trecho deve ser contínuo.");
    Check(!ComparadorInsumos.MatchInsumo(" ", ["AMORA"]) && !ComparadorInsumos.MatchInsumo("AMORA", [""]), "Nomes vazios.");
    Check(ComparadorInsumos.NormalizarNome(" Amóra  (  uRgEnTe )  ", true) == "AMORA", "Remover tag final.");
    Check(ComparadorInsumos.NormalizarNome("AMORA (EXTRATO)", true) == "AMORA (EXTRATO)", "Preservar outros parênteses.");
    Check(ComparadorInsumos.NormalizarNome("AMORA (Urgente) EXTRA", true) == "AMORA (URGENTE) EXTRA", "Não remover tag intermediária.");
    Check(ComparadorInsumos.NormalizarNome("AMORA (Urgente)") == "AMORA (URGENTE)", "Não remover tag do PDF.");
    Check(ComparadorInsumos.NormalizarNome("B12 (1%)") == "B12 (1%)", "Preservar números e pontuação.");
    foreach (decimal percentual in new[] { 0m, -1m, 100.1m })
    {
        bool rejeitado = false;
        try { ComparadorInsumos.MatchInsumo("A", ["A"], percentual); }
        catch (ArgumentOutOfRangeException) { rejeitado = true; }
        Check(rejeitado, "Validar configuração do percentual.");
    }

    string excel = Path.Combine(pasta, "cotacao.xlsx");
    using (var workbook = new XLWorkbook())
    {
        var aba = workbook.AddWorksheet("Cotação");
        aba.Cell(1, 1).Value = "Título personalizado";
        aba.Cell(1, 2).Value = "Insumos";
        aba.Cell(2, 1).Value = "AMORA (Urgente)";
        aba.Cell(4, 1).Value = "FRASCO";
        aba.Cell(2, 2).Value = "OUTRO PRODUTO";
        workbook.CustomProperties.Add("CotacaoGridRoles", "[\"INSUMO\",\"FORNECEDOR\"]");
        workbook.AddWorksheet("Outra aba").Cell(1, 1).Value = "IGNORAR";
        workbook.SaveAs(excel);
    }
    byte[] excelAntes = await File.ReadAllBytesAsync(excel);
    var nomes = LeitorPlanilhaInsumos.Ler(excel);
    Check(nomes.SequenceEqual(new[] { "AMORA (Urgente)", "FRASCO" }), "Metadados têm prioridade; ler todas as linhas da primeira aba.");
    var campos = ComparadorInsumos.Comparar(resultado, nomes);
    Check(campos.Insumos.Select(x => x.Encontrado).SequenceEqual(new[] { true, true, false, true }), "Comparar todos os insumos.");
    for (int i = 0; i < resultado.Insumos.Count; i++)
        Check(campos.Insumos[i].Nome == resultado.Insumos[i].Nome && campos.Insumos[i].Preco == resultado.Insumos[i].Preco &&
            campos.Insumos[i].Unidade == resultado.Insumos[i].Unidade, "Preservar valores originais e ordem.");
    string camposPath = await ExportadorInsumos.SalvarCamposAsync(campos, primeiro);
    Check(camposPath.EndsWith("_campos.json") && Path.GetDirectoryName(camposPath) == pasta, "Local e sufixo do JSON adicional.");
    using (var camposDoc = JsonDocument.Parse(await File.ReadAllTextAsync(camposPath)))
    {
        Check(camposDoc.RootElement.GetProperty("fornecedor").GetString() == resultado.Fornecedor, "Preservar fornecedor.");
        Check(camposDoc.RootElement.GetProperty("Insumos")[0].EnumerateObject().Select(x => x.Name)
            .SequenceEqual(new[] { "Nome", "Preco", "Unidade", "Encontrado", "PrecoUnitario" }), "Formato exato do JSON adicional.");
        Check(camposDoc.RootElement.GetProperty("Insumos")[0].GetProperty("PrecoUnitario").GetDecimal() == .12m,
            "Persistir preço convertido no JSON _campos.");
    }
    byte[] excelDepois = await File.ReadAllBytesAsync(excel);
    Check(excelAntes.SequenceEqual(excelDepois), "Planilha deve permanecer intacta.");
    Check(LeituraInsumos.Interpretar(await File.ReadAllTextAsync(primeiro)).Insumos.Count == 4, "JSON original intacto após comparação.");
    using (var workbook = new XLWorkbook(excel))
    {
        workbook.CustomProperties.Delete("CotacaoGridRoles");
        workbook.Worksheet(1).Cell(1, 2).Value = "SEM CABEÇALHO";
        workbook.Worksheet(1).Cell(1, 1).Value = " INSUMOS ";
        workbook.Save();
    }
    Check(LeitorPlanilhaInsumos.Ler(excel).Count == 2, "Fallback pelo cabeçalho.");
    using (var workbook = new XLWorkbook(excel))
    {
        workbook.Worksheet(1).Cell(1, 1).Value = "DESCONHECIDO";
        workbook.Save();
    }
    bool erroCabecalho = false;
    try { LeitorPlanilhaInsumos.Ler(excel); }
    catch (InvalidDataException) { erroCabecalho = true; }
    Check(erroCabecalho && File.Exists(primeiro), "Falha de leitura deve preservar o original sem falsos negativos.");

    foreach (var (unidade, preco, esperado) in new[] {
        ("Kg", 119.99m, .12m), ("L", 52m, .052m), ("Milheiro", 24.7m, .0247m),
        ("G", 1.23456m, 1.2346m), ("ML", 2.12344m, 2.1234m), ("Unidade", 3.12345m, 3.1235m),
        ("Kg", .05m, .0001m), ("L", .01m, 0m), ("G", 0m, 0m) })
        Check(ConversorPrecoUnitario.Converter(preco, unidade) == esperado, "Conversão/arredondamento de " + unidade);
    Check(ConversorPrecoUnitario.Converter(null, "Kg") == null &&
        ConversorPrecoUnitario.Converter(10m, null) == null &&
        ConversorPrecoUnitario.Converter(10m, "DESCONHECIDA") == null, "Não inventar conversão sem preço/base.");
    Check(campos.Insumos[2].PrecoUnitario == null && campos.Insumos[3].PrecoUnitario == .0247m,
        "Conversão de todos os itens após leitura, preservando null.");
    var semMatch = ComparadorInsumos.Comparar(resultado, ["ZZZZZZZZ"]);
    Check(!semMatch.Insumos[0].Encontrado && semMatch.Insumos[0].PrecoUnitario == .12m,
        "Converter também insumos não encontrados na planilha.");

    // Resposta esperada para os exemplos com cabeçalho "Preço kg/l".
    // Valida interpretação/conversão; a classificação visual pela IA exige PDF real.
    var tabelaKgL = LeituraInsumos.Interpretar("""
        {"fornecedor":"NUTRIFARM","Insumos":[
          {"Nome":"COGNIZIN (CITICOLINA)","Preco":9900.00,"Unidade":"Kg"},
          {"Nome":"AMORA PRETA EXT SOLUVEL","Preco":119.99,"Unidade":"Kg"},
          {"Nome":"VASELINA LIQUIDA","Preco":52.00,"Unidade":"L"}]}
        """);
    var convertidosKgL = ComparadorInsumos.Comparar(tabelaKgL, ["COGNIZIN", "AMORA", "VASELINA"]);
    Check(convertidosKgL.Insumos.Select(x => x.PrecoUnitario).SequenceEqual(new decimal?[] { 9.9m, .12m, .052m }),
        "Preços kg/l devem converter para grama/mililitro, preservando o preço original.");
    Check(convertidosKgL.Insumos[0].Preco == 9900m && convertidosKgL.Insumos[0].Unidade == "Kg",
        "COGNIZIN: base original Kg e preço unitário por grama.");

    string deslocada = Path.Combine(pasta, "cotacao-deslocada.xlsx");
    using (var workbook = new XLWorkbook())
    {
        var aba = workbook.AddWorksheet("INSUMOS");
        aba.Cell(2, 2).Value = "Cotação parcialmente preenchida";
        aba.Cell(5, 3).Value = "INSUMOS DROGA CAMPOS - DRUG 09/2026";
        aba.Cell(5, 4).Value = "Fornecedor mais em conta";
        aba.Cell(5, 5).Value = "QUANTIDADE A COMPRAR (MG OU ML)";
        aba.Cell(5, 6).Value = "GALENA";
        aba.Cell(6, 3).Value = "ACETIL L CARNITINA HCL";
        aba.Cell(6, 4).Value = "VALDEQUIMI";
        aba.Cell(6, 5).Value = 3000;
        aba.Cell(8, 3).Value = "AMORA (Urgente)";
        aba.Cell(8, 6).Value = .0405;
        aba.Cell(9, 3).Value = "ÚLTIMO INSUMO";
        workbook.SaveAs(deslocada);
    }
    byte[] antesDeslocada = await File.ReadAllBytesAsync(deslocada);
    Check(LeitorPlanilhaInsumos.Ler(deslocada).SequenceEqual(new[] {
        "ACETIL L CARNITINA HCL", "AMORA (Urgente)", "ÚLTIMO INSUMO"
    }), "Título completo deslocado: ler todas as linhas, inclusive após lacunas, sem fornecedores ou preços.");
    byte[] depoisDeslocada = await File.ReadAllBytesAsync(deslocada);
    Check(antesDeslocada.SequenceEqual(depoisDeslocada), "Não alterar cotação parcialmente preenchida.");
    string jsonAntesPersistencia = await File.ReadAllTextAsync(primeiro);
    string camposAntesPersistencia = await File.ReadAllTextAsync(camposPath);
    TestesPersistencia.Executar(pasta);
    TestesPreenchimento.Executar(pasta);
    TestesXmlPlanilha.Executar(pasta, args.FirstOrDefault());
    Check(await File.ReadAllTextAsync(primeiro) == jsonAntesPersistencia &&
        await File.ReadAllTextAsync(camposPath) == camposAntesPersistencia,
        "Persistência, inclusive em falha, não deve alterar JSONs já gravados.");
}
finally
{
    if (Directory.Exists(pasta))
    {
        foreach (string arquivo in Directory.GetFiles(pasta)) File.Delete(arquivo);
        Directory.Delete(pasta);
    }
}
Console.WriteLine("PASS: formato JSON, unidades, precisão, itens indeterminados, rejeição de respostas interrompidas e gravação sem sobrescrita.");
Console.WriteLine("PASS: comparação, percentual, tag Urgente, leitura da planilha, JSON _campos e preservação de dados.");
