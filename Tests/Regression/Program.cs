using COTACAO_INSUMO;
using ClosedXML.Excel;
void Check(bool value, string message) { if (!value) throw new Exception(message); }
ItemCotacaoIA Amora() => new() { ProdutoPdf = "AMORA PRETA EXT SOLUVEL", UnidadeOriginal = "KG", Quantidade = 500, ValorUnitario = 119.99m, ValorTotal = 60, Encontrado = false, Confianca = 0 };
var item = Amora();
Check(NormalizadorCotacao.Calcular(item).PodeConverter && item.PrecoNormalizado == 0.11999m && item.TipoPreco == "GRAMA", "Amora sem match");
foreach (var (unit, price, expected, type) in new[] { ("500 G", 60m, 0.12m, "GRAMA"), ("L", 52m, 0.052m, "MILILITRO"), ("500 MG", 10m, 20m, "GRAMA"), ("MLH", 24.7m, 0.0247m, "UNIDADE") }) {
    var sample = new ItemCotacaoIA { UnidadeOriginal = unit, ValorUnitario = price };
    Check(NormalizadorCotacao.Calcular(sample).PodeConverter && sample.PrecoNormalizado == expected && sample.TipoPreco == type, unit);
}
var path = Path.Combine(AppContext.BaseDirectory, "regression.xlsx");
using (var book = new XLWorkbook()) { var sheet = book.AddWorksheet("Cotacao"); sheet.Cell(1,1).Value = "INSUMO"; sheet.Cell(2,1).Value = "OUTRO PRODUTO"; book.SaveAs(path); }
var valid = new Fake(_ => new() { Fornecedor = "NUTRIFARM", Itens = [Amora()] });
var result = await new CotacaoProcessor(valid).ProcessarAsync(path, ["cotacao.pdf"]);
Check(valid.Calls == 1 && result.NaoEncontrados.Count == 1, "Sem match deve ser preservado");
Check(result.NaoEncontrados[0].PrecoNormalizado == 0.11999m && result.NaoEncontrados[0].TipoPreco == "GRAMA", "Preco e tipo na lista final");
var retry = new Fake(n => new() { Fornecedor = "NUTRIFARM", Itens = [n == 1 ? new() { ProdutoPdf = "AMORA PRETA EXT SOLUVEL" } : Amora()] });
var retried = await new CotacaoProcessor(retry).ProcessarAsync(path, ["cotacao.pdf"]);
Check(retry.Calls == 2 && retried.NaoEncontrados[0].TipoPreco == "GRAMA", "Recuperacao pela segunda leitura");
var invalid = new Fake(_ => new() { Itens = [new() { ProdutoPdf = "AMORA" }] });
var incomplete = await new CotacaoProcessor(invalid).ProcessarAsync(path, ["cotacao.pdf", "outra.pdf"]);
Check(invalid.Calls == 6 && incomplete.NaoEncontrados.Count == 2, "Pendencias nao bloqueiam os demais PDFs");
var failingRetry = new Fake(n => n == 1 ? new() { Fornecedor = "NUTRIFARM", Itens = [Amora(), new() { ProdutoPdf = "PENDENTE" }] } : throw new HttpRequestException("Falha na releitura"));
var recovered = await new CotacaoProcessor(failingRetry).ProcessarAsync(path, ["cotacao.pdf"]);
Check(recovered.NaoEncontrados.Count == 2 && recovered.NaoEncontrados[0].PrecoNormalizado == .11999m, "Releitura com falha preserva os itens");
var lostItem = new Fake(n => new() { Fornecedor = "NUTRIFARM", Itens = n == 1 ? [Amora(), new() { ProdutoPdf = "PENDENTE" }] : [new() { ProdutoPdf = "PENDENTE" }] });
var preserved = await new CotacaoProcessor(lostItem).ProcessarAsync(path, ["cotacao.pdf"]);
Check(preserved.NaoEncontrados.Count == 2 && preserved.NaoEncontrados[0].TipoPreco == "GRAMA", "Releitura incompleta nao descarta itens validos");
var wrongBasis = new ItemCotacaoIA { ProdutoPdf = "AMORA", UnidadeOriginal = "500 G", UnidadeQuantidade = "G", Quantidade = 500, ValorUnitario = 119.99m, ValorTotal = 60 };
Check(!NormalizadorCotacao.Calcular(wrongBasis).PodeConverter && wrongBasis.ValorUnitario == 119.99m, "Divergencia exige releitura sem trocar preco");
var wrongVolume = new ItemCotacaoIA { UnidadeOriginal = "800 ML", UnidadeQuantidade = "ML", Quantidade = 800, ValorUnitario = 52, ValorTotal = 41.60m };
Check(!NormalizadorCotacao.Calcular(wrongVolume).PodeConverter && wrongVolume.ValorUnitario == 52m, "Nao reconstruir preco pelo total de volume");
var exactKg = Amora(); exactKg.UnidadeQuantidade = "G";
Check(NormalizadorCotacao.Calcular(exactKg).PrecoNormalizado == .11999m, "Preservar preco KG com arredondamento do total");
var missingBasis = new ItemCotacaoIA { UnidadeQuantidade = "G", Quantidade = 25, ValorTotal = 247.5m };
Check(!NormalizadorCotacao.Calcular(missingBasis).PodeConverter && missingBasis.ValorUnitario == 0, "Nao inferir preco a partir de quantidade e total");
var repo = new NaoEncontradosRepository(Path.Combine(AppContext.BaseDirectory, "legacy-" + Guid.NewGuid() + ".db"));
repo.SalvarLote("TESTE", 10, 2026, [new() { Fornecedor = "NUTRIFARM", Insumo = "AMORA", UnidadeOriginal = "g", TipoPreco = "NAO_IDENTIFICADO", PrecoNormalizado = 0 }], DateTime.Now);
var legacy = repo.Consultar("TESTE", 10, 2026).Single();
Check(legacy.TipoPreco == "GRAMA" && legacy.PrecoNormalizado == 0, "Identificar unidade antiga sem inventar preco");
Console.WriteLine("PASS: validacao das bases KG e L, arredondamento e unidade dos registros antigos.");
Check(!CalculadoraPreco.Calcular("KG", 2, 200, 0).PodeConverter, "Calculadora nao usa total como preco alternativo");
var partial = NormalizadorCotacao.Complementar(
    new() { Itens = [new() { ProdutoPdf = "AMORA", UnidadeOriginal = "KG", ValorUnitario = 0 }] },
    new() { Itens = [new() { ProdutoPdf = "AMORA", ValorUnitario = 119.99m }] });
Check(partial!.Itens.Single().PrecoNormalizado == .11999m, "Combinar preco relido com unidade ja extraida");
Console.WriteLine("PASS: preco somente do campo comercial, divergencias pedem releitura, complemento parcial e continuidade.");
var readingPrompt = PromptCotacao.Criar(PromptCotacao.PrepararReleitura([Amora()]));
Check(!readingPrompt.Contains("LISTA OFICIAL") && readingPrompt.Contains("AMORA PRETA EXT SOLUVEL"), "Releitura separada da comparacao com Excel");
var fragmented = new Fake(n => new() { Fornecedor = "NUTRIFARM", Itens = [n switch {
    1 => new() { ProdutoPdf = "AMORA PRETA EXT SOLUVEL" },
    2 => new() { ProdutoPdf = "AMORA PRETA EXT SOLUVEL", ValorUnitario = 119.99m },
    _ => new() { ProdutoPdf = "AMORA PRETA EXT SOLUVEL", UnidadeOriginal = "KG" }
}] });
var assembled = await new CotacaoProcessor(fragmented).ProcessarAsync(path, ["cotacao.pdf"]);
Check(fragmented.Calls == 3 && assembled.NaoEncontrados.Single().PrecoNormalizado == .11999m, "Preservar preco e recuperar unidade em etapas");
var tiny = new Fake(_ => new() { Fornecedor = "TESTE", Itens = [new() { ProdutoPdf = "SEM MATCH", UnidadeOriginal = "KG", ValorUnitario = .10m }] });
var tinyResult = await new CotacaoProcessor(tiny).ProcessarAsync(path, ["cotacao.pdf"]);
Check(tinyResult.NaoEncontrados.Single().PrecoNormalizado == .0001m, "Preco pequeno nao pode virar zero");
var many = new Fake(n => new() { Fornecedor = "TESTE", Itens = n == 1
    ? Enumerable.Range(0, 13).Select(i => new ItemCotacaoIA { ProdutoPdf = "ITEM " + i }).ToList()
    : System.Text.Json.JsonSerializer.Deserialize<ResultadoCotacaoIA>("{\"Itens\":[]}")!.Itens });
await new CotacaoProcessor(many).ProcessarAsync(path, ["cotacao.pdf"]);
Check(many.Calls == 7, "Releitura em lotes de seis, duas rodadas sem bloquear");
Console.WriteLine("PASS: releitura exclusiva, complemento em etapas, pequenos precos e lotes limitados.");
File.Delete(path);
Console.WriteLine("PASS: amora sem match, massa/volume/unidades, lista final, releitura, continuidade com pendencias e preservacao apos falha.");
class Fake(Func<int, ResultadoCotacaoIA> response) : IAgenteIA {
    public int Calls;
    public Task<bool> TestarConexaoAsync() => Task.FromResult(true);
    public Task<ResultadoCotacaoIA?> AnalisarPdfAsync(string path, string prompt) => Task.FromResult<ResultadoCotacaoIA?>(response(++Calls));
}
