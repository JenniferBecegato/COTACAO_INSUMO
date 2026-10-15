using ClosedXML.Excel;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace COTACAO_INSUMO;

public sealed record ResultadoPreenchimento(int PrecosPreenchidos, int FornecedoresCriados,
    IReadOnlyList<string> Pendencias);

public static class PreenchedorCotacao
{
    public const decimal PercentualMinimoFornecedor = 90m;

    public static string NormalizarFornecedor(string nome)
    {
        string texto = ComparadorInsumos.NormalizarNome(nome);
        texto = Regex.Replace(texto, @"\bS\s*\.\s*A\s*\.?\b", " SA ");
        texto = Regex.Replace(texto, @"[^\p{L}\p{N}\s]", " ");
        texto = Regex.Replace(texto, @"\b(?:LTDA|LIMITADA|SA)\b", " ");
        return Regex.Replace(texto, @"\s+", " ").Trim();
    }

    private static (int Indice, bool Ambiguo) MelhorMatch(string nome,
        IEnumerable<(int Indice, string Nome)> candidatos, Func<string, string> normalizar,
        decimal minimo, bool fornecedor = false)
    {
        var avaliados = candidatos.Select(c =>
        {
            string normalizado = normalizar(c.Nome);
            bool igual = nome == normalizado && nome.Length > 0;
            decimal nota = fornecedor && Math.Min(nome.Length, normalizado.Length) < 4 && !igual
                ? 0 : ComparadorInsumos.CalcularSimilaridade(nome, normalizado);
            return (c.Indice, Igual: igual, Nota: nota);
        }).Where(x => x.Nota >= minimo).ToArray();
        if (avaliados.Any(x => x.Igual)) avaliados = avaliados.Where(x => x.Igual).ToArray();
        if (avaliados.Length == 0) return (0, false);
        decimal melhor = avaliados.Max(x => x.Nota);
        var finalistas = avaliados.Where(x => x.Nota == melhor).ToArray();
        return finalistas.Length == 1 ? (finalistas[0].Indice, false) : (0, true);
    }

    public static bool EhFornecedor(string titulo)
    {
        string nome = ComparadorInsumos.NormalizarNome(titulo);
        return nome.Length > 0 && !nome.StartsWith("INSUMO", StringComparison.Ordinal) &&
            !nome.Contains("FORNECEDOR ANTERIOR") && !nome.Contains("FORNECEDOR MAIS EM CONTA") &&
            !nome.Contains("QUANTIDADE") && nome != "QTD";
    }

    public static ResultadoPreenchimento Preencher(string caminho, InsumosCampos dados)
    {
        string destino = Path.GetFullPath(caminho);
        string temporario = Path.Combine(Path.GetDirectoryName(destino)!,
            $".{Path.GetFileNameWithoutExtension(destino)}.{Guid.NewGuid():N}.tmp.xlsx");
        var pendencias = new List<string>();
        int preenchidos = 0, criados = 0;
        // Confere o conteúdo antes de substituir alterações de outro processo.
        byte[] original = File.ReadAllBytes(destino);
        try
        {
            using (var workbook = ArquivoPlanilhaExcel.Abrir(original))
            {
                var estrutura = LeitorPlanilhaInsumos.LerEstrutura(workbook);
                var aba = workbook.Worksheets.First();
                var gravacoes = new List<(int Linha, decimal Preco)>();
                foreach (var item in dados.Insumos.Where(x => x.Encontrado))
                {
                    if (item.PrecoUnitario is not > 0)
                    {
                        pendencias.Add($"{item.Nome}: preço unitário ausente ou zero; célula preservada.");
                        continue;
                    }
                    var match = MelhorMatch(ComparadorInsumos.NormalizarNome(item.Nome),
                        estrutura.Insumos.Select(x => (x.Linha, x.Nome)),
                        nome => ComparadorInsumos.NormalizarNome(nome, true), ComparadorInsumos.PercentualMinimoMatch);
                    if (match.Indice == 0)
                    {
                        pendencias.Add($"{item.Nome}: " + (match.Ambiguo ? "match de linha ambíguo." : "linha não localizada."));
                        continue;
                    }
                    gravacoes.Add((match.Indice, Math.Round(item.PrecoUnitario.Value, 4, MidpointRounding.AwayFromZero)));
                }
                if (gravacoes.Count == 0) return new(0, 0, pendencias);

                string Papel(int coluna) => estrutura.PapeisColunas?[coluna - 1]
                    ?? aba.Cell(estrutura.LinhaCabecalho, coluna).GetString();
                var fornecedores = Enumerable.Range(1, estrutura.UltimaColuna)
                    .Where(c => c != estrutura.ColunaInsumos && EhFornecedor(Papel(c))).ToList();
                string nomeFornecedor = NormalizarFornecedor(dados.Fornecedor);
                if (nomeFornecedor.Length == 0)
                    return new(0, 0, pendencias.Append("Fornecedor sem nome válido.").ToArray());
                var matchFornecedor = MelhorMatch(nomeFornecedor,
                    fornecedores.Select(c => (c, aba.Cell(estrutura.LinhaCabecalho, c).GetString())),
                    NormalizarFornecedor, PercentualMinimoFornecedor, true);
                if (matchFornecedor.Ambiguo)
                    return new(0, 0, pendencias.Append($"{dados.Fornecedor}: match de fornecedor ambíguo; nenhuma coluna criada.").ToArray());
                var colunasMenor = Enumerable.Range(1, estrutura.UltimaColuna)
                    .Where(c => ComparadorInsumos.NormalizarNome(Papel(c)).Contains("FORNECEDOR MAIS EM CONTA")).ToArray();
                if (colunasMenor.Length > 1)
                    return new(0, 0, pendencias.Append("Coluna Fornecedor mais em conta ambígua.").ToArray());

                int ultimaColuna = estrutura.UltimaColuna;
                var papeis = Enumerable.Range(1, ultimaColuna).Select(Papel).ToList();
                int AdicionarColuna(string titulo, string papel, int modelo)
                {
                    int coluna = ++ultimaColuna;
                    int ultimaLinha = estrutura.Insumos.Max(x => x.Linha);
                    for (int linha = estrutura.LinhaCabecalho; linha <= ultimaLinha; linha++)
                        aba.Cell(linha, coluna).Style = aba.Cell(linha, modelo).Style;
                    aba.Column(coluna).Width = aba.Column(modelo).Width;
                    aba.Cell(estrutura.LinhaCabecalho, coluna).Value = titulo;
                    papeis.Add(papel);
                    return coluna;
                }
                int fornecedor = matchFornecedor.Indice;
                if (fornecedor == 0)
                {
                    fornecedor = AdicionarColuna(dados.Fornecedor, dados.Fornecedor,
                        fornecedores.LastOrDefault(estrutura.UltimaColuna));
                    fornecedores.Add(fornecedor);
                    criados++;
                }
                int colunaMenor = colunasMenor.Length == 1 ? colunasMenor[0]
                    : AdicionarColuna("Fornecedor mais em conta", "FORNECEDOR MAIS EM CONTA", estrutura.ColunaInsumos);
                foreach (int coluna in fornecedores)
                    foreach (var linha in estrutura.Insumos)
                        aba.Cell(linha.Linha, coluna).Style.NumberFormat.Format = "0.0000";

                foreach (var gravacao in gravacoes)
                {
                    aba.Cell(gravacao.Linha, fornecedor).Value = gravacao.Preco;
                    preenchidos++;
                    decimal? menor = null;
                    int vencedor = 0;
                    foreach (int coluna in fornecedores.Order())
                    {
                        var celula = aba.Cell(gravacao.Linha, coluna);
                        if (!TentarPreco(celula, out decimal preco) || preco <= 0) continue;
                        if (menor == null || preco < menor)
                        {
                            menor = preco;
                            vencedor = coluna;
                        }
                    }
                    aba.Cell(gravacao.Linha, colunaMenor).Value = vencedor == 0 ? ""
                        : aba.Cell(estrutura.LinhaCabecalho, vencedor).GetString();
                }
                workbook.CustomProperties.Delete("CotacaoGridRoles");
                workbook.CustomProperties.Add("CotacaoGridRoles", JsonSerializer.Serialize(papeis));
                workbook.SaveAs(temporario);
            }
            ArquivoPlanilhaExcel.ValidarArquivoTemporario(temporario);
            if (!original.AsSpan().SequenceEqual(File.ReadAllBytes(destino)))
                throw new IOException("A planilha foi alterada durante o processamento. As alterações externas foram preservadas; processe novamente.");
            File.Replace(temporario, destino, null);
            return new(preenchidos, criados, pendencias);
        }
        finally
        {
            if (File.Exists(temporario)) File.Delete(temporario);
        }
    }

    private static bool TentarPreco(IXLCell celula, out decimal preco)
    {
        if (celula.DataType == XLDataType.Number) return celula.TryGetValue(out preco);
        string texto = celula.GetString().Trim();
        // Sem vírgula, o ponto representa decimal; evita ler "0.0405" como 405.
        var cultura = texto.Contains(',') ? CultureInfo.GetCultureInfo("pt-BR") : CultureInfo.InvariantCulture;
        return decimal.TryParse(texto, NumberStyles.Number, cultura, out preco);
    }
}
