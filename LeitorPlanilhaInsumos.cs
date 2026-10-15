using ClosedXML.Excel;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace COTACAO_INSUMO;

public sealed record LinhaInsumoPlanilha(int Linha, string Nome);
public sealed record EstruturaPlanilhaInsumos(int LinhaCabecalho, int ColunaInsumos,
    int UltimaColuna, IReadOnlyList<LinhaInsumoPlanilha> Insumos, string[]? PapeisColunas);

public static class LeitorPlanilhaInsumos
{
    public static List<string> Ler(string caminho)
    {
        using var workbook = ArquivoPlanilhaExcel.Abrir(caminho);
        return LerEstrutura(workbook).Insumos.Select(x => x.Nome).ToList();
    }

    public static EstruturaPlanilhaInsumos LerEstrutura(XLWorkbook workbook)
    {
        var aba = workbook.Worksheets.First();
        var intervalo = aba.RangeUsed() ?? throw new InvalidDataException("A planilha está vazia.");
        int coluna = 0, linhaCabecalho = 1;
        string[]? papeisValidos = null;
        var propriedade = workbook.CustomProperties.FirstOrDefault(p => p.Name == "CotacaoGridRoles");
        if (propriedade?.Value is string json)
        {
            try
            {
                var papeis = JsonSerializer.Deserialize<string[]>(json);
                if (papeis != null && papeis.Length == intervalo.LastColumn().ColumnNumber())
                {
                    papeisValidos = papeis;
                    coluna = Array.FindIndex(papeis, EhCabecalhoInsumo) + 1;
                }
            }
            catch (JsonException) { /* Sem metadados válidos, procurar o cabeçalho. */ }
        }
        if (coluna > 0)
        {
            // Metadados descrevem as colunas; o cabeçalho pode ter mudado de linha.
            var cabecalho = intervalo.Rows().Select(linha => aba.Cell(linha.RowNumber(), coluna))
                .FirstOrDefault(celula => EhCabecalhoInsumo(celula.GetString()));
            if (cabecalho != null) linhaCabecalho = cabecalho.Address.RowNumber;
        }
        if (coluna == 0)
            foreach (var linha in intervalo.Rows())
            {
                var celula = linha.Cells().FirstOrDefault(c => EhCabecalhoInsumo(c.GetString()));
                if (celula == null) continue;
                coluna = celula.Address.ColumnNumber;
                linhaCabecalho = celula.Address.RowNumber;
                break;
            }
        if (coluna == 0)
            throw new InvalidDataException("Não foi possível identificar o título da lista de insumos na planilha.");
        var nomes = new List<LinhaInsumoPlanilha>();
        for (int linha = linhaCabecalho + 1; linha <= intervalo.LastRow().RowNumber(); linha++)
        {
            string nome = aba.Cell(linha, coluna).GetString();
            if (!string.IsNullOrWhiteSpace(nome)) nomes.Add(new(linha, nome));
        }
        if (!nomes.Any(nome => ComparadorInsumos.NormalizarNome(nome.Nome, true).Length > 0))
            throw new InvalidDataException("A coluna de insumos da planilha está vazia.");
        return new(linhaCabecalho, coluna, intervalo.LastColumn().ColumnNumber(), nomes, papeisValidos);
    }

    private static bool EhCabecalhoInsumo(string? nome) =>
        Regex.IsMatch(ComparadorInsumos.NormalizarNome(nome), @"^INSUMOS?(?:$|[\s:–—-])",
            RegexOptions.CultureInvariant);
}
