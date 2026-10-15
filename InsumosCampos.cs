using System.Globalization;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace COTACAO_INSUMO;

public sealed class InsumosCampos
{
    [JsonPropertyName("fornecedor")]
    public required string Fornecedor { get; init; }
    public required List<InsumoCampo> Insumos { get; init; }
}

public sealed class InsumoCampo
{
    public required string Nome { get; init; }
    public required decimal? Preco { get; init; }
    public required string? Unidad { get; init; }
    public required bool Encontrado { get; init; }
    public required decimal? PrecoUnitario { get; init; }
}

public static class ConversorPrecoUnitario
{
    // Kg -> G, L -> ML e Milheiro -> Unidade. Não altera o preço/base original.
    public static decimal? Converter(decimal? preco, string? unidade)
    {
        if (preco == null) return null;
        decimal? convertido = unidade switch
        {
            "Kg" or "L" or "Milheiro" => preco / 1000m,
            "G" or "ML" or "Unidade" => preco,
            _ => null
        };
        return convertido.HasValue
            ? Math.Round(convertido.Value, 4, MidpointRounding.AwayFromZero)
            : null;
    }
}

public static class ComparadorInsumos
{
    // Percentual mínimo de correspondência, configurável entre > 0 e 100.
    public const decimal PercentualMinimoMatch = 75m;

    public static string NormalizarNome(string? nome, bool nomeDaPlanilha = false)
    {
        string texto = nome ?? "";
        if (nomeDaPlanilha)
            texto = Regex.Replace(texto, @"\s*\(\s*Urgente\s*\)\s*$", "",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        var semAcentos = new StringBuilder();
        foreach (char caractere in texto.Normalize(NormalizationForm.FormD))
            if (CharUnicodeInfo.GetUnicodeCategory(caractere) != UnicodeCategory.NonSpacingMark)
                semAcentos.Append(caractere);
        return Regex.Replace(semAcentos.ToString().Normalize(NormalizationForm.FormC), @"\s+", " ")
            .Trim().ToUpperInvariant();
    }

    public static bool MatchInsumo(string? nomeJson, IEnumerable<string> nomesPlanilha,
        decimal percentualMinimo = PercentualMinimoMatch)
    {
        ValidarPercentual(percentualMinimo);
        string nome = NormalizarNome(nomeJson);
        return nomesPlanilha.Any(candidato => Corresponde(nome, NormalizarNome(candidato, true), percentualMinimo));
    }

    public static InsumosCampos Comparar(ExtracaoInsumos resultado, IEnumerable<string> nomesPlanilha,
        decimal percentualMinimo = PercentualMinimoMatch)
    {
        ValidarPercentual(percentualMinimo);
        var nomes = nomesPlanilha.Select(x => NormalizarNome(x, true)).Where(x => x.Length > 0)
            .Distinct(StringComparer.Ordinal).ToArray();
        if (nomes.Length == 0)
            throw new InvalidDataException("A planilha não contém nomes de insumos válidos para comparação.");
        return new InsumosCampos
        {
            Fornecedor = resultado.Fornecedor,
            Insumos = resultado.Insumos.Select(item => new InsumoCampo
            {
                Nome = item.Nome, Preco = item.Preco, Unidad = item.Unidad,
                PrecoUnitario = ConversorPrecoUnitario.Converter(item.Preco, item.Unidad),
                Encontrado = nomes.Any(nome => Corresponde(NormalizarNome(item.Nome), nome, percentualMinimo))
            }).ToList()
        };
    }

    private static void ValidarPercentual(decimal percentual)
    {
        if (percentual <= 0 || percentual > 100)
            throw new ArgumentOutOfRangeException(nameof(percentual), "O percentual deve ser maior que zero e até 100.");
    }

    private static bool Corresponde(string a, string b, decimal percentual)
    {
        return a.Length > 0 && b.Length > 0 && CalcularSimilaridade(a, b) >= percentual;
    }

    // Recebe nomes já normalizados; compartilhado pelo match e pelo preenchimento.
    public static decimal CalcularSimilaridade(string a, string b)
    {
        if (a.Length == 0 || b.Length == 0) return 0;
        if (a == b) return 100;
        if (a.Length > b.Length) (a, b) = (b, a);
        // Maior substring comum, com memória proporcional ao menor nome.
        int[] comprimentos = new int[a.Length + 1];
        int maior = 0;
        foreach (char caractere in b)
            for (int j = a.Length; j > 0; j--)
            {
                comprimentos[j] = a[j - 1] == caractere ? comprimentos[j - 1] + 1 : 0;
                maior = Math.Max(maior, comprimentos[j]);
            }
        return 100m * maior / a.Length;
    }
}
