using System.Globalization;
using System.Text.RegularExpressions;

namespace COTACAO_INSUMO;

public static class CalculadoraPreco
{
    // Reconhece somente a unidade e a embalagem explícitas, nunca o nome do produto.
    public static ResultadoPreco Calcular(string unidadeOriginal, decimal quantidade,
        decimal valorTotal, decimal valorUnitario)
    {
        if (quantidade <= 0 || valorTotal < 0 || valorUnitario < 0)
            return NaoIdentificado();

        string unidade = (unidadeOriginal ?? "").Trim().ToUpperInvariant();
        Match match = Regex.Match(unidade,
            @"^(?:(\d+(?:[.,]\d+)?)\s*)?(KG|QUILO[S]?|QUILOGRAMA[S]?|G|GR|GRS|GRAMA[S]?|MG|MILIGRAMA[S]?|ML|MILILITRO[S]?|L|LT|LITRO[S]?|MLH|MIL|MILHEIRO[S]?|UN|UND|UNID|UNIDADE[S]?|CAP|CAPS|CÁPSULA[S]?|CAPSULA[S]?)\.?$",
            RegexOptions.CultureInvariant);
        if (!match.Success) return NaoIdentificado();

        decimal embalagem = 1m;
        if (match.Groups[1].Success &&
            (!decimal.TryParse(match.Groups[1].Value.Replace(',', '.'),
                NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out embalagem) || embalagem <= 0))
            return NaoIdentificado();

        string medida = match.Groups[2].Value;
        decimal fator;
        string tipo, referencia;
        if (medida is "KG" || medida.StartsWith("QUILO"))
            (fator, tipo, referencia) = (1000m, "GRAMA", "g");
        else if (medida is "MG" || medida.StartsWith("MILIGRAMA"))
            (fator, tipo, referencia) = (0.001m, "GRAMA", "g");
        else if (medida is "G" or "GR" or "GRS" || medida.StartsWith("GRAMA"))
            (fator, tipo, referencia) = (1m, "GRAMA", "g");
        else if (medida is "ML" || medida.StartsWith("MILILITRO"))
            (fator, tipo, referencia) = (1m, "MILILITRO", "ml");
        else if (medida is "L" or "LT" || medida.StartsWith("LITRO"))
            (fator, tipo, referencia) = (1000m, "MILILITRO", "ml");
        else if (medida is "MLH" or "MIL" || medida.StartsWith("MILHEIRO"))
            (fator, tipo, referencia) = (1000m, "UNIDADE", "un");
        else
            (fator, tipo, referencia) = (1m, "UNIDADE", "un");

        decimal total = valorTotal > 0 ? valorTotal : quantidade * valorUnitario;
        if (total <= 0) return NaoIdentificado();
        // Total e unitário incompatíveis indicam leitura ambígua da linha.
        if (valorTotal > 0 && valorUnitario > 0 &&
            Math.Abs(valorTotal - quantidade * valorUnitario) > 0.02m)
            return NaoIdentificado();

        decimal preco = total / (quantidade * embalagem * fator);
        if (preco <= 0) return NaoIdentificado();
        return new ResultadoPreco { PrecoNormalizado = preco, TipoPreco = tipo,
            UnidadeReferencia = referencia, PodeConverter = true };
    }

    private static ResultadoPreco NaoIdentificado() => new()
    {
        TipoPreco = "NAO_IDENTIFICADO", PodeConverter = false
    };
}

public class ResultadoPreco
{
    public decimal PrecoNormalizado { get; set; }
    public string TipoPreco { get; set; } = "";
    public string UnidadeReferencia { get; set; } = "";
    public bool PodeConverter { get; set; }
}
