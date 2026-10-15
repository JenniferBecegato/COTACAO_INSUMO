using System.Text.RegularExpressions;

namespace COTACAO_INSUMO;

public static class NormalizadorCotacao
{
    // O vínculo com o Excel nunca participa da identificação da base de preço.
    public static ResultadoPreco Calcular(ItemCotacaoIA item)
    {
        if (string.IsNullOrWhiteSpace(item.UnidadeOriginal)
            || Regex.IsMatch(item.UnidadeOriginal.Trim(), @"^\d+(?:[.,]\d+)?$"))
        {
            var embalagem = Regex.Match(item.ProdutoPdf ?? "",
                @"(?i)(?:^|\s)(\d+(?:[.,]\d+)?\s*(?:KG|MG|GR|G|ML|LT|L|MIL|MLH))\s*$");
            if (embalagem.Success)
            {
                string numeroAnterior = (item.UnidadeOriginal ?? "").Trim().Replace(',', '.');
                string numeroEmbalagem = Regex.Match(embalagem.Groups[1].Value, @"\d+(?:[.,]\d+)?").Value.Replace(',', '.');
                if (string.IsNullOrWhiteSpace(numeroAnterior) || numeroAnterior == numeroEmbalagem)
                    item.UnidadeOriginal = embalagem.Groups[1].Value;
            }
        }

        var preco = CalculadoraPreco.Calcular(item.UnidadeOriginal ?? "",
            item.Quantidade, item.ValorTotal, item.ValorUnitario);
        // Total e quantidade servem apenas para conferir a extração do preço.
        // Uma divergência exige releitura do campo comercial, não outro cálculo de preço.
        var medida = CalculadoraPreco.Calcular(item.UnidadeQuantidade ?? "", 1, 0, 1);
        if (medida.PodeConverter && item.Quantidade > 0 && item.ValorTotal > 0)
        {
            decimal quantidadeBase = item.Quantidade / medida.PrecoNormalizado;
            bool mesmaDimensao = preco.TipoPreco == medida.TipoPreco;
            if (preco.PodeConverter && mesmaDimensao
                && Math.Abs(preco.PrecoNormalizado * quantidadeBase - item.ValorTotal) > 0.02m)
            {
                preco.PodeConverter = false;
                preco.PrecoNormalizado = 0;
                preco.Motivo = "Preço e base extraídos não conferem com a linha. Releia o mesmo campo comercial de preço e seu cabeçalho no PDF.";
            }
        }
        item.PrecoNormalizado = preco.PrecoNormalizado;
        item.TipoPreco = preco.TipoPreco;
        return preco;
    }

    public static List<string> Validar(ResultadoCotacaoIA? cotacao)
    {
        if (cotacao?.Itens == null || cotacao.Itens.Count == 0)
            return ["Nenhum item foi extraído do PDF."];

        var pendencias = new List<string>();
        foreach (var item in cotacao.Itens)
        {
            var preco = Calcular(item);
            if (!preco.PodeConverter)
                pendencias.Add($"{item.ProdutoPdf}: {preco.Motivo} "
                    + $"Base extraída='{item.UnidadeOriginal}', preço extraído={item.ValorUnitario}.");
        }
        return pendencias;
    }

    public static ResultadoCotacaoIA? Complementar(ResultadoCotacaoIA? original, ResultadoCotacaoIA? releitura)
    {
        if (original?.Itens == null || original.Itens.Count == 0)
            return releitura ?? original;
        if (releitura?.Itens == null || releitura.Itens.Count == 0)
            return original;

        // Compare ocorrências para preservar também produtos repetidos no orçamento.
        var candidatos = releitura.Itens
            .GroupBy(i => (i.ProdutoPdf ?? "").Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => new Queue<ItemCotacaoIA>(g), StringComparer.OrdinalIgnoreCase);
        var itens = new List<ItemCotacaoIA>();
        foreach (var item in original.Itens)
        {
            ItemCotacaoIA? alternativa = null;
            if (candidatos.TryGetValue((item.ProdutoPdf ?? "").Trim(), out var fila) && fila.Count > 0)
                alternativa = fila.Dequeue();
            if (!Calcular(item).PodeConverter && alternativa != null)
            {
                // A releitura pode recuperar só o preço; preserve a unidade já lida
                // quando ela não vier novamente, sem substituir preço pelo total.
                var complemento = new ItemCotacaoIA
                {
                    ProdutoPdf = item.ProdutoPdf ?? "",
                    InsumoExcel = item.InsumoExcel,
                    Encontrado = item.Encontrado,
                    Confianca = item.Confianca,
                    UnidadeOriginal = string.IsNullOrWhiteSpace(alternativa.UnidadeOriginal)
                        ? item.UnidadeOriginal : alternativa.UnidadeOriginal,
                    ValorUnitario = alternativa.ValorUnitario > 0 ? alternativa.ValorUnitario : item.ValorUnitario,
                    Quantidade = alternativa.Quantidade > 0 ? alternativa.Quantidade : item.Quantidade,
                    UnidadeQuantidade = string.IsNullOrWhiteSpace(alternativa.UnidadeQuantidade)
                        ? item.UnidadeQuantidade : alternativa.UnidadeQuantidade,
                    ValorTotal = alternativa.ValorTotal > 0 ? alternativa.ValorTotal : item.ValorTotal
                };
                if (Calcular(complemento).PodeConverter
                    || (item.ValorUnitario <= 0 && complemento.ValorUnitario > 0)
                    || (string.IsNullOrWhiteSpace(item.UnidadeOriginal)
                        && !string.IsNullOrWhiteSpace(complemento.UnidadeOriginal)))
                {
                    itens.Add(complemento);
                    continue;
                }
            }
            itens.Add(item);
        }
        foreach (var fila in candidatos.Values)
            itens.AddRange(fila);
        original.Itens = itens;
        if (string.IsNullOrWhiteSpace(original.Fornecedor))
            original.Fornecedor = releitura.Fornecedor;
        return original;
    }
}
