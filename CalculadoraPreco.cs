using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace COTACAO_INSUMO
{
    public static class CalculadoraPreco
    {
        public static ResultadoPreco Calcular(
            string unidadeOriginal,
            decimal quantidade,
            decimal valorTotal,
            decimal valorUnitario)
        {
            string unidade =
                (unidadeOriginal ?? "")
                .Trim()
                .ToUpperInvariant();

            unidade =
                unidade.Replace(",", ".");

            // =================================================
            // TENTAR DESCOBRIR NÚMERO EMBUTIDO NA UNIDADE
            //
            // Exemplos:
            // 250 G
            // 100 G
            // 2 KG
            // 5 MIL
            // =================================================

            decimal quantidadePorItem =
                ExtrairQuantidadeDaUnidade(
                    unidade
                );

            // Se não houver número:
            // KG -> 1 KG
            // G  -> 1 G
            // MLH -> 1 milheiro
            if (quantidadePorItem <= 0)
            {
                quantidadePorItem =
                    1m;
            }

            // =================================================
            // VALOR A USAR
            // =================================================

            decimal total =
                valorTotal;

            // Caso a IA não tenha conseguido obter o total,
            // tenta quantidade x valor unitário.
            if (
                total <= 0 &&
                valorUnitario > 0
            )
            {
                total =
                    quantidade > 0
                        ? quantidade * valorUnitario
                        : valorUnitario;
            }

            if (total <= 0)
            {
                return NaoIdentificado();
            }

            decimal qtdLinha =
                quantidade > 0
                    ? quantidade
                    : 1m;

            // =================================================
            // CÁPSULAS / MILHEIRO
            // =================================================

            if (
                unidade.Contains("MLH") ||
                unidade.Contains("MILHEIRO") ||
                unidade.Contains("MIL")
            )
            {
                decimal unidadesPorItem;

                // 5 MIL = 5000 cápsulas
                if (
                    unidade.Contains("MIL") &&
                    quantidadePorItem > 1
                )
                {
                    unidadesPorItem =
                        quantidadePorItem * 1000m;
                }

                // MLH sem número = 1000 cápsulas
                else
                {
                    unidadesPorItem =
                        1000m;
                }

                decimal totalUnidades =
                    qtdLinha *
                    unidadesPorItem;

                if (totalUnidades <= 0)
                    return NaoIdentificado();

                return new ResultadoPreco
                {
                    PrecoNormalizado =
                        total / totalUnidades,

                    TipoPreco =
                        "UNIDADE",

                    UnidadeReferencia =
                        "un",

                    PodeConverter =
                        true
                };
            }

            // =================================================
            // KG
            // =================================================

            if (
                unidade.Contains("KG") ||
                unidade.Contains("QUILO")
            )
            {
                decimal quilos;

                // Exemplo:
                // Fagron: Qtde 1 / Unidade 2 KG
                if (
                    TemNumeroNaUnidade(
                        unidade
                    )
                )
                {
                    quilos =
                        qtdLinha *
                        quantidadePorItem;
                }

                // Exemplo:
                // Sixty: Qtde 0,500 / Unidade KG
                else
                {
                    quilos =
                        qtdLinha;
                }

                decimal totalGramas =
                    quilos * 1000m;

                if (totalGramas <= 0)
                    return NaoIdentificado();

                return new ResultadoPreco
                {
                    PrecoNormalizado =
                        total / totalGramas,

                    TipoPreco =
                        "GRAMA",

                    UnidadeReferencia =
                        "g",

                    PodeConverter =
                        true
                };
            }

            // =================================================
            // MG
            // =================================================

            if (
                Regex.IsMatch(
                    unidade,
                    @"\bMG\b"
                )
            )
            {
                decimal miligramas;

                if (
                    TemNumeroNaUnidade(
                        unidade
                    )
                )
                {
                    miligramas =
                        qtdLinha *
                        quantidadePorItem;
                }
                else
                {
                    miligramas =
                        qtdLinha;
                }

                decimal totalGramas =
                    miligramas / 1000m;

                if (totalGramas <= 0)
                    return NaoIdentificado();

                return new ResultadoPreco
                {
                    PrecoNormalizado =
                        total / totalGramas,

                    TipoPreco =
                        "GRAMA",

                    UnidadeReferencia =
                        "g",

                    PodeConverter =
                        true
                };
            }

            // =================================================
            // GRAMAS
            // =================================================

            if (
                Regex.IsMatch(
                    unidade,
                    @"\bG\b"
                )
                ||
                unidade.Contains("GR")
                ||
                unidade.Contains("GRAMA")
            )
            {
                decimal totalGramas;

                // Fagron:
                // qtd 1
                // unidade 250 G
                if (
                    TemNumeroNaUnidade(
                        unidade
                    )
                )
                {
                    totalGramas =
                        qtdLinha *
                        quantidadePorItem;
                }

                // Outro fornecedor:
                // qtd 250
                // unidade G
                else
                {
                    totalGramas =
                        qtdLinha;
                }

                if (totalGramas <= 0)
                    return NaoIdentificado();

                return new ResultadoPreco
                {
                    PrecoNormalizado =
                        total / totalGramas,

                    TipoPreco =
                        "GRAMA",

                    UnidadeReferencia =
                        "g",

                    PodeConverter =
                        true
                };
            }

            // =================================================
            // NÃO CONHECIDO
            // =================================================

            return NaoIdentificado();
        }

        // =====================================================
        // EXTRAIR QUANTIDADE
        //
        // "250 G" -> 250
        // "2 KG"  -> 2
        // "5 MIL" -> 5
        // =====================================================

        private static decimal ExtrairQuantidadeDaUnidade(
            string unidade)
        {
            Match match =
                Regex.Match(
                    unidade,
                    @"(\d+(?:[.,]\d+)?)"
                );

            if (!match.Success)
                return 0m;

            string numero =
                match.Groups[1]
                    .Value
                    .Replace(",", ".");

            if (
                decimal.TryParse(
                    numero,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out decimal valor
                )
            )
            {
                return valor;
            }

            return 0m;
        }

        private static bool TemNumeroNaUnidade(
            string unidade)
        {
            return Regex.IsMatch(
                unidade,
                @"\d"
            );
        }

        private static ResultadoPreco NaoIdentificado()
        {
            return new ResultadoPreco
            {
                PrecoNormalizado =
                    0m,

                TipoPreco =
                    "NAO_IDENTIFICADO",

                UnidadeReferencia =
                    "",

                PodeConverter =
                    false
            };
        }
    }

    public class ResultadoPreco
    {
        public decimal PrecoNormalizado { get; set; }

        public string TipoPreco { get; set; } = "";

        public string UnidadeReferencia { get; set; } = "";

        public bool PodeConverter { get; set; }
    }
}