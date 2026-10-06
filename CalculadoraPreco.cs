using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace COTACAO_INSUMO
{
    public static class CalculadoraPreco
    {
        public static ResultadoPreco Calcular(string unidadeOriginal, decimal quantidade, decimal valorTotal, decimal valorUnitario)
        {
            var resultado = CalcularInterno(unidadeOriginal, quantidade, valorTotal, valorUnitario);
            if (resultado.PodeConverter) return resultado;
            string unidade = NormalizarUnidade(unidadeOriginal);
            if (UnidadeEhMg(unidade) || UnidadeEhKg(unidade) || UnidadeEhGrama(unidade))
            { resultado.TipoPreco = "GRAMA"; resultado.UnidadeReferencia = "g"; }
            else if (Regex.IsMatch(unidade, @"^(?:\d+(?:\.\d+)?\s*)?(ML|MILILITROS?|L|LT|LTS|LITROS?)$"))
            { resultado.TipoPreco = "MILILITRO"; resultado.UnidadeReferencia = "ml"; }
            else if (UnidadeEhMlh(unidade) || UnidadeEhMil(unidade) || Regex.IsMatch(unidade, @"^(?:\d+(?:\.\d+)?\s*)?(UN|UND|UNID|UNIDADE|UNIDADES)$"))
            { resultado.TipoPreco = "UNIDADE"; resultado.UnidadeReferencia = "un"; }
            return resultado;
        }

        private static ResultadoPreco CalcularInterno(
            string unidadeOriginal,
            decimal quantidade,
            decimal valorTotal,
            decimal valorUnitario)
        {
            string unidade =
                NormalizarUnidade(
                    unidadeOriginal
                );

            if (
                string.IsNullOrWhiteSpace(
                    unidade
                )
            )
            {
                return NaoIdentificado(
                    "Unidade não informada."
                );
            }

            // Recover a missing commercial unit price only from an explicit line total and quantity.
            if (valorUnitario <= 0 && valorTotal > 0 && quantidade > 0)
                valorUnitario = valorTotal / quantidade;

            // Liquids retain their volume basis; never assume a density to convert them to grams.
            if (Regex.IsMatch(unidade, @"^(?:\d+(?:\.\d+)?\s*)?(ML|MILILITROS?|L|LT|LTS|LITROS?)$"))
            {
                bool mililitros = Regex.IsMatch(unidade, @"(?:^|\s)(ML|MILILITROS?)$");
                decimal volume = ExtrairNumero(unidade);
                if (volume <= 0) volume = 1;
                decimal divisor = volume * (mililitros ? 1m : 1000m);
                return CriarResultado(valorUnitario / divisor, "MILILITRO", "ml");
            }
            if (Regex.IsMatch(unidade, @"^(?:\d+(?:\.\d+)?\s*)?(UN|UND|UNID|UNIDADE|UNIDADES)$"))
            {
                decimal unidades = ExtrairNumero(unidade);
                return CriarResultado(valorUnitario / (unidades > 0 ? unidades : 1m), "UNIDADE", "un");
            }

            // =================================================
            // REGRA 1 - PREÇO POR KG
            //
            // Exemplo:
            //
            // Valor KG = 580,00
            //
            // 580 / 1000
            // = 0,580 por grama
            // =================================================

            if (
                UnidadeEhKg(
                    unidade
                )
            )
            {
                if (valorUnitario <= 0)
                {
                    return NaoIdentificado(
                        "Preço por KG não identificado."
                    );
                }

                decimal quilosEmbalagem = ExtrairNumero(unidade);
                decimal precoPorGrama = valorUnitario / ((quilosEmbalagem > 0 ? quilosEmbalagem : 1m) * 1000m);

                return CriarResultado(
                    precoPorGrama,
                    "GRAMA",
                    "g"
                );
            }

            // =================================================
            // REGRA 2 - EMBALAGEM EM GRAMAS
            //
            // Exemplo:
            //
            // UnidadeOriginal = 250 G
            // Preço = 109,00
            //
            // 109 / 250
            // = 0,436 por grama
            // =================================================

            if (
                UnidadeEhGrama(
                    unidade
                )
            )
            {
                decimal gramas =
                    ExtrairNumero(
                        unidade
                    );

                // ---------------------------------------------
                // "250 G", "100 G", "50 G"...
                // ---------------------------------------------

                if (gramas > 0)
                {
                    if (valorUnitario <= 0)
                    {
                        return NaoIdentificado(
                            "Preço da embalagem não identificado."
                        );
                    }

                    decimal precoPorGrama =
                        valorUnitario /
                        gramas;

                    return CriarResultado(
                        precoPorGrama,
                        "GRAMA",
                        "g"
                    );
                }

                // ---------------------------------------------
                // Se vier apenas G / GR / GRAMA,
                // considera que o preço já é por grama.
                // ---------------------------------------------

                if (valorUnitario > 0)
                {
                    return CriarResultado(
                        valorUnitario,
                        "GRAMA",
                        "g"
                    );
                }

                return NaoIdentificado(
                    "Preço em gramas não identificado."
                );
            }

            // =================================================
            // REGRA 3 - MILIGRAMAS
            //
            // Exemplo:
            //
            // 500 MG por R$ 10
            //
            // 500 mg = 0,5 g
            //
            // 10 / 0,5
            // = 20 por grama
            // =================================================

            if (
                UnidadeEhMg(
                    unidade
                )
            )
            {
                decimal mg =
                    ExtrairNumero(
                        unidade
                    );

                // ---------------------------------------------
                // Embalagem, ex.: 500 MG
                // ---------------------------------------------

                if (mg > 0)
                {
                    if (valorUnitario <= 0)
                    {
                        return NaoIdentificado(
                            "Preço da embalagem em MG não identificado."
                        );
                    }

                    decimal gramas =
                        mg /
                        1000m;

                    if (gramas <= 0)
                    {
                        return NaoIdentificado(
                            "Quantidade em MG inválida."
                        );
                    }

                    decimal precoPorGrama =
                        valorUnitario /
                        gramas;

                    return CriarResultado(
                        precoPorGrama,
                        "GRAMA",
                        "g"
                    );
                }

                // ---------------------------------------------
                // Se vier somente MG e o preço realmente
                // representar 1 mg:
                //
                // preço por grama = preço por mg * 1000
                // ---------------------------------------------

                if (valorUnitario > 0)
                {
                    decimal precoPorGrama =
                        valorUnitario *
                        1000m;

                    return CriarResultado(
                        precoPorGrama,
                        "GRAMA",
                        "g"
                    );
                }

                return NaoIdentificado(
                    "Preço em MG não identificado."
                );
            }

            // =================================================
            // REGRA 4 - MLH / MILHEIRO
            //
            // Exemplo:
            //
            // R$ 24,70 por MLH
            //
            // 24,70 / 1000
            // = 0,0247 por unidade
            // =================================================

            if (
                UnidadeEhMlh(
                    unidade
                )
            )
            {
                if (valorUnitario <= 0)
                {
                    return NaoIdentificado(
                        "Preço por milheiro não identificado."
                    );
                }

                decimal precoPorUnidade =
                    valorUnitario /
                    1000m;

                return CriarResultado(
                    precoPorUnidade,
                    "UNIDADE",
                    "un"
                );
            }

            // =================================================
            // REGRA 5 - "5 MIL", "10 MIL" ETC.
            //
            // Exemplo:
            //
            // 5 MIL cápsulas = 5000
            // Valor embalagem = R$ 99,50
            //
            // 99,50 / 5000
            // =================================================

            if (
                UnidadeEhMil(
                    unidade
                )
            )
            {
                decimal quantidadeMil =
                    ExtrairNumero(
                        unidade
                    );

                if (quantidadeMil <= 0)
                {
                    quantidadeMil =
                        1m;
                }

                if (valorUnitario <= 0)
                {
                    return NaoIdentificado(
                        "Preço das unidades não identificado."
                    );
                }

                decimal totalUnidades =
                    quantidadeMil *
                    1000m;

                decimal precoPorUnidade =
                    valorUnitario /
                    totalUnidades;

                return CriarResultado(
                    precoPorUnidade,
                    "UNIDADE",
                    "un"
                );
            }

            // =================================================
            // NÃO IDENTIFICADO
            //
            // Exemplo:
            // ML
            // L
            // CX
            // FR
            // KIT
            //
            // enquanto não definirmos regra.
            // =================================================

            return NaoIdentificado(
                "Unidade não reconhecida: " +
                unidadeOriginal
            );
        }

        // =====================================================
        // RESULTADO VÁLIDO
        // =====================================================

        private static ResultadoPreco CriarResultado(
            decimal preco,
            string tipo,
            string unidadeReferencia)
        {
            if (preco <= 0)
            {
                return NaoIdentificado(
                    "O preço calculado é inválido."
                );
            }

            return new ResultadoPreco
            {
                PrecoNormalizado =
                    preco,

                TipoPreco =
                    tipo,

                UnidadeReferencia =
                    unidadeReferencia,

                PodeConverter =
                    true,

                Motivo =
                    ""
            };
        }

        // =====================================================
        // NÃO IDENTIFICADO
        // =====================================================

        private static ResultadoPreco NaoIdentificado(
            string motivo)
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
                    false,

                Motivo =
                    motivo
            };
        }

        // =====================================================
        // NORMALIZAR UNIDADE
        // =====================================================

        private static string NormalizarUnidade(
            string texto)
        {
            if (
                string.IsNullOrWhiteSpace(
                    texto
                )
            )
            {
                return "";
            }

            string resultado =
                texto
                    .Trim()
                    .ToUpperInvariant()
                    .Replace(",", ".")
                    .Replace("\r", " ")
                    .Replace("\n", " ");

            while (
                resultado.Contains("  ")
            )
            {
                resultado =
                    resultado.Replace(
                        "  ",
                        " "
                    );
            }

            // Accept compact packaging labels (1KG, 50G, 500MG, 800ML).
            resultado = Regex.Replace(resultado, @"(?<=\d)(?=[A-Z])", " ");
            return resultado;
        }

        // =====================================================
        // EXTRAIR NÚMERO
        //
        // 250 G  -> 250
        // 5 MIL  -> 5
        // 500 MG -> 500
        // =====================================================

        private static decimal ExtrairNumero(
            string texto)
        {
            Match match =
                Regex.Match(
                    texto,
                    @"(\d+(?:[.,]\d+)?)"
                );

            if (!match.Success)
            {
                return 0m;
            }

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

        // =====================================================
        // KG
        // =====================================================

        private static bool UnidadeEhKg(
            string unidade)
        {
            // Para evitar pegar algo indevido,
            // procuramos KG como unidade.
            return
                Regex.IsMatch(
                    unidade,
                    @"(^|\s)KG($|\s)"
                )
                ||
                unidade == "KG"
                ||
                unidade.Contains(
                    "QUILOGRAMA"
                )
                ||
                unidade.Contains(
                    "QUILO"
                );
        }

        // =====================================================
        // GRAMA
        // =====================================================

        private static bool UnidadeEhGrama(
            string unidade)
        {
            return
                Regex.IsMatch(
                    unidade,
                    @"(^|\s)G($|\s)"
                )
                ||
                Regex.IsMatch(
                    unidade,
                    @"(^|\s)GR($|\s)"
                )
                ||
                Regex.IsMatch(unidade, @"(?:^|\s)GRAMAS?(?:$|\s)");
        }

        // =====================================================
        // MG
        // =====================================================

        private static bool UnidadeEhMg(
            string unidade)
        {
            return
                Regex.IsMatch(
                    unidade,
                    @"(^|\s)MG($|\s)"
                )
                ||
                unidade.Contains(
                    "MILIGRAMA"
                );
        }

        // =====================================================
        // MLH
        // =====================================================

        private static bool UnidadeEhMlh(
            string unidade)
        {
            return
                unidade == "MH"
                || unidade == "MLH"
                ||
                unidade.Contains(
                    "MILHEIRO"
                );
        }

        // =====================================================
        // MIL
        // =====================================================

        private static bool UnidadeEhMil(
            string unidade)
        {
            return
                Regex.IsMatch(
                    unidade,
                    @"(^|\s)MIL($|\s)"
                );
        }
    }

    // =========================================================
    // RESULTADO
    // =========================================================

    public class ResultadoPreco
    {
        public decimal PrecoNormalizado { get; set; }

        public string TipoPreco { get; set; }
            = "";

        public string UnidadeReferencia { get; set; }
            = "";

        public bool PodeConverter { get; set; }

        public string Motivo { get; set; }
            = "";
    }
}