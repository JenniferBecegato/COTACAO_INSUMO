using System;

namespace COTACAO_INSUMO
{
    public static class CalculadoraPreco
    {
        public static ResultadoPreco Calcular(
            string unidade,
            decimal valorUnitario)
        {
            string unidadeNormalizada =
                (unidade ?? "")
                .Trim()
                .ToUpperInvariant();

            switch (unidadeNormalizada)
            {
                // =============================================
                // PREÇO POR GRAMA
                // =============================================

                case "KG":
                case "KILO":
                case "QUILO":
                case "QUILOGRAMA":
                    return new ResultadoPreco
                    {
                        PrecoNormalizado =
                            valorUnitario / 1000m,

                        TipoPreco =
                            "GRAMA",

                        UnidadeReferencia =
                            "g",

                        PodeConverter =
                            true
                    };

                case "G":
                case "GR":
                case "GRAMA":
                case "GRAMAS":
                    return new ResultadoPreco
                    {
                        PrecoNormalizado =
                            valorUnitario,

                        TipoPreco =
                            "GRAMA",

                        UnidadeReferencia =
                            "g",

                        PodeConverter =
                            true
                    };

                case "MG":
                case "MILIGRAMA":
                case "MILIGRAMAS":
                    return new ResultadoPreco
                    {
                        // se o preço é por mg,
                        // 1 g possui 1000 mg
                        PrecoNormalizado =
                            valorUnitario * 1000m,

                        TipoPreco =
                            "GRAMA",

                        UnidadeReferencia =
                            "g",

                        PodeConverter =
                            true
                    };

                // =============================================
                // CÁPSULA / MILHEIRO
                // =============================================

                case "MLH":
                case "MILHEIRO":
                case "MIL":
                    return new ResultadoPreco
                    {
                        // 1 MLH = 1000 cápsulas
                        PrecoNormalizado =
                            valorUnitario / 1000m,

                        TipoPreco =
                            "UNIDADE",

                        UnidadeReferencia =
                            "un",

                        PodeConverter =
                            true
                    };

                // =============================================
                // UNIDADE DESCONHECIDA
                // =============================================

                default:
                    return new ResultadoPreco
                    {
                        PrecoNormalizado =
                            0m,

                        TipoPreco =
                            "NAO_IDENTIFICADO",

                        UnidadeReferencia =
                            unidadeNormalizada,

                        PodeConverter =
                            false
                    };
            }
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