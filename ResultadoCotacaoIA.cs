using System.Collections.Generic;

namespace COTACAO_INSUMO
{
    public class ResultadoCotacaoIA
    {
        public string Fornecedor { get; set; } = "";

        public List<ItemCotacaoIA> Itens { get; set; }
            = new List<ItemCotacaoIA>();
    }

    public class ItemCotacaoIA
    {
        public string ProdutoPdf { get; set; } = "";

        public string InsumoExcel { get; set; } = "";

        // Ex:
        // KG
        // 250 G
        // 2 KG
        // 5 MIL
        // MLH
        public string UnidadeOriginal { get; set; } = "";

        // Quantidade da linha do orçamento.
        // Ex:
        // 0.500
        // 1
        // 5
        public decimal Quantidade { get; set; }

        public decimal ValorUnitario { get; set; }

        public decimal ValorTotal { get; set; }

        public decimal PrecoNormalizado { get; set; }

        public string TipoPreco { get; set; } = "";

        public bool Encontrado { get; set; }

        public decimal Confianca { get; set; }
    }
}