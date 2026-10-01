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
        // Nome exatamente como aparece no PDF
        public string ProdutoPdf { get; set; } = "";

        // Nome correspondente encontrado no Excel
        public string InsumoExcel { get; set; } = "";

        // Unidade original do orçamento:
        // KG, G, GR, MG, MLH...
        public string UnidadeOriginal { get; set; } = "";

        // Quantidade apresentada no PDF
        public decimal Quantidade { get; set; }

        // Valor unitário apresentado no PDF
        public decimal ValorUnitario { get; set; }

        // Valor total apresentado no PDF
        public decimal ValorTotal { get; set; }

        // Valor normalizado que irá para a planilha
        public decimal PrecoNormalizado { get; set; }

        // GRAMA ou UNIDADE
        public string TipoPreco { get; set; } = "";

        // A IA encontrou correspondência no Excel?
        public bool Encontrado { get; set; }

        // Confiança de 0 a 1
        public decimal Confianca { get; set; }
    }
}