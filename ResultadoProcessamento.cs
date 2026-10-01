using System.Collections.Generic;

namespace COTACAO_INSUMO
{
    public class ResultadoProcessamento
    {
        public int TotalPreenchidos { get; set; }

        public List<ItemNaoEncontradoProcessado> NaoEncontrados
        { get; set; }
            = new List<ItemNaoEncontradoProcessado>();
    }

    public class ItemNaoEncontradoProcessado
    {
        public string Fornecedor { get; set; } = "";

        public string Insumo { get; set; } = "";

        public decimal PrecoNormalizado { get; set; }

        public string TipoPreco { get; set; } = "";

        public string UnidadeOriginal { get; set; } = "";
    }
}