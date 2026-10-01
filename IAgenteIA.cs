using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public interface IAgenteIA
    {
        Task<bool> TestarConexaoAsync();

        Task<ResultadoCotacaoIA?> AnalisarPdfAsync(
            string caminhoPdf,
            string listaInsumosExcel
        );
    }
}