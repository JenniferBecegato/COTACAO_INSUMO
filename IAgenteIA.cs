namespace COTACAO_INSUMO;

public interface IAgenteIA
{
    Task<bool> TestarConexaoAsync();
    Task<ExtracaoInsumos> ExtrairInsumosAsync(string caminhoPdf);
}
