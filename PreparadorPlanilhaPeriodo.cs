namespace COTACAO_INSUMO;

public static class PreparadorPlanilhaPeriodo
{
    public static void Preparar(string planilhaBase, string destino)
    {
        if (File.Exists(destino)) return;
        if (!File.Exists(planilhaBase))
            throw new FileNotFoundException("Selecione a planilha-base para criar a cotação do período.", planilhaBase);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(destino))!);
        string temporario = destino + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.Copy(planilhaBase, temporario);
            File.Move(temporario, destino);
        }
        finally
        {
            if (File.Exists(temporario)) File.Delete(temporario);
        }
    }
}
