namespace COTACAO_INSUMO;

public static class LocalBanco
{
    private static readonly string PastaPadrao = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "COTACAO_INSUMO");
    private static readonly string ArquivoLocal = Path.Combine(PastaPadrao, "local-banco.txt");
    // Mantém todas as operações da sessão no mesmo banco até reiniciar.
    public static string CaminhoAtual { get; } = Resolver(CarregarPasta());

    public static string CarregarPasta() => File.Exists(ArquivoLocal) ? File.ReadAllText(ArquivoLocal).Trim() : "";

    private static string Resolver(string pasta) => Path.Combine(
        string.IsNullOrWhiteSpace(pasta) ? PastaPadrao : pasta, "configuracoes.db");

    public static void SalvarPasta(string pasta)
    {
        pasta = pasta.Trim();
        if (pasta.Length > 0 && !Path.IsPathFullyQualified(pasta))
            throw new ArgumentException("Informe o caminho completo da pasta do banco.");
        if (pasta.Length > 0) pasta = Path.GetFullPath(pasta);
        string destino = Resolver(pasta);
        Directory.CreateDirectory(Path.GetDirectoryName(destino)!);
        if (!string.Equals(Path.GetFullPath(CaminhoAtual), Path.GetFullPath(destino), StringComparison.OrdinalIgnoreCase))
        {
            if (!File.Exists(destino) && File.Exists(CaminhoAtual))
            {
                using var origem = new ConfiguracaoIARepository.Banco(CaminhoAtual);
                using var copiar = origem.Preparar("VACUUM INTO ?1;");
                copiar.Vincular(1, destino);
                copiar.Ler();
            }
            using var validar = new ConfiguracaoIARepository.Banco(destino);
        }
        Directory.CreateDirectory(PastaPadrao);
        string temporario = ArquivoLocal + ".tmp";
        File.WriteAllText(temporario, pasta);
        File.Move(temporario, ArquivoLocal, true);
    }
}
