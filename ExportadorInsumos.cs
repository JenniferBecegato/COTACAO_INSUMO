using System.Text.Encodings.Web;
using System.Text.Json;

namespace COTACAO_INSUMO;

public static class ExportadorInsumos
{
    public static async Task<string> SalvarAsync(ExtracaoInsumos resultado, string pdf, string pasta)
    {
        var json = JsonSerializer.Serialize(resultado, new JsonSerializerOptions
        {
            WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
        LeituraInsumos.Interpretar(json);
        Directory.CreateDirectory(pasta);
        string nome = Path.GetFileNameWithoutExtension(pdf);
        if (nome.Length > 80) nome = nome[..80];
        string destino = Path.Combine(pasta, $"{nome}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.json");
        return await GravarAsync(json, destino);
    }

    public static Task<string> SalvarCamposAsync(InsumosCampos resultado, string pdf, string pasta)
    {
        Directory.CreateDirectory(pasta);
        string nome = Path.GetFileNameWithoutExtension(pdf);
        if (nome.Length > 80) nome = nome[..80];
        string referencia = Path.Combine(pasta, $"{nome}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.json");
        return SalvarCamposAsync(resultado, referencia);
    }

    public static async Task<string> SalvarCamposAsync(InsumosCampos resultado, string caminhoOriginal)
    {
        string json = JsonSerializer.Serialize(resultado, new JsonSerializerOptions
        {
            WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
        string destino = Path.Combine(Path.GetDirectoryName(caminhoOriginal)!,
            Path.GetFileNameWithoutExtension(caminhoOriginal) + "_campos.json");
        return await GravarAsync(json, destino);
    }

    private static async Task<string> GravarAsync(string json, string destino)
    {
        string temporario = destino + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await File.WriteAllTextAsync(temporario, json);
            File.Move(temporario, destino);
            return destino;
        }
        finally
        {
            if (File.Exists(temporario)) File.Delete(temporario);
        }
    }
}
