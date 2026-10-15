using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace COTACAO_INSUMO;

public sealed class CatalogoModelosIA
{
    private readonly string pasta;
    public CatalogoModelosIA(string? pastaCache = null) => pasta = pastaCache ?? Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "COTACAO_INSUMO", "modelos");

    // O cache é separado por credencial e nunca contém a chave da API.
    private string Caminho(string provedor, string chave) => Path.Combine(pasta,
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(provedor + "\n" + chave))) + ".json");

    public string[] LerCache(string provedor, string chave)
    {
        try { return JsonSerializer.Deserialize<string[]>(File.ReadAllText(Caminho(provedor, chave))) ?? []; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return []; }
    }

    public async Task<string[]> ConsultarAsync(string provedor, string chave, CancellationToken cancelamento)
    {
        using var cliente = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
        string url = provedor switch
        {
            "OpenAI" => "https://api.openai.com/v1/models",
            "Claude (Anthropic)" => "https://api.anthropic.com/v1/models?limit=100",
            "Gemini (Google)" => "https://generativelanguage.googleapis.com/v1beta/models?pageSize=1000",
            _ => throw new ArgumentException("Provedor desconhecido.")
        };
        if (provedor == "OpenAI") cliente.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", chave);
        else if (provedor == "Claude (Anthropic)")
        {
            cliente.DefaultRequestHeaders.Add("x-api-key", chave);
            cliente.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
        }
        else cliente.DefaultRequestHeaders.Add("x-goog-api-key", chave);

        var modelos = new HashSet<string>(StringComparer.Ordinal);
        var paginas = new HashSet<string>(StringComparer.Ordinal);
        string? proxima = url;
        while (proxima != null)
        {
            if (!paginas.Add(proxima)) throw new InvalidDataException("Paginação inválida do catálogo.");
            using var resposta = await cliente.GetAsync(proxima, cancelamento);
            resposta.EnsureSuccessStatusCode();
            using var documento = JsonDocument.Parse(await resposta.Content.ReadAsStringAsync(cancelamento));
            var raiz = documento.RootElement;
            bool gemini = provedor == "Gemini (Google)";
            foreach (var item in raiz.GetProperty(gemini ? "models" : "data").EnumerateArray())
            {
                string id = item.GetProperty(gemini ? "name" : "id").GetString() ?? "";
                if (gemini) id = id.StartsWith("models/", StringComparison.Ordinal) ? id[7..] : id;
                if (Compativel(provedor, id, item)) modelos.Add(id);
            }
            proxima = null;
            if (gemini && raiz.TryGetProperty("nextPageToken", out var token) && !string.IsNullOrEmpty(token.GetString()))
                proxima = url + "&pageToken=" + Uri.EscapeDataString(token.GetString()!);
            else if (provedor == "Claude (Anthropic)" && raiz.TryGetProperty("has_more", out var mais) && mais.GetBoolean())
                proxima = url + "&after_id=" + Uri.EscapeDataString(raiz.GetProperty("last_id").GetString()!);
        }
        var resultado = modelos.Order(StringComparer.Ordinal).ToArray();
        if (resultado.Length == 0) throw new InvalidDataException("Nenhum modelo de texto compatível foi retornado.");
        cancelamento.ThrowIfCancellationRequested();
        try
        {
            Directory.CreateDirectory(pasta);
            await File.WriteAllTextAsync(Caminho(provedor, chave), JsonSerializer.Serialize(resultado), cancelamento);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* A consulta funciona mesmo sem cache. */ }
        return resultado;
    }

    internal static bool Compativel(string provedor, string id, JsonElement item)
    {
        if (provedor == "Claude (Anthropic)")
            return id.StartsWith("claude-", StringComparison.Ordinal) &&
                !(item.TryGetProperty("lifecycle", out var ciclo) && ciclo.GetString() == "retired");
        // Os catálogos não informam todas as capacidades de PDF/Responses.
        // Excluímos famílias especializadas; o teste de conexão continua disponível.
        string[] especializadas = ["audio", "realtime", "tts", "image", "embedding", "moderation", "transcribe", "codex", "search", "deep-research", "robotics"];
        if (especializadas.Any(parte => id.Contains(parte, StringComparison.OrdinalIgnoreCase))) return false;
        if (provedor == "Gemini (Google)")
            return id.StartsWith("gemini-", StringComparison.Ordinal) &&
                item.TryGetProperty("supportedGenerationMethods", out var metodos) &&
                metodos.EnumerateArray().Any(m => m.GetString() == "generateContent");
        return id.StartsWith("gpt-", StringComparison.Ordinal) ||
            (id.Length > 1 && id[0] == 'o' && char.IsDigit(id[1]));
    }
}
