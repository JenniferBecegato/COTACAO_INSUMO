using System.Net.Http.Json;
using System.Text.Json;

namespace COTACAO_INSUMO;

internal static class ExtracaoPdfIA
{
    public static async Task<ExtracaoInsumos> ExtrairAsync(HttpClient client, string url,
        string provedor, string modelo, string caminho)
    {
        var base64 = Convert.ToBase64String(await File.ReadAllBytesAsync(caminho));
        object corpo = provedor switch
        {
            "OpenAI" => new
            {
                model = modelo,
                input = new[] { new { role = "user", content = new object[] {
                    new { type = "input_file", filename = Path.GetFileName(caminho), file_data = "data:application/pdf;base64," + base64 },
                    new { type = "input_text", text = LeituraInsumos.Prompt }
                } } }
            },
            "Gemini" => new
            {
                contents = new[] { new { role = "user", parts = new object[] {
                    new { inline_data = new { mime_type = "application/pdf", data = base64 } },
                    new { text = LeituraInsumos.Prompt }
                } } },
                generationConfig = new { responseMimeType = "application/json" }
            },
            "Claude" => new
            {
                model = modelo, max_tokens = 32000,
                messages = new[] { new { role = "user", content = new object[] {
                    new { type = "document", source = new { type = "base64", media_type = "application/pdf", data = base64 } },
                    new { type = "text", text = LeituraInsumos.Prompt }
                } } }
            },
            _ => throw new ArgumentException("Provedor não reconhecido.")
        };
        using var response = await client.PostAsJsonAsync(url, corpo);
        await ServicoIAUtil.ValidarRespostaHttpAsync(response);
        using var documento = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        return LeituraInsumos.InterpretarRespostaPdf(ExtrairTexto(documento.RootElement, provedor));
    }

    internal static string ExtrairTexto(JsonElement raiz, string provedor)
    {
        if (provedor == "OpenAI")
        {
            if (!raiz.TryGetProperty("status", out var status) || status.GetString() != "completed")
                throw new RespostaIAException("A leitura do PDF não foi concluída pela OpenAI.");
            var partes = raiz.GetProperty("output").EnumerateArray()
                .Where(x => x.GetProperty("type").GetString() == "message")
                .SelectMany(x => x.GetProperty("content").EnumerateArray());
            var textos = new List<string>();
            foreach (var parte in partes)
            {
                if (parte.GetProperty("type").GetString() != "output_text")
                    throw new RespostaIAException("A IA não retornou a extração solicitada.");
                textos.Add(parte.GetProperty("text").GetString() ?? "");
            }
            return string.Concat(textos);
        }
        if (provedor == "Gemini")
        {
            var candidato = raiz.GetProperty("candidates")[0];
            if (candidato.GetProperty("finishReason").GetString() != "STOP")
                throw new RespostaIAException("O Gemini interrompeu a leitura; o resultado não será salvo como completo.");
            return string.Concat(candidato.GetProperty("content").GetProperty("parts").EnumerateArray()
                .Where(x => !x.TryGetProperty("thought", out var thought) || !thought.GetBoolean())
                .Where(x => x.TryGetProperty("text", out _)).Select(x => x.GetProperty("text").GetString()));
        }
        if (raiz.GetProperty("stop_reason").GetString() != "end_turn")
            throw new RespostaIAException("O Claude interrompeu a leitura; o resultado não será salvo como completo.");
        return string.Concat(raiz.GetProperty("content").EnumerateArray()
            .Where(x => x.GetProperty("type").GetString() == "text")
            .Select(x => x.GetProperty("text").GetString()));
    }
}
