using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public static class ServicoIAUtil
    {
        public static string LimparJson(
            string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return "{}";

            texto = texto.Trim();

            if (
                texto.StartsWith(
                    "```json",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                texto =
                    texto.Substring(7);
            }
            else if (
                texto.StartsWith("```")
            )
            {
                texto =
                    texto.Substring(3);
            }

            if (
                texto.EndsWith("```")
            )
            {
                texto =
                    texto.Substring(
                        0,
                        texto.Length - 3
                    );
            }

            return texto.Trim();
        }

        public static ResultadoCotacaoIA?
            ConverterResultado(
                string json)
        {
            json =
                LimparJson(json);

            return
                JsonSerializer.Deserialize<
                    ResultadoCotacaoIA>(
                    json,
                    new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive =
                            true
                    }
                );
        }

        public static async Task
            ValidarRespostaHttpAsync(
                HttpResponseMessage response)
        {
            if (response.IsSuccessStatusCode)
                return;

            string erro =
                await response.Content
                    .ReadAsStringAsync();

            throw new Exception(
                $"Erro HTTP {(int)response.StatusCode} " +
                $"({response.StatusCode})\n\n{erro}"
            );
        }
    }
}