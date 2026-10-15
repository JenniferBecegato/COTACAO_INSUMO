using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public class GeminiService :
        IAgenteIA
    {
        private readonly string apiKey;
        private readonly string modelo;
        private readonly int timeoutSegundos;

        public GeminiService(
            string apiKey,
            string modelo,
            int timeoutSegundos = 3600)
        {
            this.apiKey =
                apiKey;

            this.modelo =
                modelo;

            this.timeoutSegundos =
                timeoutSegundos;
        }

        private HttpClient CriarCliente()
        {
            HttpClient client =
                new HttpClient();

            client.Timeout =
                TimeSpan.FromSeconds(
                    timeoutSegundos
                );

            client.DefaultRequestHeaders.Add(
                "x-goog-api-key",
                apiKey
            );

            return client;
        }

        private string ObterUrl()
        {
            return
                "https://generativelanguage.googleapis.com/" +
                "v1beta/models/" +
                Uri.EscapeDataString(modelo) +
                ":generateContent";
        }

        // =====================================================
        // TESTAR
        // =====================================================

        public async Task<bool>
            TestarConexaoAsync()
        {
            var corpo =
                new
                {
                    contents =
                        new object[]
                        {
                            new
                            {
                                parts =
                                    new object[]
                                    {
                                        new
                                        {
                                            text =
                                                "Responda somente com OK."
                                        }
                                    }
                            }
                        }
                };

            string json =
                JsonSerializer.Serialize(
                    corpo
                );

            using HttpClient client =
                CriarCliente();

            using StringContent content =
                new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json"
                );

            HttpResponseMessage response =
                await client.PostAsync(
                    ObterUrl(),
                    content
                );

            await ServicoIAUtil
                .ValidarRespostaHttpAsync(
                    response
                );

            return true;
        }
        public async Task<ExtracaoInsumos> ExtrairInsumosAsync(string caminhoPdf)
        {
            using var client = CriarCliente();
            return await ExtracaoPdfIA.ExtrairAsync(client, ObterUrl(),
                "Gemini", modelo, caminhoPdf);
        }
    }
}
