using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public class ClaudeService :
        IAgenteIA
    {
        private readonly string apiKey;
        private readonly string modelo;
        private readonly int timeoutSegundos;

        public ClaudeService(
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
                "x-api-key",
                apiKey
            );

            client.DefaultRequestHeaders.Add(
                "anthropic-version",
                "2023-06-01"
            );

            return client;
        }

        // =====================================================
        // TESTAR
        // =====================================================

        public async Task<bool>
            TestarConexaoAsync()
        {
            using HttpClient client =
                CriarCliente();

            var corpo =
                new
                {
                    model =
                        modelo,

                    max_tokens =
                        30,

                    messages =
                        new object[]
                        {
                            new
                            {
                                role =
                                    "user",

                                content =
                                    "Responda somente com OK."
                            }
                        }
                };

            string json =
                JsonSerializer.Serialize(
                    corpo
                );

            using StringContent content =
                new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json"
                );

            HttpResponseMessage response =
                await client.PostAsync(
                    "https://api.anthropic.com/v1/messages",
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
            return await ExtracaoPdfIA.ExtrairAsync(client, "https://api.anthropic.com/v1/messages",
                "Claude", modelo, caminhoPdf);
        }
    }
}
