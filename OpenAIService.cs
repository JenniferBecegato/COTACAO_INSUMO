using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public class OpenAIService :
        IAgenteIA
    {
        private readonly string apiKey;
        private readonly string modelo;
        private readonly int timeoutSegundos;

        public OpenAIService(
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

            client.DefaultRequestHeaders
                .Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    apiKey
                );

            return client;
        }

        // =====================================================
        // TESTAR CONEXÃO
        // =====================================================

        public async Task<bool>
            TestarConexaoAsync()
        {
            using HttpClient client =
                CriarCliente();

            var corpo =
                new
                {
                    model = modelo,

                    input =
                        "Responda somente com OK."
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
                    "https://api.openai.com/v1/responses",
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
            return await ExtracaoPdfIA.ExtrairAsync(client, "https://api.openai.com/v1/responses",
                "OpenAI", modelo, caminhoPdf);
        }
    }
}
