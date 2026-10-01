using System;
using System.IO;
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

        // =====================================================
        // ANALISAR PDF
        // =====================================================

        public async Task<ResultadoCotacaoIA?>
            AnalisarPdfAsync(
                string caminhoPdf,
                string listaInsumosExcel)
        {
            if (!File.Exists(caminhoPdf))
            {
                throw new FileNotFoundException(
                    "PDF não encontrado.",
                    caminhoPdf
                );
            }

            byte[] pdf =
                await File.ReadAllBytesAsync(
                    caminhoPdf
                );

            string base64 =
                Convert.ToBase64String(pdf);

            string prompt =
                PromptCotacao.Criar(
                    listaInsumosExcel
                );

            var corpo =
                new
                {
                    model = modelo,

                    input = new object[]
                    {
                        new
                        {
                            role = "user",

                            content = new object[]
                            {
                                new
                                {
                                    type =
                                        "input_file",

                                    filename =
                                        Path.GetFileName(
                                            caminhoPdf
                                        ),

                                    file_data =
                                        "data:application/pdf;base64,"
                                        + base64
                                },

                                new
                                {
                                    type =
                                        "input_text",

                                    text =
                                        prompt
                                }
                            }
                        }
                    }
                };

            string requestJson =
                JsonSerializer.Serialize(
                    corpo
                );

            using HttpClient client =
                CriarCliente();

            using StringContent content =
                new StringContent(
                    requestJson,
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

            string respostaJson =
                await response.Content
                    .ReadAsStringAsync();

            string texto =
                ExtrairTexto(
                    respostaJson
                );

            return ServicoIAUtil
                .ConverterResultado(
                    texto
                );
        }

        // =====================================================
        // EXTRAIR TEXTO DA RESPOSTA
        // =====================================================

        private string ExtrairTexto(
            string json)
        {
            using JsonDocument documento =
                JsonDocument.Parse(json);

            JsonElement raiz =
                documento.RootElement;

            if (
                !raiz.TryGetProperty(
                    "output",
                    out JsonElement output
                )
            )
            {
                throw new Exception(
                    "A OpenAI não retornou o campo 'output'."
                );
            }

            foreach (
                JsonElement item
                in output.EnumerateArray()
            )
            {
                if (
                    !item.TryGetProperty(
                        "content",
                        out JsonElement contents
                    )
                )
                {
                    continue;
                }

                foreach (
                    JsonElement content
                    in contents.EnumerateArray()
                )
                {
                    if (
                        content.TryGetProperty(
                            "type",
                            out JsonElement tipo
                        )
                        &&
                        tipo.GetString()
                            == "output_text"
                        &&
                        content.TryGetProperty(
                            "text",
                            out JsonElement texto
                        )
                    )
                    {
                        return
                            texto.GetString()
                            ?? "";
                    }
                }
            }

            throw new Exception(
                "Não foi possível obter o texto retornado pela OpenAI."
            );
        }
    }
}