using System;
using System.IO;
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
                Convert.ToBase64String(
                    pdf
                );

            string prompt =
                PromptCotacao.Criar(
                    listaInsumosExcel
                );

            var corpo =
                new
                {
                    model =
                        modelo,

                    max_tokens =
                        12000,

                    messages =
                        new object[]
                        {
                            new
                            {
                                role =
                                    "user",

                                content =
                                    new object[]
                                    {
                                        new
                                        {
                                            type =
                                                "document",

                                            source =
                                                new
                                                {
                                                    type =
                                                        "base64",

                                                    media_type =
                                                        "application/pdf",

                                                    data =
                                                        base64
                                                }
                                        },

                                        new
                                        {
                                            type =
                                                "text",

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
                    "https://api.anthropic.com/v1/messages",
                    content
                );

            await ServicoIAUtil
                .ValidarRespostaHttpAsync(
                    response
                );

            string resposta =
                await response.Content
                    .ReadAsStringAsync();

            string texto =
                ExtrairTexto(
                    resposta
                );

            return ServicoIAUtil
                .ConverterResultado(
                    texto
                );
        }

        private string ExtrairTexto(
            string json)
        {
            using JsonDocument documento =
                JsonDocument.Parse(json);

            JsonElement raiz =
                documento.RootElement;

            if (
                !raiz.TryGetProperty(
                    "content",
                    out JsonElement content
                )
            )
            {
                throw new Exception(
                    "Claude não retornou conteúdo."
                );
            }

            foreach (
                JsonElement bloco
                in content.EnumerateArray()
            )
            {
                if (
                    bloco.TryGetProperty(
                        "type",
                        out JsonElement tipo
                    )
                    &&
                    tipo.GetString()
                        == "text"
                    &&
                    bloco.TryGetProperty(
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

            throw new Exception(
                "Não foi possível obter a resposta do Claude."
            );
        }
    }
}