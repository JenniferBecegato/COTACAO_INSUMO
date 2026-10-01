using System;
using System.IO;
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
                    contents =
                        new object[]
                        {
                            new
                            {
                                role =
                                    "user",

                                parts =
                                    new object[]
                                    {
                                        new
                                        {
                                            inline_data =
                                                new
                                                {
                                                    mime_type =
                                                        "application/pdf",

                                                    data =
                                                        base64
                                                }
                                        },

                                        new
                                        {
                                            text =
                                                prompt
                                        }
                                    }
                            }
                        },

                    generationConfig =
                        new
                        {
                            responseMimeType =
                                "application/json"
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
                    ObterUrl(),
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
                    "candidates",
                    out JsonElement candidates
                )
                ||
                candidates.GetArrayLength() == 0
            )
            {
                throw new Exception(
                    "O Gemini não retornou candidatos."
                );
            }

            JsonElement primeiro =
                candidates[0];

            if (
                !primeiro.TryGetProperty(
                    "content",
                    out JsonElement content
                )
                ||
                !content.TryGetProperty(
                    "parts",
                    out JsonElement parts
                )
            )
            {
                throw new Exception(
                    "O Gemini não retornou conteúdo."
                );
            }

            StringBuilder textoFinal =
                new StringBuilder();

            foreach (
                JsonElement parte
                in parts.EnumerateArray()
            )
            {
                if (
                    parte.TryGetProperty(
                        "text",
                        out JsonElement texto
                    )
                )
                {
                    textoFinal.Append(
                        texto.GetString()
                    );
                }
            }

            if (
                textoFinal.Length == 0
            )
            {
                throw new Exception(
                    "O Gemini retornou uma resposta vazia."
                );
            }

            return textoFinal.ToString();
        }
    }
}