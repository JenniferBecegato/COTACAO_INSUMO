using System;

namespace COTACAO_INSUMO
{
    public static class AgenteIAFactory
    {
        public static IAgenteIA Criar(
            string provedor,
            string apiKey,
            string modelo,
            int timeoutSegundos)
        {
            switch (provedor)
            {
                case "OpenAI":

                    return new OpenAIService(
                        apiKey,
                        modelo,
                        timeoutSegundos
                    );

                case "Claude (Anthropic)":

                    return new ClaudeService(
                        apiKey,
                        modelo,
                        timeoutSegundos
                    );

                case "Gemini (Google)":

                    return new GeminiService(
                        apiKey,
                        modelo,
                        timeoutSegundos
                    );

                default:

                    throw new Exception(
                        "Provedor de IA não reconhecido: "
                        + provedor
                    );
            }
        }
    }
}