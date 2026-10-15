using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public static class ServicoIAUtil
    {
        public static async Task
            ValidarRespostaHttpAsync(
                HttpResponseMessage response)
        {
            if (response.IsSuccessStatusCode)
                return;

            string erro =
                await response.Content
                    .ReadAsStringAsync();

            throw new ErroIAException(response.StatusCode, erro);
        }
    }
}
