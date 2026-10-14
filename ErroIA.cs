using System.Net;
using System.Text.Json;

namespace COTACAO_INSUMO;

public sealed class RespostaIAException(string mensagem) : Exception(mensagem);

public sealed class ErroIAException : Exception
{
    public HttpStatusCode StatusHttp { get; }
    public string Codigo { get; }
    public string RespostaOriginal { get; }

    public ErroIAException(HttpStatusCode status, string resposta)
        : base($"HTTP {(int)status}\n{resposta}")
    {
        StatusHttp = status;
        RespostaOriginal = resposta;
        Codigo = "";
        try
        {
            using var doc = JsonDocument.Parse(resposta);
            var erro = doc.RootElement;
            if (erro.ValueKind == JsonValueKind.Object && erro.TryGetProperty("error", out var interno)) erro = interno;
            if (erro.ValueKind == JsonValueKind.Object)
            {
                Codigo = string.Join(" ", new[] { "code", "status", "type" }
                    .Where(nome => erro.TryGetProperty(nome, out _))
                    .Select(nome => erro.GetProperty(nome).ToString()));
            }
        }
        catch (JsonException) { }
    }
}

public sealed record MensagemErroIA(string Titulo, string Descricao, string Orientacao, int? StatusHttp)
{
    public static MensagemErroIA Classificar(Exception ex)
    {
        var http = ex as ErroIAException;
        int? status = http != null ? (int)http.StatusHttp : (ex as HttpRequestException)?.StatusCode is { } s ? (int)s : null;
        string texto = ((http?.Codigo ?? "") + " " + ex.Message).ToLowerInvariant();
        bool Tem(params string[] termos) => termos.Any(texto.Contains);
        MensagemErroIA Criar(string titulo, string descricao, string orientacao) => new(titulo, descricao, orientacao, status);

        string codigo = http?.Codigo.ToLowerInvariant() ?? "";
        bool CodigoTem(params string[] termos) => termos.Any(codigo.Contains);
        int? categoriaHttp = CodigoTem("api_key_invalid", "invalid_api_key", "authentication_error") ? 401
            : CodigoTem("permission_denied", "permission_error") ? 403
            : CodigoTem("model_not_found", "not_found_error") ? 404
            : CodigoTem("rate_limit", "resource_exhausted") ? 429
            : CodigoTem("overloaded", "unavailable") ? 503
            : status;

        if (Tem("insufficient_quota", "credit_balance_exhausted", "credit balance", "billing", "quota exceeded", "quota_exceeded"))
            return Criar("Créditos ou cota esgotados", "A conta da API atingiu sua cota ou está sem créditos disponíveis.", "Verifique o faturamento e os limites da conta no provedor. Adicione créditos ou aguarde a renovação da cota antes de tentar novamente.");
        if (Tem("api_key_invalid", "invalid_api_key", "invalid api key", "authentication_error") || categoriaHttp == 401)
            return Criar("Chave da API inválida", "O provedor não aceitou a autenticação da API.", "Confira a chave em Configurações. Se necessário, gere uma nova chave no provedor e teste a conexão.");
        if (Tem("permission_denied", "permission_error") || categoriaHttp == 403)
            return Criar("Acesso não permitido", "A conta ou chave da API não tem permissão para esta solicitação.", "Confira as permissões da chave e o acesso ao modelo selecionado nas configurações do provedor.");
        if (Tem("model_not_found", "not_found_error") || categoriaHttp == 404)
            return Criar("Modelo ou recurso indisponível", "O modelo ou recurso solicitado não foi encontrado ou não está disponível para esta conta.", "Confira o modelo em Configurações e selecione um modelo disponível para sua chave.");
        if (Tem("rate_limit", "rate limit", "resource_exhausted") || categoriaHttp == 429)
            return Criar("Limite de requisições atingido", "O provedor limitou temporariamente as solicitações à API.", "Aguarde alguns instantes e tente novamente. Se persistir, confira os limites de uso no provedor.");
        if (Tem("overloaded", "high demand", "unavailable") || categoriaHttp is 500 or 502 or 503 or 529)
            return Criar("Serviço de IA temporariamente indisponível", "O serviço está indisponível ou enfrentando alta demanda. Esses picos costumam ser temporários.", "Aguarde alguns instantes e tente novamente. Se persistir, selecione outro modelo ou provedor nas configurações.");
        if (ex is TaskCanceledException or TimeoutException || categoriaHttp is 408 or 504 || Tem("timeout", "timed out"))
            return Criar("Tempo de resposta excedido", "O serviço de IA não respondeu dentro do tempo disponível.", "Tente novamente. Se persistir, aumente o tempo limite em Configurações ou processe menos PDFs por vez.");
        if (ex is HttpRequestException)
            return Criar("Falha na comunicação com a IA", "Não foi possível concluir a comunicação com o provedor.", "Verifique sua conexão com a internet e o acesso ao provedor e tente novamente.");
        if (ex is JsonException or RespostaIAException || Tem("não retornou", "resposta vazia", "obter a resposta"))
            return Criar("Resposta inválida da IA", "O provedor não retornou conteúdo no formato esperado para a cotação.", "Tente novamente. Se persistir, selecione outro modelo e confira o conteúdo dos PDFs.");
        if (categoriaHttp is 400 or 413 or 422)
            return Criar("Solicitação não aceita pela IA", "O provedor não conseguiu aceitar os dados ou parâmetros enviados.", "Confira o modelo e os arquivos enviados. Tente com menos PDFs ou arquivos menores e consulte os detalhes técnicos.");
        return Criar("Não foi possível concluir a operação", "Ocorreu um erro que não pôde ser identificado automaticamente.", "Consulte os detalhes técnicos. Confira as configurações e tente novamente; se persistir, copie os detalhes para solicitar suporte.");
    }
}
