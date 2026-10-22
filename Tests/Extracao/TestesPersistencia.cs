using COTACAO_INSUMO;

internal static class TestesPersistencia
{
    public static void Executar(string pasta)
    {
        void Check(bool ok, string mensagem)
        {
            if (!ok) throw new Exception(mensagem);
        }
        InsumoCampo Item(string nome, decimal? preco, string? unidade, bool encontrado = false) => new()
        {
            Nome = nome, Preco = preco, Unidade = unidade, Encontrado = encontrado,
            PrecoUnitario = ConversorPrecoUnitario.Converter(preco, unidade)
        };
        InsumosCampos Lote(string fornecedor, params InsumoCampo[] itens) => new()
        {
            Fornecedor = fornecedor, Insumos = itens.ToList()
        };

        var repo = new NaoEncontradosRepository(Path.Combine(pasta, "persistencia.db"));
        var primeiraData = new DateTime(2026, 10, 15, 10, 0, 0, DateTimeKind.Local);
        var segundaData = primeiraData.AddHours(1);
        var lote = Lote("FORNECEDOR", Item("AMORA", 119.99m, "Kg"), Item("ENCONTRADO", 10m, "G", true));
        Check(repo.SalvarLote("Orlando", 10, 2026, lote, primeiraData) == 1, "Gravar apenas não encontrados.");
        var original = repo.Consultar("Orlando", 10, 2026).Single();
        Check(original.Insumo == "AMORA" && original.PrecoNormalizado == .12m && original.TipoPreco == "GRAMA" &&
            original.UnidadeOriginal == "Kg" && original.DataRegistro == primeiraData, "Mapeamento dos campos do banco.");
        repo.SalvarLote("Orlando", 10, 2026, lote, primeiraData);
        Check(repo.Consultar("Orlando", 10, 2026).Single().Id == original.Id, "Reprocessamento sem duplicar nem trocar id.");
        repo.SalvarLote("Orlando", 10, 2026, Lote("FORNECEDOR", Item("AMORA", 52m, "L")), segundaData);
        var atualizado = repo.Consultar("Orlando", 10, 2026).Single();
        Check(atualizado.Id == original.Id && atualizado.PrecoNormalizado == .052m && atualizado.TipoPreco == "MILILITRO" &&
            atualizado.UnidadeOriginal == "L" && atualizado.DataRegistro == segundaData, "Atualizar valores e data mantendo id.");
        repo.SalvarLote("Orlando", 10, 2026, Lote("FORNECEDOR", Item("AMORA", null, null)), segundaData);
        var incompleto = repo.Consultar("Orlando", 10, 2026).Single();
        Check(incompleto.PrecoNormalizado == 0 && incompleto.TipoPreco == "NAO_IDENTIFICADO" &&
            incompleto.UnidadeOriginal == "" && incompleto.Id == original.Id, "Nova leitura incompleta substitui dados antigos.");

        foreach (var (unidade, tipo, preco) in new[] {
            ("Kg", "GRAMA", .002m), ("G", "GRAMA", 2m), ("L", "MILILITRO", .002m),
            ("ML", "MILILITRO", 2m), ("Milheiro", "UNIDADE", .002m), ("Unidade", "UNIDADE", 2m) })
        {
            repo.SalvarLote("Drugstore", 10, 2026, Lote("FORNECEDOR", Item(unidade, 2m, unidade)), primeiraData);
            var salvo = repo.Consultar("Drugstore", 10, 2026).Single(x => x.Insumo == unidade);
            Check(salvo.TipoPreco == tipo && salvo.PrecoNormalizado == preco && salvo.UnidadeOriginal == unidade,
                "Mapear unidade " + unidade);
        }
        repo.SalvarLote("Orlando", 11, 2026, lote, primeiraData);
        repo.SalvarLote("Orlando", 10, 2027, lote, primeiraData);
        repo.SalvarLote("Orlando", 10, 2026, Lote("OUTRO", Item("AMORA", 3m, "G")), primeiraData);
        Check(repo.Consultar("Orlando", 11, 2026).Count == 1 && repo.Consultar("Orlando", 10, 2027).Count == 1 &&
            repo.Consultar("Orlando", 10, 2026).Count == 2, "Isolar período e fornecedor.");

        repo.SalvarLote("Orlando", 10, 2026,
            Lote("FORNECEDOR", Item("AMORA", 1m, "G"), Item("AMORA", 3m, "G")), segundaData);
        Check(repo.Consultar("Orlando", 10, 2026).Single(x => x.Id == original.Id).PrecoNormalizado == 3m,
            "Última ocorrência do lote prevalece.");
        repo.SalvarLote("Orlando", 10, 2026, Lote("FORNECEDOR", Item("AMORA", 4m, "G", true)), segundaData);
        Check(repo.Consultar("Orlando", 10, 2026).Single(x => x.Id == original.Id).PrecoNormalizado == 3m,
            "Insumo agora encontrado não exclui nem altera histórico.");

        bool falhou = false;
        try
        {
            // O último item viola NOT NULL depois de uma atualização e uma inserção válidas.
            repo.SalvarLote("Orlando", 10, 2026, Lote("FORNECEDOR",
                Item("AMORA", 99m, "G"), Item("NOVO", 10m, "G"), Item(null!, 1m, "G")), segundaData);
        }
        catch (IOException) { falhou = true; }
        var aposFalha = repo.Consultar("Orlando", 10, 2026);
        Check(falhou && aposFalha.Count == 2 && aposFalha.Single(x => x.Id == original.Id).PrecoNormalizado == 3m,
            "Rollback deve desfazer inserções e atualizações do PDF inteiro.");
        Console.WriteLine("PASS: persistência, upsert, id estável, dados ausentes, isolamento, última ocorrência e rollback.");
    }
}
