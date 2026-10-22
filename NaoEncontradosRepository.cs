using System.Globalization;

namespace COTACAO_INSUMO;

public sealed record ItemNaoEncontradoSalvo(long Id, string Loja, int Mes, int Ano,
    string Fornecedor, string Insumo, decimal PrecoNormalizado, string TipoPreco,
    string UnidadeOriginal, DateTime DataRegistro);

public sealed class NaoEncontradosRepository
{
    private readonly string caminho;
    public NaoEncontradosRepository(string? caminhoBanco = null)
    {
        caminho = caminhoBanco ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "COTACAO_INSUMO", "configuracoes.db");
    }

    private ConfiguracaoIARepository.Banco Abrir()
    {
        var banco = new ConfiguracaoIARepository.Banco(caminho);
        try
        {
            banco.Executar("""
                CREATE TABLE IF NOT EXISTS insumos_nao_encontrados (
                    id INTEGER PRIMARY KEY,
                    loja TEXT NOT NULL, mes INTEGER NOT NULL CHECK(mes BETWEEN 1 AND 12),
                    ano INTEGER NOT NULL CHECK(ano BETWEEN 1900 AND 9999),
                    fornecedor TEXT NOT NULL, insumo TEXT NOT NULL,
                    preco_normalizado TEXT NOT NULL, tipo_preco TEXT NOT NULL,
                    unidade_original TEXT NOT NULL, data_registro TEXT NOT NULL,
                    UNIQUE(loja, mes, ano, fornecedor, insumo)
                );
                """);
            return banco;
        }
        catch { banco.Dispose(); throw; }
    }

    public int SalvarLote(string loja, int mes, int ano, InsumosCampos resultado, DateTime dataRegistro)
    {
        using var banco = Abrir();
        int gravados = 0;
        EmTransacao(banco, () =>
        {
            foreach (var item in resultado.Insumos.Where(item => !item.Encontrado))
            {
                using var comando = banco.Preparar("""
                    INSERT INTO insumos_nao_encontrados
                    (loja, mes, ano, fornecedor, insumo, preco_normalizado, tipo_preco, unidade_original, data_registro)
                    VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)
                    ON CONFLICT(loja, mes, ano, fornecedor, insumo) DO UPDATE SET
                        preco_normalizado = excluded.preco_normalizado,
                        tipo_preco = excluded.tipo_preco,
                        unidade_original = excluded.unidade_original,
                        data_registro = excluded.data_registro;
                    """);
                VincularPeriodo(comando, loja, mes, ano);
                comando.Vincular(4, resultado.Fornecedor);
                comando.Vincular(5, item.Nome);
                comando.Vincular(6, (item.PrecoUnitario ?? 0m).ToString(CultureInfo.InvariantCulture));
                comando.Vincular(7, item.Unidade switch
                {
                    "Kg" or "G" => "GRAMA",
                    "L" or "ML" => "MILILITRO",
                    "Unidade" or "Milheiro" => "UNIDADE",
                    _ => "NAO_IDENTIFICADO"
                });
                comando.Vincular(8, item.Unidade ?? "");
                comando.Vincular(9, dataRegistro.ToString("O", CultureInfo.InvariantCulture));
                comando.Ler();
                gravados++;
            }
        });
        return gravados;
    }

    public List<ItemNaoEncontradoSalvo> Consultar(string loja, int mes, int ano)
    {
        using var banco = Abrir();
        using var consulta = banco.Preparar("""
            SELECT id, loja, mes, ano, fornecedor, insumo, preco_normalizado,
                tipo_preco, unidade_original, data_registro
            FROM insumos_nao_encontrados WHERE loja = ?1 COLLATE NOCASE AND mes = ?2 AND ano = ?3
            ORDER BY id;
            """);
        VincularPeriodo(consulta, loja, mes, ano);
        var itens = new List<ItemNaoEncontradoSalvo>();
        while (consulta.Ler())
            itens.Add(new(long.Parse(consulta.Texto(0)), consulta.Texto(1),
                int.Parse(consulta.Texto(2)), int.Parse(consulta.Texto(3)),
                consulta.Texto(4), consulta.Texto(5),
                decimal.Parse(consulta.Texto(6), CultureInfo.InvariantCulture),
                consulta.Texto(7), consulta.Texto(8),
                DateTime.Parse(consulta.Texto(9), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)));
        return itens;
    }

    public void Restaurar(IEnumerable<ItemNaoEncontradoSalvo> itens, string? loja = null, int mes = 0, int ano = 0)
    {
        using var banco = Abrir();
        EmTransacao(banco, () =>
        {
            if (loja != null)
            {
                using var excluir = banco.Preparar("DELETE FROM insumos_nao_encontrados WHERE loja = ?1 COLLATE NOCASE AND mes = ?2 AND ano = ?3;");
                VincularPeriodo(excluir, loja, mes, ano); excluir.Ler();
            }
            foreach (var item in itens)
            {
                using var comando = banco.Preparar("""
                    INSERT INTO insumos_nao_encontrados
                    (id, loja, mes, ano, fornecedor, insumo, preco_normalizado, tipo_preco, unidade_original, data_registro)
                    VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10);
                    """);
                comando.Vincular(1, item.Id.ToString(CultureInfo.InvariantCulture));
                comando.Vincular(2, item.Loja); comando.Vincular(3, item.Mes.ToString()); comando.Vincular(4, item.Ano.ToString());
                comando.Vincular(5, item.Fornecedor); comando.Vincular(6, item.Insumo);
                comando.Vincular(7, item.PrecoNormalizado.ToString(CultureInfo.InvariantCulture));
                comando.Vincular(8, item.TipoPreco); comando.Vincular(9, item.UnidadeOriginal);
                comando.Vincular(10, item.DataRegistro.ToString("O", CultureInfo.InvariantCulture)); comando.Ler();
            }
        });
    }

    public void Excluir(IEnumerable<long> ids)
    {
        using var banco = Abrir();
        EmTransacao(banco, () =>
        {
            foreach (long id in ids.Distinct())
            {
                using var comando = banco.Preparar("DELETE FROM insumos_nao_encontrados WHERE id = ?1;");
                comando.Vincular(1, id.ToString(CultureInfo.InvariantCulture));
                comando.Ler();
            }
        });
    }

    public void ExcluirPeriodo(string loja, int mes, int ano)
    {
        using var banco = Abrir();
        using var comando = banco.Preparar("DELETE FROM insumos_nao_encontrados WHERE loja = ?1 COLLATE NOCASE AND mes = ?2 AND ano = ?3;");
        VincularPeriodo(comando, loja, mes, ano);
        comando.Ler();
    }

    private static void VincularPeriodo(ConfiguracaoIARepository.Consulta comando, string loja, int mes, int ano)
    {
        comando.Vincular(1, loja);
        comando.Vincular(2, mes.ToString(CultureInfo.InvariantCulture));
        comando.Vincular(3, ano.ToString(CultureInfo.InvariantCulture));
    }

    private static void EmTransacao(ConfiguracaoIARepository.Banco banco, Action acao)
    {
        banco.Executar("BEGIN IMMEDIATE;");
        try { acao(); banco.Executar("COMMIT;"); }
        catch { banco.Executar("ROLLBACK;"); throw; }
    }
}

