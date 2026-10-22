using ClosedXML.Excel;
using COTACAO_INSUMO;

internal static class TestesPreenchimento
{
    public static void Executar(string pasta)
    {
        void Check(bool ok, string mensagem)
        {
            if (!ok) throw new Exception(mensagem);
        }
        InsumoCampo Item(string nome, decimal? preco, bool encontrado = true) => new()
        {
            Nome = nome, Preco = preco, PrecoUnitario = preco, Unidade = "G", Encontrado = encontrado
        };
        InsumosCampos Lote(string fornecedor, params InsumoCampo[] itens) => new()
        {
            Fornecedor = fornecedor, Insumos = itens.ToList()
        };
        string Criar(string nome, Action<XLWorkbook>? modificar = null)
        {
            string caminho = Path.Combine(pasta, nome + ".xlsx");
            using var livro = new XLWorkbook();
            var aba = livro.AddWorksheet("Cotação");
            aba.Cell(4, 3).Value = "INSUMOS DROGA CAMPOS - DRUG 09/2026";
            aba.Cell(4, 4).Value = "Fornecedor mais em conta";
            aba.Cell(4, 5).Value = "QUANTIDADE A COMPRAR (MG OU ML)";
            aba.Cell(4, 6).Value = "GALENA";
            aba.Cell(4, 7).Value = "PURIFARMA";
            aba.Cell(5, 3).Value = "AMORA (Urgente)";
            aba.Cell(5, 4).Value = "GALENA";
            aba.Cell(5, 5).Value = 2000;
            aba.Cell(5, 6).Value = .3m;
            aba.Cell(5, 7).Value = .4m;
            aba.Cell(10, 2).FormulaA1 = "1+2";
            aba.Cell(6, 3).Value = "ACIDO CITRICO";
            aba.Cell(6, 6).Value = 0;
            aba.Cell(5, 6).Style.Fill.BackgroundColor = XLColor.Yellow;
            aba.Column(7).Width = 18;
            livro.AddWorksheet("Outra aba").Cell(1, 1).FormulaA1 = "1+2";
            modificar?.Invoke(livro);
            livro.SaveAs(caminho);
            return caminho;
        }
        var arquivosAntes = Directory.GetFiles(pasta).ToHashSet();
        string caminho = Criar("preenchimento");
        var resultado = PreenchedorCotacao.Preencher(caminho, Lote("Galéna S.A.", Item("AMORA", .1234m)));
        Check(resultado.PrecosPreenchidos == 1 && resultado.FornecedoresCriados == 0, "Fornecedor normalizado deve reutilizar coluna.");
        using (var livro = new XLWorkbook(caminho))
        {
            var aba = livro.Worksheet(1);
            Check(aba.Cell(5, 6).GetValue<decimal>() == .1234m && aba.Cell(5, 6).Style.NumberFormat.Format == "0.0000",
                "Preço deve preservar quatro casas ao reabrir.");
            Check(aba.Cell(5, 4).GetString() == "GALENA" && aba.Cell(5, 3).GetString() == "AMORA (Urgente)",
                "Atualizar menor fornecedor e preservar nome original.");
            Check(aba.Cell(5, 5).GetValue<int>() == 2000 && aba.Cell(5, 6).Style.Fill.BackgroundColor == XLColor.Yellow &&
                aba.Cell(10, 2).FormulaA1 == "1+2" && livro.Worksheet(2).Cell(1, 1).FormulaA1 == "1+2",
                "Preservar quantidades, estilo, outras abas e fórmulas.");
        }
        PreenchedorCotacao.Preencher(caminho, Lote("GALENA LTDA", Item("AMORA", .5m)));
        using (var livro = new XLWorkbook(caminho))
            Check(livro.Worksheet(1).Cell(5, 4).GetString() == "PURIFARMA", "Recalcular vencedor quando preço sobe.");
        PreenchedorCotacao.Preencher(caminho, Lote("GALENA", Item("AMORA", .4m)));
        using (var livro = new XLWorkbook(caminho))
            Check(livro.Worksheet(1).Cell(5, 4).GetString() == "GALENA", "Empate mantém primeira coluna.");
        byte[] antesInvalidos = File.ReadAllBytes(caminho);
        var invalidos = PreenchedorCotacao.Preencher(caminho,
            Lote("GALENA", Item("AMORA", null), Item("AMORA", 0), Item("ACIDO CITRICO", 9m, false)));
        Check(invalidos.PrecosPreenchidos == 0 && antesInvalidos.SequenceEqual(File.ReadAllBytes(caminho)),
            "Nulos, zeros e não encontrados não alteram arquivo.");
        var novo = PreenchedorCotacao.Preencher(caminho, Lote("NUTRIFARM", Item("ACIDO CITRICO", .0123m)));
        Check(novo.FornecedoresCriados == 1 && novo.PrecosPreenchidos == 1, "Criar fornecedor inexistente.");
        PreenchedorCotacao.Preencher(caminho, Lote("NUTRIFARM LIMITADA", Item("ACIDO CITRICO", .0234m)));
        using (var livro = new XLWorkbook(caminho))
        {
            var aba = livro.Worksheet(1);
            Check(aba.LastColumnUsed()!.ColumnNumber() == 8 && aba.Cell(6, 8).GetValue<decimal>() == .0234m,
                "Vários PDFs reutilizam coluna criada, inclusive com cabeçalho deslocado e metadados.");
            Check(aba.Cell(6, 4).GetString() == "NUTRIFARM" && aba.Column(8).Width == 18,
                "Ignorar preços zero e copiar largura da coluna vizinha.");
        }
        Check(Directory.GetFiles(pasta).Where(x => !arquivosAntes.Contains(x)).SequenceEqual(new[] { caminho }),
            "Nenhuma cópia de cotação ou arquivo temporário deve restar.");

        string similar = Criar("fornecedor-similar", livro => livro.Worksheet(1).Cell(4, 6).Value = "GALENAA");
        Check(PreenchedorCotacao.Preencher(similar, Lote("GALENA", Item("AMORA", 1m))).FornecedoresCriados == 0,
            "Match por similaridade deve evitar coluna duplicada.");
        string curto = Criar("fornecedor-curto", livro => livro.Worksheet(1).Cell(4, 6).Value = "PN");
        Check(PreenchedorCotacao.Preencher(curto, Lote("PN FARMACIA", Item("AMORA", 1m))).FornecedoresCriados == 1,
            "Sigla curta exige igualdade.");

        string ambiguo = Criar("fornecedor-ambiguo", livro =>
        {
            livro.Worksheet(1).Cell(4, 6).Value = "GALENA BRASIL";
            livro.Worksheet(1).Cell(4, 7).Value = "GALENA GLOBAL";
        });
        byte[] originalAmbiguo = File.ReadAllBytes(ambiguo);
        var semEscolha = PreenchedorCotacao.Preencher(ambiguo, Lote("GALENA", Item("AMORA", 1m)));
        Check(semEscolha.PrecosPreenchidos == 0 && semEscolha.Pendencias.Count == 1 &&
            originalAmbiguo.SequenceEqual(File.ReadAllBytes(ambiguo)), "Empate de fornecedor não altera nem cria coluna.");
        string exato = Criar("prioridade-exata", livro => livro.Worksheet(1).Cell(4, 7).Value = "GALENA BRASIL");
        Check(PreenchedorCotacao.Preencher(exato, Lote("GALENA", Item("AMORA", 1m))).PrecosPreenchidos == 1,
            "Igualdade do fornecedor prevalece sobre substring.");

        string linhas = Criar("linhas-ambiguas", livro =>
        {
            livro.Worksheet(1).Cell(5, 3).Value = "AMORA SECA";
            livro.Worksheet(1).Cell(6, 3).Value = "AMORA FLUIDA";
        });
        Check(PreenchedorCotacao.Preencher(linhas, Lote("GALENA", Item("AMORA", .2m))).PrecosPreenchidos == 0,
            "Empate de linhas não escolhe arbitrariamente.");
        Check(PreenchedorCotacao.Preencher(linhas, Lote("GALENA", Item("AMORA SECA", .2m))).PrecosPreenchidos == 1,
            "Igualdade de linha prevalece sobre similaridade.");
        string menorAusente = Criar("menor-ausente", livro => livro.Worksheet(1).Cell(4, 4).Clear());
        PreenchedorCotacao.Preencher(menorAusente, Lote("GALENA", Item("AMORA", .1m)));
        using (var livro = new XLWorkbook(menorAusente))
            Check(livro.Worksheet(1).Cell(4, 8).GetString() == "Fornecedor mais em conta" &&
                livro.Worksheet(1).Cell(5, 8).GetString() == "GALENA", "Criar coluna de menor fornecedor ausente.");
        string menorAmbiguo = Criar("menor-ambiguo", livro => livro.Worksheet(1).Cell(4, 7).Value = "Fornecedor mais em conta");
        Check(PreenchedorCotacao.Preencher(menorAmbiguo, Lote("GALENA", Item("AMORA", .1m))).PrecosPreenchidos == 0,
            "Duas colunas de menor fornecedor geram pendência.");
        foreach (string precoTexto in new[] { "0,0405", "0.0405" })
        {
            string comTexto = Criar("preco-texto-" + (precoTexto.Contains(',') ? "virgula" : "ponto"),
                livro => livro.Worksheet(1).Cell(5, 7).Value = precoTexto);
            PreenchedorCotacao.Preencher(comTexto, Lote("GALENA", Item("AMORA", .1m)));
            using var livro = new XLWorkbook(comTexto);
            Check(livro.Worksheet(1).Cell(5, 4).GetString() == "PURIFARMA",
                "Ler preços existentes em texto com separador decimal correto.");
        }

        byte[] antesBloqueio = File.ReadAllBytes(caminho);
        bool bloqueado = false;
        using (var trava = new FileStream(caminho, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            try { PreenchedorCotacao.Preencher(caminho, Lote("GALENA", Item("AMORA", 99m))); }
            catch (IOException) { bloqueado = true; }
        }
        Check(bloqueado && antesBloqueio.SequenceEqual(File.ReadAllBytes(caminho)), "Arquivo bloqueado deve permanecer intacto.");
        Check(!Directory.GetFiles(pasta).Any(x => x.Contains(".tmp.xlsx")), "Limpar temporários após falha.");
        string inexistente = Path.Combine(pasta, "nao-selecionada.xlsx");
        bool ausente = false;
        try { PreenchedorCotacao.Preencher(inexistente, Lote("GALENA", Item("AMORA", 1m))); }
        catch (FileNotFoundException) { ausente = true; }
        Check(ausente && !File.Exists(inexistente), "Não criar cotação quando o arquivo selecionado não existe.");
        string pastaPeriodo = Path.Combine(pasta, "2026-10");
        string cotacaoPeriodo = Path.Combine(pastaPeriodo, "Cotacao_Orlando_2026-10.xlsx");
        try
        {
            byte[] baseAntes = File.ReadAllBytes(caminho);
            PreparadorPlanilhaPeriodo.Preparar(caminho, cotacaoPeriodo);
            Check(baseAntes.SequenceEqual(File.ReadAllBytes(cotacaoPeriodo)), "Criar cotação do período a partir da base.");
            PreenchedorCotacao.Preencher(cotacaoPeriodo, Lote("GALENA", Item("AMORA", .9876m)));
            byte[] periodoPreenchido = File.ReadAllBytes(cotacaoPeriodo);
            PreparadorPlanilhaPeriodo.Preparar(caminho, cotacaoPeriodo);
            PreparadorPlanilhaPeriodo.Preparar("", cotacaoPeriodo);
            Check(periodoPreenchido.SequenceEqual(File.ReadAllBytes(cotacaoPeriodo)),
                "Cotação existente deve ser reutilizada sem sobrescrever, inclusive sem selecionar base.");
            Check(baseAntes.SequenceEqual(File.ReadAllBytes(caminho)), "Preencher o período sem alterar a planilha-base.");
        }
        finally
        {
            if (File.Exists(cotacaoPeriodo)) File.Delete(cotacaoPeriodo);
            if (Directory.Exists(pastaPeriodo)) Directory.Delete(pastaPeriodo);
        }
        bool semBase = false;
        try { PreparadorPlanilhaPeriodo.Preparar("", cotacaoPeriodo); }
        catch (FileNotFoundException) { semBase = true; }
        Check(semBase && !Directory.Exists(pastaPeriodo), "Exigir base apenas quando ainda não existir cotação do período.");
        Console.WriteLine("PASS: criação e reutilização da cotação no ano-mês, preservando a planilha-base.");
        Console.WriteLine("PASS: preenchimento no arquivo selecionado, fornecedores, ambiguidades, quatro casas, menor preço e arquivo bloqueado.");
    }
}
