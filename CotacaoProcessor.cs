using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace COTACAO_INSUMO
{
    public class CotacaoProcessor
    {
        private readonly IAgenteIA iaService;

        // Confiança mínima para preencher automaticamente
        private const decimal CONFIANCA_MINIMA = 0.90m;

        public CotacaoProcessor(
            IAgenteIA iaService)
        {
            this.iaService = iaService;
        }

        // =====================================================
        // PROCESSAMENTO PRINCIPAL
        // =====================================================

        public async Task<ResultadoProcessamento> ProcessarAsync(
            string caminhoExcel,
            List<string> pdfs)
        {
            ResultadoProcessamento resultado =
                new ResultadoProcessamento();

            using XLWorkbook workbook =
                new XLWorkbook(caminhoExcel);

            IXLWorksheet planilha =
                workbook.Worksheets.First();

            string listaInsumos =
                CriarListaInsumos(
                    planilha
                );

            foreach (string caminhoPdf in pdfs)
            {
                ResultadoCotacaoIA? resposta =
                    await iaService.AnalisarPdfAsync(
                        caminhoPdf,
                        listaInsumos
                    );

                if (resposta == null)
                    continue;

                string fornecedor =
                    resposta.Fornecedor?.Trim()
                    ?? "";

                if (
                    string.IsNullOrWhiteSpace(
                        fornecedor
                    )
                )
                {
                    fornecedor =
                        "FORNECEDOR NÃO IDENTIFICADO";
                }

                int colunaFornecedor =
                    ObterOuCriarColunaFornecedor(
                        planilha,
                        fornecedor
                    );

                foreach (
                    ItemCotacaoIA item
                    in resposta.Itens
                )
                {
                    ProcessarItem(
                        planilha,
                        colunaFornecedor,
                        fornecedor,
                        item,
                        resultado
                    );
                }
            }

            workbook.Save();

            return resultado;
        }

        // =====================================================
        // PROCESSAR ITEM
        // =====================================================

        private void ProcessarItem(
            IXLWorksheet planilha,
            int colunaFornecedor,
            string fornecedor,
            ItemCotacaoIA item,
            ResultadoProcessamento resultado)
        {
            // Cálculo feito pelo C#, não pela IA
            ResultadoPreco preco =
            CalculadoraPreco.Calcular(
                item.UnidadeOriginal,
                item.Quantidade,
                item.ValorTotal,
                item.ValorUnitario
            );

            item.PrecoNormalizado =
                preco.PrecoNormalizado;

            item.TipoPreco =
                preco.TipoPreco;

            // -------------------------------------------------
            // Unidade não reconhecida
            // -------------------------------------------------

            if (!preco.PodeConverter)
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // -------------------------------------------------
            // IA não encontrou correspondência
            // -------------------------------------------------

            if (!item.Encontrado)
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // -------------------------------------------------
            // Confiança baixa
            // -------------------------------------------------

            if (
                item.Confianca <
                CONFIANCA_MINIMA
            )
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // -------------------------------------------------
            // Nome vazio
            // -------------------------------------------------

            if (
                string.IsNullOrWhiteSpace(
                    item.InsumoExcel
                )
            )
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // -------------------------------------------------
            // Localizar linha no Excel
            // -------------------------------------------------

            int linha =
                LocalizarLinhaInsumo(
                    planilha,
                    item.InsumoExcel
                );

            if (linha <= 0)
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // -------------------------------------------------
            // PRESERVAR PRECISÃO DO PREÇO
            // -------------------------------------------------

            decimal valorArredondado =
                preco.PrecoNormalizado;

            IXLCell celula =
                planilha.Cell(
                    linha,
                    colunaFornecedor
                );

            celula.Value =
                valorArredondado;

            celula.Style
                .NumberFormat
                .Format =
                "0.000#########################";

            // -------------------------------------------------
            // ATUALIZAR FORNECEDOR MAIS BARATO
            // -------------------------------------------------

            AtualizarFornecedorMaisBarato(
                planilha,
                linha
            );

            resultado.TotalPreenchidos++;
        }

        // =====================================================
        // NÃO ENCONTRADOS
        // =====================================================

        private void AdicionarNaoEncontrado(
            ResultadoProcessamento resultado,
            string fornecedor,
            ItemCotacaoIA item)
        {
            decimal preco =
                item.PrecoNormalizado;

            resultado.NaoEncontrados.Add(
                new ItemNaoEncontradoProcessado
                {
                    Fornecedor =
                        fornecedor,

                    Insumo =
                        item.ProdutoPdf,

                    PrecoNormalizado =
                        preco,

                    TipoPreco =
                        item.TipoPreco,

                    UnidadeOriginal =
                        item.UnidadeOriginal
                }
            );
        }

        // =====================================================
        // CRIAR LISTA DOS INSUMOS DO EXCEL
        // =====================================================

        private string CriarListaInsumos(
            IXLWorksheet planilha)
        {
            StringBuilder sb =
                new StringBuilder();

            IXLRange? range =
                planilha.RangeUsed();

            if (range == null)
                return "";

            int ultimaLinha =
                range
                    .LastRow()
                    .RowNumber();

            for (
                int linha = 2;
                linha <= ultimaLinha;
                linha++)
            {
                string nome =
                    planilha
                        .Cell(
                            linha,
                            1
                        )
                        .GetFormattedString()
                        .Trim();

                if (
                    !string.IsNullOrWhiteSpace(
                        nome
                    )
                )
                {
                    sb.AppendLine(nome);
                }
            }

            return sb.ToString();
        }

        // =====================================================
        // LOCALIZAR INSUMO NO EXCEL
        // =====================================================

        private int LocalizarLinhaInsumo(
            IXLWorksheet planilha,
            string nome)
        {
            string procurar =
                NormalizarNome(nome);

            IXLRange? range =
                planilha.RangeUsed();

            if (range == null)
                return -1;

            int ultimaLinha =
                range
                    .LastRow()
                    .RowNumber();

            for (
                int linha = 2;
                linha <= ultimaLinha;
                linha++)
            {
                string atual =
                    planilha
                        .Cell(
                            linha,
                            1
                        )
                        .GetFormattedString();

                if (
                    NormalizarNome(atual)
                    ==
                    procurar
                )
                {
                    return linha;
                }
            }

            return -1;
        }

        // =====================================================
        // NORMALIZAR NOME DO INSUMO
        // =====================================================

        private string NormalizarNome(
            string texto)
        {
            if (
                string.IsNullOrWhiteSpace(
                    texto
                )
            )
            {
                return "";
            }

            string resultado =
                texto
                    .Trim()
                    .ToUpperInvariant()
                    .Replace("-", " ")
                    .Replace("/", " ");

            while (
                resultado.Contains("  ")
            )
            {
                resultado =
                    resultado.Replace(
                        "  ",
                        " "
                    );
            }

            return resultado;
        }

        // =====================================================
        // NORMALIZAR FORNECEDOR
        // =====================================================

        private string NormalizarFornecedor(
            string fornecedor)
        {
            if (
                string.IsNullOrWhiteSpace(
                    fornecedor
                )
            )
            {
                return "";
            }

            string nome =
                fornecedor
                    .Trim()
                    .ToUpperInvariant();

            // =================================================
            // APELIDOS / NOMES COMERCIAIS
            // =================================================

            if (nome.Contains("SIXTY"))
                return "SIXTY";

            if (nome.Contains("GALENA"))
                return "GALENA";

            if (nome.Contains("PURIFARMA"))
                return "PURIFARMA";

            if (nome.Contains("FAGRON"))
                return "FAGRON";

            if (nome.Contains("SOVITA"))
                return "SOVITA";

            if (nome.Contains("EMBRAFARMA"))
                return "EMBRAFARMA";

            if (nome.Contains("INFINITY"))
                return "INFINITY";

            if (nome.Contains("FLORIEN"))
                return "FLORIEN";

            if (
                nome.Contains("VALDEQUIMICA") ||
                nome.Contains("VALDEQUÍMICA")
            )
            {
                return "VALDEQUIMICA";
            }

            if (nome.Contains("AQIA"))
                return "AQIA";

            if (nome.Contains("NUTRIFARMA"))
                return "NUTRIFARMA";

            if (nome.Contains("CALDIC"))
                return "CALDIC";

            if (
                nome == "PN" ||
                nome.StartsWith("PN ")
            )
            {
                return "PN";
            }

            nome =
                nome
                    .Replace(".", "")
                    .Replace("-", " ")
                    .Replace("/", " ");

            while (
                nome.Contains("  ")
            )
            {
                nome =
                    nome.Replace(
                        "  ",
                        " "
                    );
            }

            return nome.Trim();
        }

        // =====================================================
        // OBTER OU CRIAR COLUNA DO FORNECEDOR
        // =====================================================

        private int ObterOuCriarColunaFornecedor(
            IXLWorksheet planilha,
            string fornecedor)
        {
            IXLRange? range =
                planilha.RangeUsed();

            int ultimaColuna =
                range?
                    .LastColumn()
                    .ColumnNumber()
                ?? 1;

            string fornecedorNormalizado =
                NormalizarFornecedor(
                    fornecedor
                );

            // =================================================
            // PRIMEIRO PROCURA UMA COLUNA JÁ EXISTENTE
            // =================================================

            for (
                int coluna = 1;
                coluna <= ultimaColuna;
                coluna++)
            {
                string cabecalho =
                    planilha
                        .Cell(
                            1,
                            coluna
                        )
                        .GetFormattedString()
                        .Trim();

                if (
                    string.IsNullOrWhiteSpace(
                        cabecalho
                    )
                )
                {
                    continue;
                }

                string cabecalhoNormalizado =
                    NormalizarFornecedor(
                        cabecalho
                    );

                if (
                    cabecalhoNormalizado ==
                    fornecedorNormalizado
                )
                {
                    return coluna;
                }
            }

            // =================================================
            // NÃO EXISTE:
            // CRIA NOVA COLUNA
            // =================================================

            int novaColuna =
                ultimaColuna + 1;

            IXLCell cabecalhoNovo =
                planilha.Cell(
                    1,
                    novaColuna
                );

            // Usa o nome normalizado se existir
            cabecalhoNovo.Value =
                string.IsNullOrWhiteSpace(
                    fornecedorNormalizado
                )
                    ? fornecedor
                    : fornecedorNormalizado;

            // Copia estilo da coluna anterior
            if (ultimaColuna > 0)
            {
                cabecalhoNovo.Style =
                    planilha
                        .Cell(
                            1,
                            ultimaColuna
                        )
                        .Style;
            }

            return novaColuna;
        }

        // =====================================================
        // ATUALIZAR FORNECEDOR MAIS BARATO
        // =====================================================

        private void AtualizarFornecedorMaisBarato(
            IXLWorksheet planilha,
            int linha)
        {
            IXLRange? range =
                planilha.RangeUsed();

            if (range == null)
                return;

            int ultimaColuna =
                range
                    .LastColumn()
                    .ColumnNumber();

            int colunaFornecedorMaisBarato = -1;
            int colunaQuantidade = -1;

            // =========================================================
            // LOCALIZAR COLUNAS FIXAS
            // =========================================================

            for (
                int coluna = 1;
                coluna <= ultimaColuna;
                coluna++
            )
            {
                string cabecalho =
                    NormalizarNome(
                        planilha
                            .Cell(1, coluna)
                            .GetFormattedString()
                    );

                if (
                    cabecalho.Contains(
                        "FORNECEDOR MAIS EM CONTA"
                    )
                )
                {
                    colunaFornecedorMaisBarato =
                        coluna;
                }

                if (
                    cabecalho.Contains(
                        "QUANTIDADE A COMPRAR"
                    )
                    ||
                    cabecalho == "QUANTIDADE"
                    ||
                    cabecalho == "QTD"
                )
                {
                    colunaQuantidade =
                        coluna;
                }
            }

            if (
                colunaFornecedorMaisBarato <= 0
            )
            {
                return;
            }

            // =========================================================
            // FORNECEDORES COMEÇAM DEPOIS DAS COLUNAS FIXAS
            // =========================================================

            int primeiraColunaFornecedor =
                Math.Max(
                    colunaFornecedorMaisBarato,
                    colunaQuantidade
                ) + 1;

            decimal? menorPreco = null;

            string fornecedorMaisBarato = "";

            // =========================================================
            // VERIFICAR SOMENTE FORNECEDORES
            // =========================================================

            for (
                int coluna = primeiraColunaFornecedor;
                coluna <= ultimaColuna;
                coluna++
            )
            {
                string cabecalho =
                    planilha
                        .Cell(1, coluna)
                        .GetFormattedString()
                        .Trim();

                if (
                    string.IsNullOrWhiteSpace(
                        cabecalho
                    )
                )
                {
                    continue;
                }

                IXLCell celula =
                    planilha.Cell(
                        linha,
                        coluna
                    );

                if (
                    celula.TryGetValue<decimal>(
                        out decimal preco
                    )
                )
                {
                    if (preco <= 0)
                        continue;

                    if (
                        menorPreco == null ||
                        preco < menorPreco.Value
                    )
                    {
                        menorPreco =
                            preco;

                        fornecedorMaisBarato =
                            cabecalho;
                    }
                }
            }

            IXLCell celulaMaisBarato =
                planilha.Cell(
                    linha,
                    colunaFornecedorMaisBarato
                );

            // =========================================================
            // NENHUM PREÇO PREENCHIDO
            // =========================================================

            if (
                menorPreco == null ||
                string.IsNullOrWhiteSpace(
                    fornecedorMaisBarato
                )
            )
            {
                celulaMaisBarato.Clear(
                    XLClearOptions.Contents
                );

                celulaMaisBarato
                    .Style
                    .Fill
                    .BackgroundColor =
                    XLColor.NoColor;

                return;
            }

            // =========================================================
            // ENCONTROU MENOR PREÇO
            // =========================================================

            celulaMaisBarato.Value =
                fornecedorMaisBarato;

            celulaMaisBarato
                .Style
                .Fill
                .BackgroundColor =
                XLColor.LightGreen;
        }
    }
}

