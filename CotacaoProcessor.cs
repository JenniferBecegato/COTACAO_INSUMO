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

        // Confiança mínima para preenchimento automático
        private const decimal CONFIANCA_MINIMA =
            0.90m;

        public CotacaoProcessor(
    IAgenteIA iaService)
        {
            this.iaService =
                iaService;
        }

        public async Task<ResultadoProcessamento>
            ProcessarAsync(
                string caminhoExcel,
                List<string> pdfs)
        {
            ResultadoProcessamento resultado =
                new ResultadoProcessamento();

            using XLWorkbook workbook =
                new XLWorkbook(
                    caminhoExcel
                );

            IXLWorksheet planilha =
                workbook.Worksheets.First();

            string listaInsumos =
                CriarListaInsumos(
                    planilha
                );

            foreach (
                string caminhoPdf
                in pdfs)
            {
                ResultadoCotacaoIA? resposta =
                    await iaService
                        .AnalisarPdfAsync(
                            caminhoPdf,
                            listaInsumos
                        );

                if (resposta == null)
                {
                    continue;
                }

                string fornecedor =
                    resposta.Fornecedor.Trim();

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
                    in resposta.Itens)
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
        // ITEM
        // =====================================================

        private void ProcessarItem(
            IXLWorksheet planilha,
            int colunaFornecedor,
            string fornecedor,
            ItemCotacaoIA item,
            ResultadoProcessamento resultado)
        {
            ResultadoPreco preco =
                CalculadoraPreco.Calcular(
                    item.UnidadeOriginal,
                    item.ValorUnitario
                );

            item.PrecoNormalizado =
                preco.PrecoNormalizado;

            item.TipoPreco =
                preco.TipoPreco;

            // Unidade não reconhecida
            if (!preco.PodeConverter)
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // IA não encontrou
            if (!item.Encontrado)
            {
                AdicionarNaoEncontrado(
                    resultado,
                    fornecedor,
                    item
                );

                return;
            }

            // Confiança insuficiente
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

            // =================================================
            // ESCREVER PREÇO NORMALIZADO
            // =================================================

            IXLCell celula =
                planilha.Cell(
                    linha,
                    colunaFornecedor
                );

            decimal valorArredondado =
    Math.Round(
        preco.PrecoNormalizado,
        5,
        MidpointRounding.AwayFromZero
    );

            celula.Value =
                valorArredondado;

            celula.Style.NumberFormat.Format =
                "0.00000";

            AtualizarFornecedorMaisBarato(
    planilha,
    linha
);

            resultado.TotalPreenchidos++;
        }

        // =====================================================
        // NÃO ENCONTRADO
        // =====================================================

        private void AdicionarNaoEncontrado(
            ResultadoProcessamento resultado,
            string fornecedor,
            ItemCotacaoIA item)
        {
            resultado.NaoEncontrados.Add(
                new ItemNaoEncontradoProcessado
                {
                    Fornecedor =
                        fornecedor,

                    Insumo =
                        item.ProdutoPdf,

                    PrecoNormalizado =
                        item.PrecoNormalizado,

                    TipoPreco =
                        item.TipoPreco,

                    UnidadeOriginal =
                        item.UnidadeOriginal
                }
            );
        }

        // =====================================================
        // LISTA DE INSUMOS
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
                    sb.AppendLine(
                        nome
                    );
                }
            }

            return sb.ToString();
        }

        // =====================================================
        // LOCALIZAR INSUMO
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
                    NormalizarNome(atual) ==
                    procurar
                )
                {
                    return linha;
                }
            }

            return -1;
        }

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

            return texto
                .Trim()
                .ToUpperInvariant()
                .Replace("-", " ")
                .Replace("/", " ")
                .Replace("  ", " ");
        }

        // =====================================================
        // FORNECEDOR
        // =====================================================

        private string NormalizarFornecedor(
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

            // Não encontrou fornecedor equivalente:
            // cria nova coluna
            int novaColuna =
                ultimaColuna + 1;

            IXLCell cabecalhoNovo =
                planilha.Cell(
                    1,
                    novaColuna
                );

            cabecalhoNovo.Value =
                fornecedor;

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

            // aliases conhecidos
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

            if (nome.Contains("VALDEQUIMICA"))
                return "VALDEQUIMICA";

            if (nome.Contains("AQIA"))
                return "AQIA";

            if (nome.Contains("NUTRIFARMA"))
                return "NUTRIFARMA";

            if (nome.Contains("CALDIC"))
                return "CALDIC";

            if (nome == "PN")
                return "PN";

            return nome
                .Replace(".", "")
                .Replace("-", " ")
                .Replace("/", " ")
                .Replace("  ", " ");
        }

        private void ProcessarItem(
    IXLWorksheet planilha,
    int linha)
        {
            // Procura a coluna "Fornecedor mais em conta"
            int colunaFornecedorMaisBarato =
                -1;

            IXLRange? range =
                planilha.RangeUsed();

            if (range == null)
                return;

            int ultimaColuna =
                range
                    .LastColumn()
                    .ColumnNumber();

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
                    cabecalho.Equals(
                        "Fornecedor mais em conta",
                        StringComparison.OrdinalIgnoreCase
                    )
                )
                {
                    colunaFornecedorMaisBarato =
                        coluna;

                    break;
                }
            }

            if (
                colunaFornecedorMaisBarato <= 0
            )
            {
                return;
            }

            decimal? menorPreco =
                null;

            string fornecedorMenorPreco =
                "";

            for (
                int coluna = 1;
                coluna <= ultimaColuna;
                coluna++)
            {
                if (
                    coluna ==
                    colunaFornecedorMaisBarato
                )
                {
                    continue;
                }

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

                // evita colunas que não são fornecedores
                if (
                    cabecalho.Equals(
                        "INSUMO",
                        StringComparison.OrdinalIgnoreCase
                    )
                    ||
                    cabecalho.Contains(
                        "fornecedor",
                        StringComparison.OrdinalIgnoreCase
                    )
                    ||
                    cabecalho.Contains(
                        "quantidade",
                        StringComparison.OrdinalIgnoreCase
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
                    &&
                    preco > 0
                )
                {
                    if (
                        menorPreco == null
                        ||
                        preco < menorPreco
                    )
                    {
                        menorPreco =
                            preco;

                        fornecedorMenorPreco =
                            cabecalho;
                    }
                }
            }

            if (
                !string.IsNullOrWhiteSpace(
                    fornecedorMenorPreco
                )
            )
            {
                planilha
                    .Cell(
                        linha,
                        colunaFornecedorMaisBarato
                    )
                    .Value =
                    fornecedorMenorPreco;
            }
        }
    }
}