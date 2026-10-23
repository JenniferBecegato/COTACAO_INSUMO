using ClosedXML.Excel;
using OpenAI.Responses;

#pragma warning disable OPENAI001

namespace COTACAO_INSUMO
{
    public partial class Form1 : Form
    {
        private sealed record AcaoHistorico(Action Desfazer, Action Refazer);
        private readonly List<AcaoHistorico> acoesDesfazer = new();
        private readonly List<AcaoHistorico> acoesRefazer = new();
        private HistoricoTabela? historicoTabelaAtivo;
        private Action? indicarAlteracaoTabela;
        private bool desfazendo;
        private void RegistrarDesfazer(Action desfazer, Action refazer)
        {
            if (desfazendo) return;
            acoesRefazer.Clear();
            acoesDesfazer.Add(new(desfazer, refazer));
            if (acoesDesfazer.Count > 50) acoesDesfazer.RemoveAt(0);
        }
        private void DesfazerUltimaAcao() => ExecutarHistorico(false);
        private void RefazerUltimaAcao() => ExecutarHistorico(true);
        private void ExecutarHistorico(bool refazer)
        {
            historicoTabelaAtivo?.Confirmar();
            var origem = refazer ? acoesRefazer : acoesDesfazer;
            var destino = refazer ? acoesDesfazer : acoesRefazer;
            if (origem.Count == 0) return;
            try
            {
                desfazendo = true;
                var acao = origem[^1];
                if (refazer) acao.Refazer(); else acao.Desfazer();
                origem.RemoveAt(origem.Count - 1);
                destino.Add(acao);
            }
            catch (Exception ex) { MessageBox.Show("Não foi possível " + (refazer ? "refazer" : "desfazer") + ".\n\n" + ex.Message,
                "Histórico", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { desfazendo = false; }
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.C) || keyData == (Keys.Control | Keys.X) ||
                keyData == (Keys.Control | Keys.V))
            {
                var tabela = painelConteudo.Controls.Find("tabelaCotacao", true)
                    .OfType<DataGridView>().FirstOrDefault();
                if (tabela?.ContainsFocus == true && !tabela.IsCurrentCellInEditMode)
                {
                    ExecutarAreaTransferencia(tabela, keyData & Keys.KeyCode);
                    return true;
                }
            }
            if (keyData == (Keys.Control | Keys.L))
            {
                var tabela = painelConteudo.Controls.Find("tabelaCotacao", true)
                    .OfType<DataGridView>().FirstOrDefault();
                if (tabela != null && tabela.Visible)
                {
                    LocalizarNaCotacao.Abrir(this, tabela);
                    return true;
                }
            }
            if (keyData == (Keys.Control | Keys.Z))
            {
                if (historicoTabelaAtivo != null)
                {
                    // Cancel the in-progress edit first; committed edits use the application history.
                    var editor = ActiveControl;
                    while (editor is ContainerControl container && container.ActiveControl != null) editor = container.ActiveControl;
                    if (editor is DataGridView grid && grid.IsCurrentCellInEditMode) { grid.CancelEdit(); return true; }
                }
                Control? controle = ActiveControl;
                while (controle is ContainerControl container && container.ActiveControl != null) controle = container.ActiveControl;
                if (controle is TextBoxBase texto && texto.CanUndo) { texto.Undo(); return true; }
                DesfazerUltimaAcao(); return true;
            }
            if (keyData == (Keys.Control | Keys.Y))
            {
                RefazerUltimaAcao(); return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ExecutarAreaTransferencia(DataGridView tabela, Keys tecla)
        {
            try
            {
                if (!tabela.EndEdit()) return;
                historicoTabelaAtivo?.Confirmar();
                if (tecla == Keys.V) AreaTransferenciaCelulas.Colar(tabela);
                else AreaTransferenciaCelulas.Copiar(tabela, tecla == Keys.X);
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                MessageBox.Show("A área de transferência está ocupada. Tente novamente.",
                    "Área de transferência", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { historicoTabelaAtivo?.Confirmar(); }
        }

        // =========================================================
        // CONFIGURAÇÃO DA IA
        // =========================================================

        private string provedorIA =
            "OpenAI";

        private string modeloIA =
            "gpt-5.6-sol";

        private string apiKeyIA =
            "";

        private int timeoutIA = 3600;
        private bool gravarJsonIA;
        private decimal percentualMinimoMatch = ComparadorInsumos.PercentualMinimoMatch;

        private readonly ConfiguracaoIARepository configuracoesIA = new();

        private bool configuracaoIASalva =
            false;



        private async Task<bool>
        TestarConexaoOpenAIAsync(
        string apiKey,
        string modelo)
        {
            ResponsesClient client =
                new ResponsesClient(
                    apiKey
                );

            ResponseResult resposta =
                await client.CreateResponseAsync(
                    modelo,
                    "Responda somente com OK."
                );

            string texto =
                resposta.GetOutputText();

            return
                !string.IsNullOrWhiteSpace(
                    texto
                );
        }

        private enum TelaAtual
        {
            Inicio,
            NovaCotacao,
            Consultar,
            NaoEncontrados,
            Configuracoes,
            Ajuda,
            Concluido
        }

        private TelaAtual telaAtual = TelaAtual.Inicio;

        private readonly Color corFundo = Color.FromArgb(18, 18, 18);
        private readonly Color corMenu = Color.FromArgb(12, 48, 87);
        private readonly Color corTopo = Color.FromArgb(9, 44, 80);
        private readonly Color corAzul = Color.FromArgb(45, 117, 235);
        private readonly Color corCard = Color.FromArgb(22, 22, 22);
        private readonly Color corBorda = Color.FromArgb(72, 72, 72);
        private readonly Color corTexto = Color.FromArgb(238, 238, 238);
        private readonly Color corTextoSecundario = Color.FromArgb(180, 180, 180);

        private Panel painelMenu = null!;
        private Panel painelTopo = null!;
        private Panel painelConteudo = null!;

        private Button btnInicio = null!;
        private Button btnNovaCotacao = null!;
        private Button btnConsultar = null!;
        private Button btnNaoEncontrados = null!;
        private Button btnConfiguracoes = null!;
        private Button btnAjuda = null!;

        private readonly List<string> pdfsSelecionados = new List<string>();

        private string caminhoExcelSelecionado = "";
        private string lojaSelecionada = "";
        private string lojaConsultaSelecionada = "Orlando";
        private int anoFiltroConsulta = DateTime.Now.Year;

        private int mesSelecionado = DateTime.Now.Month;
        private int anoSelecionado = DateTime.Now.Year;




        private bool EhColunaFornecedorGrid(DataGridViewColumn coluna) =>
            PreenchedorCotacao.EhFornecedor(coluna.Tag as string ?? coluna.HeaderText);
        private void AtualizarFornecedorMaisBaratoGrid(
      DataGridView dgv,
      int indiceLinha)
        {
            if (
                indiceLinha < 0 ||
                indiceLinha >= dgv.Rows.Count
            )
            {
                return;
            }

            DataGridViewRow linha =
                dgv.Rows[indiceLinha];

            int colunaFornecedorMaisBarato = -1;
            int colunaQuantidade = -1;

            // =========================================================
            // LOCALIZAR COLUNAS FIXAS
            // =========================================================

            foreach (
                DataGridViewColumn coluna
                in dgv.Columns
            )
            {
                string cabecalho =
                    NormalizarCabecalhoGrid(
                        coluna.Tag as string ?? coluna.HeaderText
                    );

                if (
                    cabecalho.Contains(
                        "FORNECEDOR MAIS EM CONTA"
                    )
                )
                {
                    colunaFornecedorMaisBarato =
                        coluna.Index;
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
                        coluna.Index;
                }
            }

            if (
                colunaFornecedorMaisBarato < 0
            )
            {
                return;
            }

            // =========================================================
            // IDENTIFICAR FORNECEDORES PELO CABEÇALHO
            //
            // No seu Excel:
            //
            // Insumo
            // Fornecedor mais em conta
            // Quantidade
            // GALENA
            // PURIFARMA
            // FAGRON
            // ...
            // =========================================================

            decimal? menorValor = null;

            string fornecedorMaisBarato = "";

            // =========================================================
            // PROCURAR SOMENTE NAS COLUNAS DOS FORNECEDORES
            // =========================================================

            for (
                int coluna = 0;
                coluna < dgv.Columns.Count;
                coluna++
            )
            {
                if (!EhColunaFornecedorGrid(dgv.Columns[coluna]))
                    continue;

                string nomeFornecedor =
                    dgv.Columns[coluna]
                        .HeaderText?
                        .Trim()
                    ?? "";

                if (
                    string.IsNullOrWhiteSpace(
                        nomeFornecedor
                    )
                )
                {
                    continue;
                }

                DataGridViewCell celula =
                    linha.Cells[coluna];

                if (
                    celula.Value == null ||
                    string.IsNullOrWhiteSpace(
                        celula.Value.ToString()
                    )
                )
                {
                    continue;
                }

                string texto =
                    celula.Value
                        .ToString()!
                        .Trim();

                decimal valor;

                bool conseguiu =
                    decimal.TryParse(
                        texto,
                        System.Globalization.NumberStyles.Any,
                        new System.Globalization.CultureInfo(
                            "pt-BR"
                        ),
                        out valor
                    );

                if (!conseguiu)
                {
                    conseguiu =
                        decimal.TryParse(
                            texto,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out valor
                        );
                }

                if (!conseguiu)
                    continue;

                // Zero não é cotação válida
                if (valor <= 0)
                    continue;

                if (
                    menorValor == null ||
                    valor < menorValor.Value
                )
                {
                    menorValor =
                        valor;

                    fornecedorMaisBarato =
                        nomeFornecedor;
                }
            }

            DataGridViewCell celulaFornecedor =
                linha.Cells[
                    colunaFornecedorMaisBarato
                ];

            // =========================================================
            // NENHUM FORNECEDOR TEM PREÇO
            // =========================================================

            if (
                menorValor == null ||
                string.IsNullOrWhiteSpace(
                    fornecedorMaisBarato
                )
            )
            {
                // Limpa fornecedor incorreto antigo
                celulaFornecedor.Value =
                    "";

                celulaFornecedor.Style.BackColor =
                    Color.FromArgb(
                        24,
                        24,
                        24
                    );

                celulaFornecedor.Style.ForeColor =
                    Color.White;

                return;
            }

            // =========================================================
            // ENCONTROU MENOR PREÇO
            // =========================================================

            celulaFornecedor.Value =
                fornecedorMaisBarato;

            celulaFornecedor.Style.BackColor =
                Color.LightGreen;

            celulaFornecedor.Style.ForeColor =
                Color.Black;
        }

        private string NormalizarCabecalhoGrid(
    string? texto)
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
                    .Replace("\r", " ")
                    .Replace("\n", " ");

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

        private void FormatarValorGridQuatroCasas(
    DataGridView dgv,
    int indiceLinha,
    int indiceColuna)
        {
            if (
                indiceLinha < 0 ||
                indiceColuna < 0
            )
            {
                return;
            }

            DataGridViewCell celula =
                dgv.Rows[indiceLinha]
                   .Cells[indiceColuna];

            if (
                celula.Value == null ||
                string.IsNullOrWhiteSpace(
                    celula.Value.ToString()
                )
            )
            {
                return;
            }

            string texto =
                celula.Value
                    .ToString()!
                    .Trim();

            decimal valor;

            bool conseguiu =
                decimal.TryParse(
                    texto,
                    System.Globalization.NumberStyles.Any,
                    new System.Globalization.CultureInfo(
                        "pt-BR"
                    ),
                    out valor
                );

            if (!conseguiu)
            {
                conseguiu =
                    decimal.TryParse(
                        texto,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out valor
                    );
            }

            if (!conseguiu)
                return;

            valor =
                Math.Round(
                    valor,
                    4,
                    MidpointRounding.AwayFromZero
                );

            celula.Value =
                valor;

            celula.Style.Format =
                "0.0000";
        }

        private readonly NaoEncontradosRepository naoEncontradosRepository = new();

        public Form1() : this(null)
        {
        }

        private Form1(string? caminhoConfiguracao)
        {
            configuracoesIA = new ConfiguracaoIARepository(caminhoConfiguracao);
            InitializeComponent();

            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1250, 780);

            BackColor = corFundo;
            Font = new Font("Segoe UI", 10F);

            try
            {
                pastaRaizCotacoes = configuracoesIA.CarregarPastaCotacoes();
                gravarJsonIA = configuracoesIA.CarregarGravacaoJson();
                percentualMinimoMatch = configuracoesIA.CarregarPercentualMinimoMatch();
                ConfiguracaoIA? salva = configuracoesIA.Carregar();
                if (salva != null)
                {
                    provedorIA = salva.Provedor;
                    modeloIA = salva.Modelo;
                    apiKeyIA = salva.ChaveApi;
                    timeoutIA = salva.TimeoutSegundos;
                    configuracaoIASalva = !string.IsNullOrWhiteSpace(apiKeyIA);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Não foi possível carregar as configurações salvas. Acesse Configurações para configurá-las novamente.",
                    "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            MontarBase();
            AbrirTelaInicio();
        }

        // =========================================================
        // BASE
        // =========================================================

        private void MontarBase()
        {
            Controls.Clear();

            painelTopo = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = corTopo
            };

            Label lblSistema = new Label
            {
                Text = "Sistema de cotação de insumos",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F),
                AutoSize = true,
                Location = new Point(18, 12)
            };

            painelTopo.Controls.Add(lblSistema);

            painelMenu = new Panel
            {
                Dock = DockStyle.Left,
                Width = 220,
                BackColor = corMenu
            };

            btnInicio = CriarBotaoMenu("⌂  Início");
            btnNovaCotacao = CriarBotaoMenu("▣  Nova cotação");
            btnConsultar = CriarBotaoMenu("⌕  Consultar cotações");
            btnNaoEncontrados = CriarBotaoMenu("⚠  Não encontrados");
            btnConfiguracoes = CriarBotaoMenu("⚙  Configurações");
            btnAjuda = CriarBotaoMenu("?  Ajuda");

            btnInicio.Location = new Point(10, 24);
            btnNovaCotacao.Location = new Point(10, 88);
            btnConsultar.Location = new Point(10, 152);
            btnNaoEncontrados.Location = new Point(10, 226);
            btnConfiguracoes.Location = new Point(10, 300);
            btnAjuda.Location = new Point(10, 374);

            btnInicio.Click += (s, e) => AbrirTelaInicio();
            btnNovaCotacao.Click += (s, e) => AbrirTelaNovaCotacao();
            btnConsultar.Click += (s, e) => AbrirTelaConsultar();
            btnNaoEncontrados.Click += (s, e) => AbrirTelaNaoEncontrados();
            btnConfiguracoes.Click += (s, e) => AbrirTelaConfiguracoes();
            btnAjuda.Click += (s, e) => AbrirTelaAjuda();

            painelMenu.Controls.Add(btnInicio);
            painelMenu.Controls.Add(btnNovaCotacao);
            painelMenu.Controls.Add(btnConsultar);
            painelMenu.Controls.Add(btnNaoEncontrados);
            painelMenu.Controls.Add(btnConfiguracoes);
            painelMenu.Controls.Add(btnAjuda);

            painelConteudo = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = corFundo,
                AutoScroll = true
            };

            Controls.Add(painelConteudo);
            Controls.Add(painelMenu);
            Controls.Add(painelTopo);
        }

        private Button CriarBotaoMenu(string texto)
        {
            Button btn = new Button();

            btn.Text = texto;
            btn.Width = 198;
            btn.Height = 54;
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = corMenu;
            btn.ForeColor = Color.White;

            btn.Font = new Font(
                "Segoe UI",
                11F,
                FontStyle.Regular
            );

            btn.TextAlign = ContentAlignment.MiddleLeft;

            btn.Padding = new Padding(
                18,
                0,
                0,
                0
            );

            btn.Cursor = Cursors.Hand;

            btn.FlatAppearance.BorderSize = 0;

            btn.FlatAppearance.MouseOverBackColor =
                Color.FromArgb(20, 65, 110);

            btn.FlatAppearance.MouseDownBackColor =
                Color.FromArgb(35, 100, 180);

            return btn;
        }

        private void DestacarMenu(Button selecionado)
        {
            foreach (Control control in painelMenu.Controls)
            {
                if (control is Button botao)
                {
                    botao.BackColor = corMenu;
                    botao.FlatAppearance.BorderSize = 0;
                }
            }

            selecionado.BackColor = corAzul;
        }

        private void LimparConteudo()
        {
            historicoTabelaAtivo?.Confirmar();
            historicoTabelaAtivo = null;
            painelConteudo.Controls.Clear();
            painelConteudo.AutoScrollPosition = Point.Empty;
            painelConteudo.AutoScroll = true;
        }

        // =========================================================
        // COMPONENTES VISUAIS
        // =========================================================

        private Panel CriarContainer(int altura)
        {
            return new Panel
            {
                Location = new Point(25, 20),

                Width = Math.Max(
                    1050,
                    painelConteudo.ClientSize.Width - 50
                ),

                Height = altura,
                BackColor = corFundo
            };
        }

        private void AjustarContainerATela(Panel tela)
        {
            painelConteudo.AutoScroll = false;
            tela.Bounds = new Rectangle(25, 20,
                Math.Max(1, painelConteudo.ClientSize.Width - 50),
                Math.Max(1, painelConteudo.ClientSize.Height - 40));
            tela.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        }
        private Panel CriarCard(int largura, int altura)
        {
            return new Panel
            {
                Width = largura,
                Height = altura,
                BackColor = corCard,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private void EstilizarCard(Panel card)
        {
            card.BorderStyle = BorderStyle.None;
            card.BackColor = Color.FromArgb(23, 28, 35);
            card.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 3,
                BackColor = Color.FromArgb(47, 91, 145) });
        }

        private void EstilizarGrade(DataGridView grade)
        {
            grade.BorderStyle = BorderStyle.None;
            grade.BackgroundColor = Color.FromArgb(20, 24, 30);
            grade.GridColor = Color.FromArgb(43, 52, 65);
            grade.ColumnHeadersHeight = 40;
            grade.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(32, 44, 60);
            grade.DefaultCellStyle.BackColor = Color.FromArgb(23, 28, 35);
            grade.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(28, 34, 42);
            grade.DefaultCellStyle.Padding = new Padding(8, 4, 8, 4);
            grade.RowTemplate.Height = 38;
            foreach (DataGridViewColumn coluna in grade.Columns)
                if (coluna is DataGridViewButtonColumn botao)
                {
                    botao.FlatStyle = FlatStyle.Flat;
                    botao.DefaultCellStyle.BackColor = Color.FromArgb(35, 59, 87);
                    botao.DefaultCellStyle.ForeColor = Color.FromArgb(204, 224, 250);
                }
        }

        private Label CriarTitulo(string texto)
        {
            return new Label
            {
                Text = texto,
                AutoSize = true,
                ForeColor = corTexto,
                Font = new Font(
                    "Segoe UI",
                    18F,
                    FontStyle.Bold
                )
            };
        }

        private Label CriarTexto(string texto)
        {
            return new Label
            {
                Text = texto,
                AutoSize = true,
                ForeColor = corTextoSecundario,
                Font = new Font(
                    "Segoe UI",
                    10F
                )
            };
        }

        private Label CriarLabelSecao(string texto)
        {
            return new Label
            {
                Text = texto,
                AutoSize = true,
                ForeColor = Color.FromArgb(
                    225,
                    225,
                    225
                ),
                Font = new Font(
                    "Segoe UI",
                    10F
                )
            };
        }

        private Button CriarBotao(
            string texto,
            int largura = 160,
            int altura = 42)
        {
            Button btn = new Button();

            btn.Text = texto;
            btn.Width = largura;
            btn.Height = altura;

            btn.BackColor = Color.FromArgb(
                24,
                24,
                24
            );

            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;

            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = corBorda;

            btn.Cursor = Cursors.Hand;

            return btn;
        }

        private Button CriarBotaoAzul(
            string texto,
            int largura = 210,
            int altura = 44)
        {
            Button btn = CriarBotao(
                texto,
                largura,
                altura
            );

            btn.BackColor = corAzul;
            btn.FlatAppearance.BorderSize = 0;

            return btn;
        }

        private ComboBox CriarCombo(int largura)
        {
            return new ComboBox
            {
                Width = largura,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(25, 25, 25),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
        }

        private NumericUpDown CriarAno(int largura)
        {
            return new NumericUpDown
            {
                Width = largura,
                Minimum = 2020,
                Maximum = 2100,
                Value = DateTime.Now.Year,
                BackColor = Color.FromArgb(25, 25, 25),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private DataGridView CriarGrid()
        {
            DataGridView dgv = new DataGridView
            {
                BackgroundColor = Color.FromArgb(18, 18, 18),
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = Color.FromArgb(65, 65, 65),

                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,

                SelectionMode =
                    DataGridViewSelectionMode.FullRowSelect,

                MultiSelect = false,

                AutoSizeColumnsMode =
                    DataGridViewAutoSizeColumnsMode.Fill,

                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34,

                ScrollBars = ScrollBars.Both,
            };

            dgv.ScrollBars = ScrollBars.Vertical;

            dgv.RowTemplate.Height = 32;

            dgv.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(24, 24, 24);

            dgv.ColumnHeadersDefaultCellStyle.ForeColor =
                Color.White;

            dgv.ColumnHeadersDefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    9.5F,
                    FontStyle.Bold
                );

            dgv.DefaultCellStyle.BackColor =
                Color.FromArgb(22, 22, 22);

            dgv.DefaultCellStyle.ForeColor =
                Color.White;

            dgv.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(40, 75, 110);

            dgv.DefaultCellStyle.SelectionForeColor =
                Color.White;

            return dgv;
        }

        // =========================================================
        // PASTAS
        // =========================================================

        private string pastaRaizCotacoes = "";

        private string ObterPastaRaiz()
        {
            if (!string.IsNullOrWhiteSpace(pastaRaizCotacoes)) return pastaRaizCotacoes;
            return Environment.GetFolderPath(
                Environment.SpecialFolder.MyDocuments
            );
        }

        private string ObterPastaEmpresa(string empresa)
        {
            return Path.Combine(
                ObterPastaRaiz(),
                empresa
            );
        }

        private string ObterPastaPeriodo(
            string empresa,
            int mes,
            int ano)
        {
            return Path.Combine(
                ObterPastaEmpresa(empresa),
                $"{ano}-{mes:00}"
            );
        }

        private string ObterCaminhoPlanilhaPeriodo(
            string empresa,
            int mes,
            int ano)
        {
            string pasta =
                ObterPastaPeriodo(
                    empresa,
                    mes,
                    ano
                );

            return Path.Combine(
                pasta,
                $"Cotacao_{empresa}_{ano}-{mes:00}.xlsx"
            );
        }

        private bool PeriodoExiste(
            string empresa,
            int mes,
            int ano)
        {
            string pasta =
                ObterPastaPeriodo(
                    empresa,
                    mes,
                    ano
                );

            return Directory.Exists(pasta);
        }

        private bool PlanilhaPeriodoExiste(
            string empresa,
            int mes,
            int ano)
        {
            string caminho =
                ObterCaminhoPlanilhaPeriodo(
                    empresa,
                    mes,
                    ano
                );

            return File.Exists(caminho);
        }

        // =========================================================
        // INÍCIO
        // =========================================================

        private void AbrirTelaInicio()
        {
            telaAtual = TelaAtual.Inicio;

            LimparConteudo();
            DestacarMenu(btnInicio);

            Panel tela = CriarContainer(650);

            Label titulo =
                CriarTitulo("Bem-vindo");

            titulo.Location =
                new Point(10, 5);

            Panel cardNova =
                CriarCard(400, 140);

            cardNova.Location =
                new Point(10, 60);

            Label tituloNova =
                new Label
                {
                    Text = "Nova cotação",
                    AutoSize = true,
                    ForeColor = Color.White,
                    Font = new Font(
                        "Segoe UI",
                        13F,
                        FontStyle.Bold
                    ),
                    Location = new Point(130, 25)
                };

            Label descNova =
                CriarTexto(
                    "Excel opcional, PDFs obrigatórios"
                );

            descNova.Location =
                new Point(90, 57);

            Button iniciar =
                CriarBotao(
                    "Iniciar",
                    120,
                    38
                );

            iniciar.Location =
                new Point(140, 88);

            iniciar.Click +=
                (s, e) =>
                {
                    PrepararNovaCotacao();
                    AbrirTelaNovaCotacao();
                };

            cardNova.Controls.Add(tituloNova);
            cardNova.Controls.Add(descNova);
            cardNova.Controls.Add(iniciar);

            Panel cardConsulta =
                CriarCard(400, 140);

            cardConsulta.Location =
                new Point(430, 60);

            Label tituloConsulta =
                new Label
                {
                    Text = "Consultar cotações",
                    AutoSize = true,
                    ForeColor = Color.White,
                    Font = new Font(
                        "Segoe UI",
                        13F,
                        FontStyle.Bold
                    ),
                    Location = new Point(105, 25)
                };

            Label descConsulta =
                CriarTexto(
                    "Planilhas geradas por mês"
                );

            descConsulta.Location =
                new Point(110, 57);

            Button consultar =
                CriarBotao(
                    "Consultar",
                    120,
                    38
                );

            consultar.Location =
                new Point(140, 88);

            consultar.Click +=
                (s, e) =>
                    AbrirTelaConsultar();

            cardConsulta.Controls.Add(tituloConsulta);
            cardConsulta.Controls.Add(descConsulta);
            cardConsulta.Controls.Add(consultar);

            Label acesso =
                CriarTexto(
                    "Acesso rápido às pastas"
                );

            acesso.Location =
                new Point(10, 230);

            Button orlando =
                CriarBotao(
                    "▣ Abrir pasta Orlando",
                    230,
                    42
                );

            orlando.Location =
                new Point(10, 260);

            orlando.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        "Orlando";

                    AbrirTelaConsultar();
                };

            Button drugstore =
                CriarBotao(
                    "▣ Abrir pasta Drugstore",
                    230,
                    42
                );

            drugstore.Location =
                new Point(255, 260);

            drugstore.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        "Drugstore";

                    AbrirTelaConsultar();
                };

            Button naoEncontrados =
                CriarBotao(
                    "⚠ Insumos não encontrados",
                    280,
                    42
                );

            naoEncontrados.Location =
                new Point(10, 318);

            naoEncontrados.Click +=
                (s, e) =>
                    AbrirTelaNaoEncontrados();

            AjustarContainerATela(tela);
            iniciar.Size = new Size(160, 44);
            consultar.Size = new Size(160, 44);
            EstilizarCard(cardNova);
            EstilizarCard(cardConsulta);
            iniciar.Text = "Criar nova cotação";
            iniciar.BackColor = corAzul;
            iniciar.FlatAppearance.BorderSize = 0;
            iniciar.Width = 210;
            consultar.Width = 210;
            consultar.Text = "Ver minhas cotações";
            descNova.Text = "Selecione os PDFs e preencha sua planilha.";
            descConsulta.Text = "Acesse suas cotações organizadas por período.";
            acesso.Text = "ACESSO RÁPIDO";
            acesso.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            orlando.Text = "Cotações Orlando";
            drugstore.Text = "Cotações Drugstore";
            naoEncontrados.Text = "Insumos não encontrados";
            naoEncontrados.ForeColor = Color.FromArgb(245, 190, 85);
            Label boasVindas = CriarTexto("Organize cotações, compare fornecedores e acompanhe seus insumos.");
            boasVindas.Location = new Point(10, 45);
            tela.Controls.Add(boasVindas);
            void AjustarLayoutInicio()
            {
                bool compacto = tela.ClientSize.Width < 820;
                int larguraCard = Math.Max(300, (tela.ClientSize.Width - (compacto ? 20 : 38)) / (compacto ? 1 : 2));
                cardNova.SetBounds(10, 96, larguraCard, 220);
                cardConsulta.SetBounds(compacto ? 10 : cardNova.Right + 18,
                    compacto ? cardNova.Bottom + 18 : cardNova.Top, larguraCard, 220);
                tituloNova.Location = new Point(24, 28);
                descNova.Location = new Point(24, 70);
                iniciar.Location = new Point(24, 145);
                tituloConsulta.Location = new Point(24, 28);
                descConsulta.Location = new Point(24, 70);
                consultar.Location = new Point(24, 145);
                descNova.MaximumSize = new Size(larguraCard - 48, 0);
                descConsulta.MaximumSize = new Size(larguraCard - 48, 0);
                acesso.Top = cardConsulta.Bottom + 32;
                orlando.Bounds = new Rectangle(10, acesso.Bottom + 16, larguraCard, 50);
                drugstore.Bounds = new Rectangle(compacto ? 10 : cardConsulta.Left,
                    compacto ? orlando.Bottom + 12 : orlando.Top, larguraCard, 50);
                naoEncontrados.Bounds = new Rectangle(10, drugstore.Bottom + 14, larguraCard, 50);
                tela.AutoScroll = true;
            }
            tela.Resize += (s, e) => AjustarLayoutInicio();
            AjustarLayoutInicio();
            tela.Controls.Add(titulo);
            tela.Controls.Add(cardNova);
            tela.Controls.Add(cardConsulta);
            tela.Controls.Add(acesso);
            tela.Controls.Add(orlando);
            tela.Controls.Add(drugstore);
            tela.Controls.Add(naoEncontrados);

            painelConteudo.Controls.Add(tela);
        }

        // =========================================================
        // NOVA COTAÇÃO
        // =========================================================

        private void AbrirTelaNovaCotacao()
        {
            telaAtual =
                TelaAtual.NovaCotacao;

            LimparConteudo();
            DestacarMenu(btnNovaCotacao);

            Panel tela =
                CriarContainer(780);

            Label titulo =
                CriarTitulo("Nova cotação");

            titulo.Location =
                new Point(10, 0);

            // -----------------------------------------------------
            // LOJA / MÊS / ANO
            // -----------------------------------------------------

            Panel cardLoja =
                CriarCard(410, 220);

            cardLoja.Location =
                new Point(10, 55);

            Label lblLoja =
                CriarLabelSecao("1. Empresa");

            lblLoja.Location =
                new Point(10, 15);

            Button btnOrlando =
                CriarBotao(
                    "Orlando",
                    150,
                    40
                );

            btnOrlando.Location =
                new Point(10, 45);

            Button btnDrugstore =
                CriarBotao(
                    "Drugstore",
                    150,
                    40
                );

            btnDrugstore.Location =
                new Point(170, 45);

            if (lojaSelecionada == "Orlando")
                btnOrlando.BackColor = corAzul;

            if (lojaSelecionada == "Drugstore")
                btnDrugstore.BackColor = corAzul;

            Label lblMesAno =
                CriarLabelSecao("2. Mês e ano ");

            lblMesAno.Location =
                new Point(10, 105);

            ComboBox cmbMes =
                CriarCombo(240);

            cmbMes.Location =
                new Point(10, 135);

            cmbMes.Items.AddRange(
                new string[]
                {
                    "Janeiro",
                    "Fevereiro",
                    "Março",
                    "Abril",
                    "Maio",
                    "Junho",
                    "Julho",
                    "Agosto",
                    "Setembro",
                    "Outubro",
                    "Novembro",
                    "Dezembro"
                }
            );

            cmbMes.SelectedIndex =
                mesSelecionado - 1;

            NumericUpDown numAno =
                CriarAno(110);

            numAno.Location =
                new Point(260, 135);

            numAno.Value =
                anoSelecionado;

            Label lblAviso =
                new Label
                {
                    AutoSize = false,
                    Width = 360,
                    Height = 30,
                    Location = new Point(10, 175),

                    Font = new Font(
                        "Segoe UI",
                        9.5F
                    ),

                    TextAlign =
                        ContentAlignment.MiddleLeft,

                    Padding =
                        new Padding(8, 6, 8, 6)
                };

            void AtualizarAviso()
            {
                if (string.IsNullOrWhiteSpace(lojaSelecionada))
                {
                    lblAviso.Text =
                        "Selecione Orlando ou Drugstore.";

                    lblAviso.ForeColor =
                        Color.LightGray;

                    lblAviso.BackColor =
                        Color.FromArgb(45, 45, 45);

                    return;
                }

                bool existe =
                    PeriodoExiste(
                        lojaSelecionada,
                        mesSelecionado,
                        anoSelecionado
                    );

                if (existe)
                {
                    lblAviso.Text =
                        $"A cotação existente será complementada na pasta {anoSelecionado}-{mesSelecionado:00}. Não é necessário selecionar outra planilha-base.";

                    lblAviso.ForeColor =
                        Color.FromArgb(255, 183, 0);

                    lblAviso.BackColor =
                        Color.FromArgb(58, 38, 0);
                }
                else
                {
                    lblAviso.Text =
                        $"A cotação será salva no período {mesSelecionado:00}/{anoSelecionado}. A cotação será criada a partir da planilha-base selecionada.";

                    lblAviso.ForeColor =
                        Color.LightGreen;

                    lblAviso.BackColor =
                        Color.FromArgb(15, 60, 25);
                }
            }

            btnOrlando.Click +=
                (s, e) =>
                {
                    lojaSelecionada =
                        "Orlando";

                    btnOrlando.BackColor =
                        corAzul;

                    btnDrugstore.BackColor =
                        Color.FromArgb(24, 24, 24);

                    AtualizarAviso();
                };

            btnDrugstore.Click +=
                (s, e) =>
                {
                    lojaSelecionada =
                        "Drugstore";

                    btnDrugstore.BackColor =
                        corAzul;

                    btnOrlando.BackColor =
                        Color.FromArgb(24, 24, 24);

                    AtualizarAviso();
                };

            cmbMes.SelectedIndexChanged +=
                (s, e) =>
                {
                    mesSelecionado =
                        cmbMes.SelectedIndex + 1;

                    AtualizarAviso();
                };

            numAno.ValueChanged +=
                (s, e) =>
                {
                    anoSelecionado =
                        (int)numAno.Value;

                    AtualizarAviso();
                };

            AtualizarAviso();

            cardLoja.Controls.Add(lblLoja);
            cardLoja.Controls.Add(btnOrlando);
            cardLoja.Controls.Add(btnDrugstore);
            cardLoja.Controls.Add(lblMesAno);
            cardLoja.Controls.Add(cmbMes);
            cardLoja.Controls.Add(numAno);
            cardLoja.Controls.Add(lblAviso);

            // -----------------------------------------------------
            // EXCEL
            // -----------------------------------------------------

            Panel cardExcel =
                CriarCard(460, 220);

            cardExcel.Location =
                new Point(425, 55);

            Label lblExcel =
                CriarLabelSecao(
                    "3. Planilha Excel (opcional)"
                );

            lblExcel.Location =
                new Point(18, 15);

            Panel painelExcel =
                new Panel
                {
                    Location =
                        new Point(15, 45),

                    Width =
                        430,

                    Height =
                        52,

                    BorderStyle =
                        BorderStyle.FixedSingle,

                    BackColor =
                        Color.FromArgb(25, 25, 25)
                };

            Label iconeExcel =
                new Label
                {
                    Text = "▣",
                    ForeColor = Color.LimeGreen,

                    Font = new Font(
                        "Segoe UI",
                        18F,
                        FontStyle.Bold
                    ),

                    AutoSize = true,
                    Location = new Point(15, 10)
                };

            Label arquivoExcel =
                new Label
                {
                    Text =
                        string.IsNullOrWhiteSpace(
                            caminhoExcelSelecionado
                        )
                        ? "Nenhuma planilha selecionada"
                        : Path.GetFileName(
                            caminhoExcelSelecionado
                        ),

                    ForeColor = Color.White,

                    AutoSize = false,

                    Width = 240,
                    Height = 30,

                    Location =
                        new Point(50, 14)
                };

            Button removerExcel =
                new Button
                {
                    Text = "×",
                    Width = 30,
                    Height = 30,

                    Location =
                        new Point(290, 9),

                    FlatStyle =
                        FlatStyle.Flat,

                    BackColor =
                        Color.FromArgb(25, 25, 25),

                    ForeColor =
                        Color.LightGray,

                    Cursor =
                        Cursors.Hand
                };

            removerExcel.FlatAppearance.BorderSize =
                0;

            removerExcel.Click +=
                (s, e) =>
                {
                    caminhoExcelSelecionado = "";
                    AbrirTelaNovaCotacao();
                };

            Button selecionarExcel =
                CriarBotao(
                    "Selecionar",
                    95,
                    34
                );

            selecionarExcel.Location =
                new Point(325, 8);

            selecionarExcel.Click +=
                (s, e) =>
                {
                    OpenFileDialog dialog =
                        new OpenFileDialog
                        {
                            Filter =
                                "Planilhas Excel|*.xlsx;*.xls"
                        };

                    if (
                        dialog.ShowDialog() ==
                        DialogResult.OK
                    )
                    {
                        caminhoExcelSelecionado =
                            dialog.FileName;

                        AbrirTelaNovaCotacao();
                    }
                };

            painelExcel.Controls.Add(iconeExcel);
            painelExcel.Controls.Add(arquivoExcel);
            painelExcel.Controls.Add(removerExcel);
            painelExcel.Controls.Add(selecionarExcel);

            Label obs1 =
     new Label
     {
         Text =
             "Nova cotação: selecione a planilha-base.\n" +
             "A cotação será salva na pasta do ano-mês e reutilizada se já existir.",

         ForeColor =
             corTextoSecundario,

         Font =
             new Font(
                 "Segoe UI",
                 9.5F
             ),

         AutoSize =
             false,

         Width =
             410,

         Height =
             42,

         Location =
             new Point(
                 18,
                 115
             )
     };

            Label obs2 =
                new Label
                {
                    Text =
                        "Cotação nova: selecione o Excel-base para criar uma nova planilha.",

                    ForeColor =
                        corTextoSecundario,

                    Font =
                        new Font(
                            "Segoe UI",
                            9.5F
                        ),

                    AutoSize =
                        false,

                    Width =
                        410,

                    Height =
                        30,

                    Location =
                        new Point(
                            18,
                            165
                        )
                };

            cardExcel.Controls.Add(lblExcel);
            cardExcel.Controls.Add(painelExcel);
            cardExcel.Controls.Add(obs1);
            cardExcel.Controls.Add(obs2);

            // -----------------------------------------------------
            // PDFS
            // -----------------------------------------------------

            Panel cardPdfs =
                CriarCard(875, 325);

            cardPdfs.Location =
                new Point(10, 290);

            Label lblPdfs =
                CriarLabelSecao(
                    "4. PDFs dos fornecedores * (até 6 ou mais)"
                );

            lblPdfs.Location =
                new Point(10, 15);

            DataGridView dgvPdfs =
                CriarGrid();

            dgvPdfs.Location =
                new Point(10, 50);

            dgvPdfs.Size =
                new Size(855, 190);

            dgvPdfs.Columns.Add(
                "Arquivo",
                "Arquivo"
            );

            dgvPdfs.Columns.Add(
                "Fornecedor",
                "Fornecedor detectado"
            );

            dgvPdfs.Columns.Add(
                "Situacao",
                "Situação"
            );

            DataGridViewButtonColumn excluirPdf =
                new DataGridViewButtonColumn
                {
                    Name = "Excluir",
                    HeaderText = "",
                    Text = "×",
                    UseColumnTextForButtonValue = true,
                    Width = 40
                };

            dgvPdfs.Columns.Add(excluirPdf);

            foreach (string pdf in pdfsSelecionados)
            {
                dgvPdfs.Rows.Add(
                    Path.GetFileName(pdf),
                    "A detectar",
                    "Aguardando"
                );
            }

            dgvPdfs.CellContentClick +=
                (s, e) =>
                {
                    if (e.RowIndex < 0)
                        return;

                    if (
                        dgvPdfs.Columns[
                            e.ColumnIndex
                        ].Name != "Excluir"
                    )
                        return;

                    string nome =
                        dgvPdfs.Rows[
                            e.RowIndex
                        ]
                        .Cells["Arquivo"]
                        .Value?
                        .ToString()
                        ?? "";

                    string? caminho =
                        pdfsSelecionados
                            .FirstOrDefault(
                                x =>
                                    Path.GetFileName(x)
                                    == nome
                            );

                    if (caminho != null)
                    {
                        pdfsSelecionados.Remove(caminho);
                        AbrirTelaNovaCotacao();
                    }
                };

            Button selecionarPdfs =
                CriarBotao(
                    "▣ Selecionar PDFs",
                    170,
                    42
                );

            selecionarPdfs.Location =
                new Point(10, 255);

            selecionarPdfs.Click +=
                (s, e) =>
                {
                    OpenFileDialog dialog =
                        new OpenFileDialog
                        {
                            Filter =
                                "Arquivos PDF|*.pdf",

                            Multiselect =
                                true
                        };

                    if (
                        dialog.ShowDialog() ==
                        DialogResult.OK
                    )
                    {
                        foreach (
                            string arquivo
                            in dialog.FileNames
                        )
                        {
                            if (
                                !pdfsSelecionados
                                    .Contains(arquivo)
                            )
                            {
                                pdfsSelecionados.Add(
                                    arquivo
                                );
                            }
                        }

                        AbrirTelaNovaCotacao();
                    }
                };

            cardPdfs.Controls.Add(lblPdfs);
            cardPdfs.Controls.Add(dgvPdfs);
            cardPdfs.Controls.Add(selecionarPdfs);

            // -----------------------------------------------------
            // PROCESSAR
            // -----------------------------------------------------

            Button processar =
                CriarBotaoAzul(
                    "▶  Processar cotações",
                    220,
                    46
                );

            processar.Location =
                new Point(665, 635);

            processar.Click += async (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(lojaSelecionada) || pdfsSelecionados.Count == 0 ||
                    (!File.Exists(caminhoExcelSelecionado) &&
                     !PlanilhaPeriodoExiste(lojaSelecionada, mesSelecionado, anoSelecionado)))
                {
                    MessageBox.Show(this, "Selecione a unidade, a planilha e os PDFs dos fornecedores.",
                        "Arquivos necessários", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (!configuracaoIASalva || string.IsNullOrWhiteSpace(apiKeyIA))
                {
                    MessageBox.Show(this, "Salve a configuração da IA antes de processar.",
                        "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string pasta = ObterPastaPeriodo(lojaSelecionada, mesSelecionado, anoSelecionado);
                var arquivos = pdfsSelecionados.ToArray();
                string caminhoComparacao = ObterCaminhoPlanilhaPeriodo(lojaSelecionada, mesSelecionado, anoSelecionado);
                var nomesPlanilha = new Lazy<List<string>>(() => LeitorPlanilhaInsumos.Ler(caminhoComparacao));
                var errosProcessamento = new List<(Exception Erro, string Arquivo)>();
                string arquivoProcessamento = caminhoExcelSelecionado;
                var falhas = new List<string>();
                var falhasPersistencia = new List<string>();
                var falhasPreenchimento = new List<string>();
                var pendenciasPreenchimento = new List<string>();
                int precosPreenchidos = 0, fornecedoresCriados = 0;
                string lojaProcessamento = lojaSelecionada;
                int mesProcessamento = mesSelecionado, anoProcessamento = anoSelecionado;
                int salvos = 0, camposSalvos = 0, pendentes = 0, registrosPersistidos = 0;
                string? resumoConclusao = null;
                tela.Enabled = false;
                tela.Visible = false;
                var telaEspera = CriarTelaEsperaProcessamento(out var lblStatus);
                painelConteudo.Controls.Add(telaEspera);
                telaEspera.BringToFront();
                painelMenu.Enabled = false;
                UseWaitCursor = true;
                try
                {
                    lblStatus.Text = "Preparando a planilha da cotação do período...";
                    await Task.Yield();
                    PreparadorPlanilhaPeriodo.Preparar(caminhoExcelSelecionado, caminhoComparacao);
                    var agente = AgenteIAFactory.Criar(provedorIA, apiKeyIA, modeloIA, timeoutIA);
                    for (int i = 0; i < arquivos.Length; i++)
                    {
                        arquivoProcessamento = arquivos[i];
                        lblStatus.Text = $"Analisando PDF {i + 1}/{arquivos.Length}...\n\n" +
                            Path.GetFileName(arquivos[i]) + "\n\n" +
                            $"PDFs extraídos: {salvos}";
                        try
                        {
                            var resultado = await agente.ExtrairInsumosAsync(arquivos[i]);
                            pendentes += resultado.Insumos.Count(x => x.Preco == null || x.Unidade == null);
                            salvos++;
                            try
                            {
                                lblStatus.Text = $"Comparando insumos do PDF {i + 1}/{arquivos.Length} com a planilha...";
                                await Task.Yield();
                                var campos = ComparadorInsumos.Comparar(resultado, nomesPlanilha.Value, percentualMinimoMatch);
                                if (gravarJsonIA)
                                {
                                    try
                                    {
                                        lblStatus.Text = $"Salvando JSON _campos do PDF {i + 1}/{arquivos.Length}...";
                                        await ExportadorInsumos.SalvarCamposAsync(campos, arquivos[i], pasta);
                                        camposSalvos++;
                                    }
                                    catch (Exception ex)
                                    {
                                        errosProcessamento.Add((ex, arquivoProcessamento));
                                        falhas.Add($"{Path.GetFileName(arquivos[i])}: falha ao gravar JSON _campos. {ex.Message}");
                                    }
                                }
                                try
                                {
                                    lblStatus.Text = $"Preenchendo planilha do PDF {i + 1}/{arquivos.Length}...";
                                    await Task.Yield();
                                    var preenchimento = PreenchedorCotacao.Preencher(caminhoComparacao, campos, percentualMinimoMatch);
                                    precosPreenchidos += preenchimento.PrecosPreenchidos;
                                    fornecedoresCriados += preenchimento.FornecedoresCriados;
                                    pendenciasPreenchimento.AddRange(preenchimento.Pendencias.Select(p =>
                                        $"{Path.GetFileName(arquivos[i])}: {p}"));
                                }
                                catch (Exception ex)
                                {
                                    errosProcessamento.Add((ex, arquivoProcessamento));
                                    falhasPreenchimento.Add($"{Path.GetFileName(arquivos[i])}: {ex.Message}");
                                }
                                try
                                {
                                    lblStatus.Text = $"Gravando não encontrados do PDF {i + 1}/{arquivos.Length}...";
                                    await Task.Yield();
                                    registrosPersistidos += naoEncontradosRepository.SalvarLote(lojaProcessamento,
                                        mesProcessamento, anoProcessamento, campos, DateTime.Now);
                                }
                                catch (Exception ex)
                                {
                                    errosProcessamento.Add((ex, arquivoProcessamento));
                                    falhasPersistencia.Add($"{Path.GetFileName(arquivos[i])}: {ex.Message}");
                                }
                            }
                            catch (Exception ex)
                            {
                                errosProcessamento.Add((ex, arquivoProcessamento));
                                falhas.Add($"{Path.GetFileName(arquivos[i])}: comparação não concluída. " +
                                    $"{ex.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            errosProcessamento.Add((ex, arquivoProcessamento));
                            falhas.Add($"{Path.GetFileName(arquivos[i])}: {ex.Message}");
                        }
                    }
                    string mensagem = $"{salvos} de {arquivos.Length} PDF(s) extraído(s) e " +
                        $"{camposSalvos} JSON(s) _campos salvo(s) em:\n{pasta}";
                    if (pendenciasPreenchimento.Count > 0)
                        mensagem += "\n\nPendências de preenchimento:\n" + string.Join("\n", pendenciasPreenchimento);
                    if (falhasPreenchimento.Count > 0)
                        mensagem += "\n\nFalhas ao salvar a planilha:\n" + string.Join("\n\n", falhasPreenchimento);
                    if (pendentes > 0)
                        mensagem += $"\n\n{pendentes} insumo(s) preservados com preço ou unidade indeterminados (null).";
                    if (falhas.Count > 0)
                        mensagem += "\n\nPDFs que precisam de nova tentativa:\n" + string.Join("\n\n", falhas);
                    if (falhasPersistencia.Count > 0)
                        mensagem += "\n\nFalhas ao gravar não encontrados no banco:\n" +
                            string.Join("\n\n", falhasPersistencia);
                    lblStatus.Text = $"Extração finalizada.\n\nPDFs extraídos: {salvos}\n" +
                        $"JSONs _campos salvos: {camposSalvos}\n" +
                        $"Falhas: {falhas.Count + falhasPersistencia.Count + falhasPreenchimento.Count}";
                    resumoConclusao = mensagem;
                }
                catch (Exception ex)
                {
                    errosProcessamento.Add((ex, arquivoProcessamento));
                }
                finally
                {
                    painelConteudo.Controls.Remove(telaEspera);
                    telaEspera.Dispose();
                    tela.Visible = true;
                    tela.Enabled = true;
                    painelMenu.Enabled = true;
                    UseWaitCursor = false;
                }
                foreach (var erro in errosProcessamento)
                    ModalErroIA.Exibir(this, erro.Erro, provedorIA, modeloIA, arquivosPreservados: true, arquivo: erro.Arquivo);

                if (resumoConclusao != null)
                    AbrirTelaConcluido(lojaProcessamento, mesProcessamento, anoProcessamento,
                        caminhoComparacao, resumoConclusao, precosPreenchidos, registrosPersistidos,
                        falhas.Count > 0 || falhasPersistencia.Count > 0 || pendentes > 0 ||
                        falhasPreenchimento.Count > 0 || pendenciasPreenchimento.Count > 0);
            };

            AjustarContainerATela(tela);
            EstilizarCard(cardLoja);
            EstilizarCard(cardExcel);
            EstilizarCard(cardPdfs);
            EstilizarGrade(dgvPdfs);
            lblLoja.Font = lblMesAno.Font = lblExcel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            selecionarPdfs.Text = "Adicionar PDFs";
            selecionarPdfs.BackColor = Color.FromArgb(35, 59, 87);
            processar.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            painelExcel.BackColor = Color.FromArgb(30, 38, 48);
            painelExcel.BorderStyle = BorderStyle.None;
            iconeExcel.Text = "X";
            Label ajudaNova = CriarTexto("Escolha a empresa, o período e os arquivos dos fornecedores para começar.");
            ajudaNova.Location = new Point(10, 43);
            tela.Controls.Add(ajudaNova);
            tela.AutoScroll = false;
            tela.AutoScrollMinSize = Size.Empty;
            tela.AutoScrollMargin = Size.Empty;
            void AjustarLayoutNovaCotacao()
            {
                int largura = Math.Max(1, tela.ClientSize.Width - 20);
                bool compacto = largura < 940;
                cardLoja.Location = new Point(10, 84);
                cardLoja.Width = compacto ? largura : (largura - 18) / 2;
                cardExcel.Left = compacto ? 10 : cardLoja.Right + 18;
                cardExcel.Width = compacto ? largura : largura - cardLoja.Width - 18;
                lblAviso.Width = Math.Max(1, cardLoja.ClientSize.Width - 20);
                lblAviso.Height = lblAviso.GetPreferredSize(new Size(lblAviso.Width, 0)).Height;
                cardLoja.Height = Math.Max(220, lblAviso.Bottom + 15);
                cardExcel.Top = compacto ? cardLoja.Bottom + 18 : cardLoja.Top;
                cardExcel.Height = cardLoja.Height;
                painelExcel.Width = cardExcel.Width - 30;
                selecionarExcel.Left = painelExcel.Width - selecionarExcel.Width - 12;
                removerExcel.Left = selecionarExcel.Left - removerExcel.Width - 8;
                arquivoExcel.Width = Math.Max(80, removerExcel.Left - arquivoExcel.Left - 8);
                cardPdfs.Top = Math.Max(cardLoja.Bottom, cardExcel.Bottom) + 15;
                cardPdfs.Width = largura;
                int topoProcessar = Math.Max(0, tela.ClientSize.Height - processar.Height - 12);
                cardPdfs.Height = Math.Max(0, topoProcessar - cardPdfs.Top - 16);
                dgvPdfs.Width = Math.Max(1, cardPdfs.ClientSize.Width - 20);
                selecionarPdfs.Top = Math.Max(0, cardPdfs.ClientSize.Height - selecionarPdfs.Height - 16);
                dgvPdfs.Height = Math.Max(0, selecionarPdfs.Top - dgvPdfs.Top - 14);
                processar.Location = new Point(Math.Max(10, cardPdfs.Right - processar.Width),
                    topoProcessar);
            }
            tela.Resize += (s, e) => AjustarLayoutNovaCotacao();
            lblAviso.TextChanged += (s, e) => AjustarLayoutNovaCotacao();
            AjustarLayoutNovaCotacao();
            tela.Controls.Add(titulo);
            tela.Controls.Add(cardLoja);
            tela.Controls.Add(cardExcel);
            tela.Controls.Add(cardPdfs);
            tela.Controls.Add(processar);

            painelConteudo.Controls.Add(tela);
        }

        // =========================================================
        // CONCLUSÃO
        // =========================================================

        private void AbrirTelaConcluido(string empresa, int mes, int ano, string caminhoPlanilha,
            string resumo, int insumosPreenchidos, int insumosNaoEncontrados, bool comPendencias)
        {
            telaAtual = TelaAtual.Concluido;
            LimparConteudo();
            Panel tela = CriarContainer(700);

            Label icone = new()
            {
                Text = comPendencias ? "!" : "✓",
                AutoSize = false,
                Size = new Size(38, 38),
                Location = new Point(15, 16),
                ForeColor = comPendencias ? Color.Orange : Color.LimeGreen,
                Font = new Font("Segoe UI", 22F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            Label titulo = CriarTitulo("Processamento concluído");
            titulo.Location = new Point(65, 15);
            Label subtitulo = CriarTexto(comPendencias
                ? "O processamento terminou com pendências. Confira o resumo abaixo."
                : "A leitura, a comparação e a gravação foram concluídas com sucesso.");
            subtitulo.Location = new Point(65, 53);

            Color verde = Color.FromArgb(75, 210, 145);
            Color amarelo = Color.FromArgb(245, 190, 85);
            icone.BackColor = comPendencias ? Color.FromArgb(54, 43, 25) : Color.FromArgb(23, 49, 37);
            icone.ForeColor = comPendencias ? amarelo : verde;
            icone.Size = new Size(46, 46);
            icone.Location = new Point(10, 16);
            titulo.Location = new Point(72, 12);
            subtitulo.Location = new Point(72, 52);

            Panel CriarIndicador(string nome, string texto, int quantidade, Color destaque)
            {
                Panel card = CriarCard(400, 132);
                card.Name = nome;
                card.BorderStyle = BorderStyle.None;
                card.BackColor = Color.FromArgb(24, 29, 36);
                Panel faixa = new() { Dock = DockStyle.Left, Width = 4, BackColor = destaque };
                Label legenda = CriarTexto(texto);
                legenda.Location = new Point(22, 18);
                legenda.Font = new Font("Segoe UI", 10F);
                legenda.ForeColor = Color.FromArgb(186, 196, 208);
                Label valor = new()
                {
                    Text = quantidade.ToString("N0"), AutoSize = true,
                    Location = new Point(18, 45), ForeColor = destaque,
                    Font = new Font("Segoe UI", 30F, FontStyle.Bold)
                };
                card.Controls.AddRange([faixa, legenda, valor]);
                return card;
            }
            Panel preenchidos = CriarIndicador("cardInsumosPreenchidos",
                "Insumos preenchidos na planilha", insumosPreenchidos, verde);
            Panel naoEncontrados = CriarIndicador("cardInsumosNaoEncontrados",
                "Insumos não encontrados", insumosNaoEncontrados, amarelo);

            Panel informacoes = CriarCard(850, 180);
            informacoes.BorderStyle = BorderStyle.None;
            informacoes.BackColor = Color.FromArgb(23, 26, 31);
            Label cabecalhoResumo = CriarLabelSecao("RESUMO DO PROCESSAMENTO");
            cabecalhoResumo.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            cabecalhoResumo.ForeColor = Color.FromArgb(151, 166, 185);
            cabecalhoResumo.Location = new Point(22, 18);
            Label periodo = CriarTexto($"{empresa}  •  {ObterNomeMes(mes)} de {ano}");
            periodo.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            periodo.ForeColor = Color.WhiteSmoke;
            periodo.Location = new Point(22, 46);
            Panel conteudoResumo = new() { Location = new Point(22, 83), AutoScroll = true };
            Label detalhes = new()
            {
                Text = resumo,
                AutoSize = true,
                MaximumSize = new Size(780, 0),
                ForeColor = Color.FromArgb(186, 196, 208),
                Font = new Font("Segoe UI", 10F),
                Location = Point.Empty
            };
            conteudoResumo.Controls.Add(detalhes);
            informacoes.Controls.AddRange([cabecalhoResumo, periodo, conteudoResumo]);

            Button visualizarPlanilha = CriarBotaoAzul("Visualizar planilha de cotação", 285, 46);
            visualizarPlanilha.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            visualizarPlanilha.Location = new Point(10, 470);
            visualizarPlanilha.Click += (s, e) =>
            {
                if (!File.Exists(caminhoPlanilha))
                {
                    MessageBox.Show(this, "A planilha desta cotação não foi encontrada.",
                        "Arquivo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                lojaConsultaSelecionada = empresa;
                anoFiltroConsulta = ano;
                AbrirTabelaCotacaoArquivo(caminhoPlanilha, empresa, mes, ano);
            };
            Button verNaoEncontrados = CriarBotao("Visualizar não encontrados", 265, 46);
            verNaoEncontrados.Font = new Font("Segoe UI", 10F);
            verNaoEncontrados.Location = new Point(310, 470);
            verNaoEncontrados.Click += (s, e) =>
            {
                lojaConsultaSelecionada = empresa;
                anoFiltroConsulta = ano;
                mesSelecionado = mes;
                anoSelecionado = ano;
                AbrirTelaNaoEncontrados();
            };
            tela.Controls.AddRange([icone, titulo, subtitulo, preenchidos, naoEncontrados, informacoes, visualizarPlanilha, verNaoEncontrados]);
            AjustarContainerATela(tela);
            void AjustarResumo()
            {
                int largura = Math.Max(300, tela.ClientSize.Width - 20);
                bool compacto = largura < 680;
                int larguraCard = compacto ? largura : (largura - 18) / 2;
                preenchidos.SetBounds(10, 108, larguraCard, 132);
                naoEncontrados.SetBounds(compacto ? 10 : preenchidos.Right + 18,
                    compacto ? preenchidos.Bottom + 14 : preenchidos.Top, larguraCard, 132);
                informacoes.SetBounds(10, naoEncontrados.Bottom + 22, largura, 180);
                conteudoResumo.Size = new Size(Math.Max(1, largura - 44), 77);
                detalhes.MaximumSize = new Size(Math.Max(200, conteudoResumo.Width -
                    SystemInformation.VerticalScrollBarWidth), 0);
                visualizarPlanilha.Location = new Point(10, informacoes.Bottom + 22);
                verNaoEncontrados.Location = new Point(compacto ? 10 : visualizarPlanilha.Right + 14,
                    compacto ? visualizarPlanilha.Bottom + 12 : visualizarPlanilha.Top);
                tela.AutoScroll = true;
                tela.AutoScrollMinSize = new Size(0, verNaoEncontrados.Bottom + 20);
            }
            tela.Resize += (s, e) => AjustarResumo();
            AjustarResumo();
            painelConteudo.Controls.Add(tela);
        }

        // =========================================================
        // TELA DE ESPERA
        // =========================================================

        private Panel CriarTelaEsperaProcessamento(out Label lblStatus)
        {
            Panel tela = CriarContainer(650);
            Label titulo = CriarTitulo("Processando cotações...");
            titulo.Location = new Point(10, 10);

            Label descricao = CriarTexto("Aguarde enquanto a IA analisa os PDFs e preenche a cotação.");
            descricao.Location = new Point(10, 50);

            ProgressBar progresso = new()
            {
                Location = new Point(10, 90),
                Width = 800,
                Height = 20,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };
            Panel card = CriarCard(800, 230);
            card.Location = new Point(10, 135);
            lblStatus = new Label
            {
                Text = "Preparando arquivos...",
                AutoSize = false,
                Width = 750,
                Height = 180,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 11F),
                Location = new Point(20, 20)
            };
            card.Controls.Add(lblStatus);
            tela.Controls.Add(titulo);
            tela.Controls.Add(descricao);
            tela.Controls.Add(progresso);
            tela.Controls.Add(card);
            return tela;
        }

        private void PrepararNovaCotacao()
        {
            pdfsSelecionados.Clear();
            caminhoExcelSelecionado = "";

            lojaSelecionada = "";

            mesSelecionado =
                DateTime.Now.Month;

            anoSelecionado =
                DateTime.Now.Year;

        }

        private void ExecutarExportacaoPdf(string tipo, string empresa, int mes, int ano,
            Func<(string[] Cabecalhos, List<string[]> Linhas)> carregar)
        {
            try
            {
                var dados = carregar();
                if (dados.Linhas.Count == 0) { MessageBox.Show("Não há dados para exportar neste período.", "Exportar PDF"); return; }
                using var salvar = new SaveFileDialog { Filter = "Arquivo PDF (*.pdf)|*.pdf", DefaultExt = "pdf",
                    AddExtension = true, FileName = $"{tipo}_{empresa}_{ano}-{mes:00}.pdf", Title = "Salvar PDF para impressão" };
                if (salvar.ShowDialog(this) != DialogResult.OK) return;
                string caminhoPdf = salvar.FileName;
                byte[]? pdfAnterior = File.Exists(caminhoPdf) ? File.ReadAllBytes(caminhoPdf) : null;
                RelatorioPdf.Salvar(salvar.FileName, tipo == "Cotacao" ? "Cotação de insumos" : "Insumos não encontrados",
                    $"{empresa} • {ObterNomeMes(mes)}/{ano}", dados.Cabecalhos, dados.Linhas, todasColunas: tipo == "Cotacao");
                byte[] pdfDepois = File.ReadAllBytes(caminhoPdf);
                RegistrarDesfazer(() =>
                {
                    if (pdfAnterior == null) { if (File.Exists(caminhoPdf)) File.Delete(caminhoPdf); }
                    else File.WriteAllBytes(caminhoPdf, pdfAnterior);
                }, () => File.WriteAllBytes(caminhoPdf, pdfDepois));
                MessageBox.Show("PDF exportado com sucesso.\n\n" + salvar.FileName, "Exportar PDF");
            }
            catch (Exception ex) { MessageBox.Show("Não foi possível exportar o PDF.\n\n" + ex.Message,
                "Exportar PDF", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void AbrirTelaConsultar()
        {
            telaAtual = TelaAtual.Consultar;
            LimparConteudo();
            DestacarMenu(btnConsultar);
            Panel tela = CriarContainer(740);
            Label titulo = CriarTitulo("Consultar cotações");
            titulo.Location = new Point(10, 5);
            Button orlando = CriarBotao("Orlando", 130, 40);
            orlando.Location = new Point(10, 65);
            Button drugstore = CriarBotao("Drugstore", 130, 40);
            drugstore.Location = new Point(150, 65);
            if (lojaConsultaSelecionada == "Orlando")
                orlando.BackColor = corAzul;
            else
                drugstore.BackColor = corAzul;
            orlando.Click += (s, e) =>
            {
                lojaConsultaSelecionada = "Orlando";
                AbrirTelaConsultar();
            };
            drugstore.Click += (s, e) =>
            {
                lojaConsultaSelecionada = "Drugstore";
                AbrirTelaConsultar();
            };

            Label lblAno = CriarLabelSecao("Ano");
            lblAno.Location = new Point(10, 120);
            NumericUpDown campoAno = new NumericUpDown
            {
                Location = new Point(10, 148), Width = 110,
                Minimum = 1900, Maximum = 9999, Value = anoFiltroConsulta,
                BackColor = Color.FromArgb(25, 25, 25), ForeColor = Color.White,
                Font = new Font("Segoe UI", 11F)
            };
            Button buscar = CriarBotao("Buscar", 130, 36);
            buscar.Location = new Point(140, 145);
            buscar.BackColor = corAzul;
            DataGridView dgv = CriarGrid();
            dgv.Location = new Point(10, 200);
            dgv.Size = new Size(960, 350);
            dgv.Columns.Add("Mes", "Mês");
            dgv.Columns.Add("Ano", "Ano");
            dgv.Columns.Add("Arquivo", "Arquivo");
            dgv.Columns.Add("Modificado", "Última modificação");
            dgv.Columns.Add(new DataGridViewButtonColumn
            {
                Name = "AbrirTabela", HeaderText = "Tabela", Text = "Abrir tabela",
                UseColumnTextForButtonValue = true, Width = 110,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dgv.Columns.Add(new DataGridViewButtonColumn
            {
                Name = "NaoEncontrados", HeaderText = "Não encontrados",
                Text = "Visualizar", UseColumnTextForButtonValue = true
            });
            dgv.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Selecionar", HeaderText = "Selecionar", Width = 90,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            foreach (DataGridViewColumn coluna in dgv.Columns)
            {
                coluna.ReadOnly = coluna.Name != "Selecionar";
                coluna.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dgv.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgv.IsCurrentCellDirty)
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            Label vazio = CriarTexto("");
            vazio.Location = new Point(10, 620);
            Button selecionarTudo = CriarBotao("Selecionar tudo", 160, 40);
            selecionarTudo.Location = new Point(10, 565);
            Button excluir = CriarBotao("Excluir tabela", 160, 40);
            excluir.Location = new Point(185, 565);

            Button exportar = CriarBotao("Exportar PDF", 140, 36);
            exportar.BackColor = corAzul;
            exportar.Click += (s, e) =>
            {
                var meses = dgv.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow)
                    .Select(r => Convert.ToString(r.Cells["Mes"].Value) ?? "").ToArray();
                if (meses.Length == 0) { MessageBox.Show("Nenhuma cotação disponível para exportar.", "Exportar PDF"); return; }
                using var escolha = new Form { Text = "Exportar cotação em PDF", Size = new Size(390, 190),
                    StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false, MinimizeBox = false };
                var label = new Label { Text = $"{lojaConsultaSelecionada} • {anoFiltroConsulta} — escolha o mês:",
                    AutoSize = true, Location = new Point(15, 15) };
                var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = new Point(15, 45), Width = 340 };
                combo.Items.AddRange(meses); combo.SelectedIndex = 0;
                var confirmar = new Button { Text = "Exportar PDF", Location = new Point(225, 85),
                    Size = new Size(130, 35), DialogResult = DialogResult.OK };
                escolha.Controls.AddRange(new Control[] { label, combo, confirmar }); escolha.AcceptButton = confirmar;
                if (escolha.ShowDialog(this) != DialogResult.OK) return;
                int mes = ObterNumeroMes(combo.Text);
                ExecutarExportacaoPdf("Cotacao", lojaConsultaSelecionada, mes, anoFiltroConsulta, () =>
                {
                    string pasta = ObterPastaPeriodo(lojaConsultaSelecionada, mes, anoFiltroConsulta);
                    string? caminho = Directory.GetFiles(pasta, "*.xlsx")
                        .Where(f => !Path.GetFileName(f).StartsWith("~$") && !Path.GetFileName(f).StartsWith("temp_"))
                        .OrderBy(f => f).FirstOrDefault();
                    if (caminho == null) throw new InvalidOperationException("A planilha deste mês ainda não foi gerada.");
                    using var workbook = ArquivoPlanilhaExcel.Abrir(caminho);
                    var planilha = workbook.Worksheets.First();
                    var range = planilha.RangeUsed() ?? throw new InvalidOperationException("A planilha está vazia.");
                    int primeira = range.FirstRow().RowNumber(), ultima = range.LastRow().RowNumber();
                    int primeiraColuna = range.FirstColumn().ColumnNumber(), quantidade = range.ColumnCount();
                    string[] Ler(int linha) => Enumerable.Range(primeiraColuna, quantidade)
                        .Select(c => planilha.Cell(linha, c).GetFormattedString()).ToArray();
                    return (Ler(primeira), Enumerable.Range(primeira + 1, ultima - primeira)
                        .Select(Ler).Where(r => r.Any(v => !string.IsNullOrWhiteSpace(v))).ToList());
                });
            };

            void CarregarResultados()
            {
                dgv.Rows.Clear();
                string pastaEmpresa = ObterPastaEmpresa(lojaConsultaSelecionada);
                try
                {
                    if (Directory.Exists(pastaEmpresa))
                    {
                        foreach (string pastaPeriodo in Directory.GetDirectories(pastaEmpresa)
                            .OrderByDescending(pasta => Path.GetFileName(pasta)))
                        {
                            string[] partes = Path.GetFileName(pastaPeriodo).Split('-');
                            if (partes.Length != 2
                                || !int.TryParse(partes[0], out int ano)
                                || ano != anoFiltroConsulta
                                || !int.TryParse(partes[1], out int mes)
                                || mes < 1 || mes > 12)
                                continue;
                        string arquivo = Directory.GetFiles(pastaPeriodo, "*.xlsx")
                            .Where(caminho => !Path.GetFileName(caminho).StartsWith("~$")
                                && !Path.GetFileName(caminho).StartsWith("temp_"))
                            .OrderBy(caminho => caminho).Select(Path.GetFileName).FirstOrDefault()
                            ?? "Planilha ainda não gerada";
                        int indice = dgv.Rows.Add(ObterNomeMes(mes), ano,
                            arquivo, Directory.GetLastWriteTime(pastaPeriodo).ToString("dd/MM/yyyy HH:mm"),
                            "Abrir tabela", "Visualizar", false);
                        dgv.Rows[indice].Tag = pastaPeriodo;
                        }
                    }
                    vazio.Text = dgv.Rows.Count == 0
                        ? "Nenhuma cotação encontrada para esta empresa e ano." : "";
                }
                catch (Exception ex)
                {
                    vazio.Text = "Não foi possível consultar as cotações.";
                    MessageBox.Show("Não foi possível buscar as tabelas.\n\n" + ex.Message,
                        "Buscar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                selecionarTudo.Enabled = dgv.Rows.Count > 0;
                excluir.Enabled = dgv.Rows.Count > 0;
            }
            buscar.Click += (s, e) =>
            {
                anoFiltroConsulta = (int)campoAno.Value;
                CarregarResultados();
            };
            selecionarTudo.Click += (s, e) =>
            {
                dgv.EndEdit();
                foreach (DataGridViewRow linha in dgv.Rows)
                    linha.Cells["Selecionar"].Value = true;
            };
            excluir.Click += (s, e) =>
            {
                dgv.EndEdit();
                var selecionadas = dgv.Rows.Cast<DataGridViewRow>()
                    .Where(linha => linha.Cells["Selecionar"].Value is true).ToArray();
                if (selecionadas.Length == 0)
                {
                    MessageBox.Show("Selecione uma tabela para excluir.", "Excluir tabela",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string periodos = string.Join("\n", selecionadas.Select(linha =>
                    $"{lojaConsultaSelecionada} - {linha.Cells["Mes"].Value}/{linha.Cells["Ano"].Value}"));
                if (MessageBox.Show("Enviar os períodos selecionados para a Lixeira?\n\n" + periodos +
                    "\n\nSerão removidos a planilha e os registros de não encontrados dessas pastas.",
                    "Excluir tabela", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                    return;
                var erros = new List<string>();
                var periodosRemovidos = new List<(string Pasta, Dictionary<string, byte[]> Arquivos, List<ItemNaoEncontradoSalvo> Itens)>();
                string empresaExcluida = lojaConsultaSelecionada;
                int anoExcluido = anoFiltroConsulta;
                string raizEmpresa = Path.GetFullPath(ObterPastaEmpresa(lojaConsultaSelecionada));
                foreach (DataGridViewRow linha in selecionadas)
                {
                    try
                    {
                        string destino = Path.GetFullPath((string)linha.Tag!);
                        if (!string.Equals(Path.GetDirectoryName(destino), raizEmpresa,
                                StringComparison.OrdinalIgnoreCase)
                            || Path.GetFileName(destino) != $"{Convert.ToInt32(linha.Cells["Ano"].Value)}-{ObterNumeroMes(Convert.ToString(linha.Cells["Mes"].Value) ?? ""):00}")
                            throw new InvalidOperationException("Pasta do período inválida.");
                        var arquivosAntes = Directory.Exists(destino) ? Directory.GetFiles(destino, "*", SearchOption.AllDirectories)
                            .ToDictionary(f => Path.GetRelativePath(destino, f), File.ReadAllBytes) : new Dictionary<string, byte[]>();
                        var itensAntes = naoEncontradosRepository.Consultar(empresaExcluida,
                            ObterNumeroMes(Convert.ToString(linha.Cells["Mes"].Value) ?? ""), anoExcluido);
                        if (Directory.Exists(destino))
                            Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(destino,
                                Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                                Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin,
                                Microsoft.VisualBasic.FileIO.UICancelOption.ThrowException);
                        naoEncontradosRepository.ExcluirPeriodo(lojaConsultaSelecionada,
                            ObterNumeroMes(Convert.ToString(linha.Cells["Mes"].Value) ?? ""),
                            Convert.ToInt32(linha.Cells["Ano"].Value));
                        periodosRemovidos.Add((destino, arquivosAntes, itensAntes));
                    }
                    catch (Exception ex)
                    {
                        erros.Add($"{linha.Cells["Mes"].Value}/{linha.Cells["Ano"].Value}: {ex.Message}");
                    }
                }
                if (periodosRemovidos.Count > 0) RegistrarDesfazer(() =>
                {
                    if (periodosRemovidos.Any(p => Directory.Exists(p.Pasta)))
                        throw new IOException("Já existe uma nova cotação no período excluído. Não será sobrescrita.");
                    var criadas = new List<string>();
                    try
                    {
                        foreach (var periodo in periodosRemovidos)
                        {
                            Directory.CreateDirectory(periodo.Pasta); criadas.Add(periodo.Pasta);
                            foreach (var arquivo in periodo.Arquivos)
                            {
                                string destino = Path.Combine(periodo.Pasta, arquivo.Key);
                                Directory.CreateDirectory(Path.GetDirectoryName(destino)!);
                                File.WriteAllBytes(destino, arquivo.Value);
                            }
                        }
                        naoEncontradosRepository.Restaurar(periodosRemovidos.SelectMany(p => p.Itens));
                    }
                    catch
                    {
                        foreach (var pastaCriada in criadas) Directory.Delete(pastaCriada, true);
                        throw;
                    }
                    lojaConsultaSelecionada = empresaExcluida; anoFiltroConsulta = anoExcluido;
                    AbrirTelaConsultar();
                }, () =>
                {
                    foreach (var periodo in periodosRemovidos)
                    {
                        if (Directory.Exists(periodo.Pasta))
                            Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(periodo.Pasta,
                                Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                                Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                        naoEncontradosRepository.Excluir(periodo.Itens.Select(i => i.Id));
                    }
                    lojaConsultaSelecionada = empresaExcluida; anoFiltroConsulta = anoExcluido;
                    AbrirTelaConsultar();
                });
                CarregarResultados();
                if (erros.Count > 0)
                    MessageBox.Show("Não foi possível excluir:\n\n" + string.Join("\n", erros),
                        "Excluir tabela", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            dgv.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0)
                    return;
                string coluna = dgv.Columns[e.ColumnIndex].Name;
                DataGridViewRow linha = dgv.Rows[e.RowIndex];
                int mes = ObterNumeroMes(Convert.ToString(linha.Cells["Mes"].Value) ?? "");
                int ano = Convert.ToInt32(linha.Cells["Ano"].Value);
                if (coluna == "AbrirTabela")
                    AbrirTabelaCotacao(lojaConsultaSelecionada, mes, ano);
                else if (coluna == "NaoEncontrados")
                {
                    lojaSelecionada = lojaConsultaSelecionada;
                    mesSelecionado = mes;
                    anoSelecionado = ano;
                    AbrirTelaNaoEncontrados();
                }
            };
            AjustarContainerATela(tela);
            EstilizarGrade(dgv);
            dgv.Columns["Arquivo"]!.FillWeight = 190;
            dgv.Columns["Mes"]!.FillWeight = 75;
            dgv.Columns["Ano"]!.FillWeight = 45;
            Panel filtros = CriarCard(900, 110);
            EstilizarCard(filtros);
            Label empresaFiltro = CriarLabelSecao("Empresa");
            empresaFiltro.Location = new Point(18, 15);
            Label subtituloConsulta = CriarTexto("Consulte as planilhas por empresa e período, abra os detalhes ou exporte em PDF.");
            subtituloConsulta.Location = new Point(10, 45);
            orlando.Location = new Point(18, 45);
            drugstore.Location = new Point(160, 45);
            lblAno.Location = new Point(320, 15);
            campoAno.Location = new Point(320, 48);
            buscar.Location = new Point(450, 45);
            buscar.Size = new Size(130, 40);
            buscar.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            exportar.Size = new Size(148, 40);
            exportar.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            excluir.ForeColor = Color.FromArgb(246, 143, 143);
            excluir.BackColor = Color.FromArgb(43, 27, 32);
            excluir.FlatAppearance.BorderColor = Color.FromArgb(105, 52, 58);
            Label resultados = CriarLabelSecao("COTAÇÕES DO PERÍODO");
            resultados.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            resultados.ForeColor = Color.FromArgb(158, 173, 192);
            void AjustarLayoutConsulta()
            {
                int largura = Math.Max(600, tela.ClientSize.Width - 20);
                filtros.SetBounds(10, 88, largura, 108);
                resultados.Location = new Point(10, filtros.Bottom + 26);
                exportar.Location = new Point(filtros.Right - exportar.Width, filtros.Bottom + 14);
                dgv.SetBounds(10, exportar.Bottom + 16, largura,
                    Math.Max(80, tela.ClientSize.Height - exportar.Bottom - 94));
                excluir.Location = new Point(dgv.Right - excluir.Width, dgv.Bottom + 16);
                selecionarTudo.Location = new Point(excluir.Left - selecionarTudo.Width - 12, excluir.Top);
                vazio.Location = new Point(10, excluir.Top + 10);
            }
            tela.Resize += (s, e) => AjustarLayoutConsulta();
            AjustarLayoutConsulta();
            tela.Controls.AddRange([titulo, subtituloConsulta, filtros, resultados,
                exportar, dgv, selecionarTudo, excluir, vazio]);
            filtros.Controls.AddRange([empresaFiltro, orlando, drugstore, lblAno, campoAno, buscar]);
            painelConteudo.Controls.Add(tela);
            CarregarResultados();
        }
        private int ObterNumeroMes(
    string mes)
        {
            switch (mes)
            {
                case "Janeiro":
                    return 1;

                case "Fevereiro":
                    return 2;

                case "Março":
                    return 3;

                case "Abril":
                    return 4;

                case "Maio":
                    return 5;

                case "Junho":
                    return 6;

                case "Julho":
                    return 7;

                case "Agosto":
                    return 8;

                case "Setembro":
                    return 9;

                case "Outubro":
                    return 10;

                case "Novembro":
                    return 11;

                case "Dezembro":
                    return 12;

                default:
                    return 0;
            }
        }

        private string ObterNomeMes(int mes)
        {
            string[] meses =
            {
                "",
                "Janeiro",
                "Fevereiro",
                "Março",
                "Abril",
                "Maio",
                "Junho",
                "Julho",
                "Agosto",
                "Setembro",
                "Outubro",
                "Novembro",
                "Dezembro"
            };

            if (
                mes < 1 ||
                mes > 12
            )
                return "";

            return meses[mes];
        }

        // =========================================================
        // NÃO ENCONTRADOS
        // =========================================================

        private void AbrirTelaNaoEncontrados()
        {
            telaAtual = TelaAtual.NaoEncontrados;

            LimparConteudo();
            DestacarMenu(btnNaoEncontrados);

            Panel tela = CriarContainer(760);

            Label titulo = CriarTitulo("Insumos não encontrados");
            titulo.Location = new Point(10, 5);

            Label subtitulo = CriarTexto(
                "Consulte os insumos que não foram localizados automaticamente na planilha."
            );

            subtitulo.Location = new Point(10, 45);

            // =========================================================
            // FILTROS
            // =========================================================

            Label lblEmpresa = CriarLabelSecao("Empresa");
            lblEmpresa.Location = new Point(10, 95);

            ComboBox cmbEmpresa = CriarCombo(220);
            cmbEmpresa.Location = new Point(10, 125);

            cmbEmpresa.Items.AddRange(
                new object[]
                {
            "Orlando",
            "Drugstore"
                }
            );

            if (
                !string.IsNullOrWhiteSpace(lojaConsultaSelecionada) &&
                cmbEmpresa.Items.Contains(lojaConsultaSelecionada)
            )
            {
                cmbEmpresa.SelectedItem =
                    lojaConsultaSelecionada;
            }
            else
            {
                cmbEmpresa.SelectedIndex = 0;
            }

            Label lblMes = CriarLabelSecao("Mês");
            lblMes.Location = new Point(250, 95);

            ComboBox cmbMes = CriarCombo(180);
            cmbMes.Location = new Point(250, 125);

            for (int i = 1; i <= 12; i++)
            {
                cmbMes.Items.Add(
                    ObterNomeMes(i)
                );
            }

            cmbMes.SelectedIndex =
                Math.Max(
                    0,
                    Math.Min(
                        mesSelecionado - 1,
                        11
                    )
                );

            Label lblAno = CriarLabelSecao("Ano");
            lblAno.Location = new Point(450, 95);

            NumericUpDown numAno =
                CriarAno(120);

            numAno.Location =
                new Point(450, 125);

            numAno.Value =
                anoSelecionado;

            Button btnBuscar =
                CriarBotaoAzul(
                    "Buscar",
                    120,
                    36
                );

            btnBuscar.Location =
                new Point(610, 123);

            // =========================================================
            // GRID
            // =========================================================

            Panel fornecedores = new Panel
            {
                Location = new Point(10, 190), Width = 870, Height = 430,
                AutoScroll = true
            };
            Label quantidade = CriarLabelSecao("Fornecedores");
            quantidade.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            quantidade.Location = new Point(10, 222);
            Label orientacao = CriarTexto("Abra os fornecedores que deseja incluir no PDF.");
            orientacao.Location = new Point(10, 252);
            orientacao.ForeColor = Color.FromArgb(158, 173, 192);
            var grupos = new List<(Panel Painel, Button Cabecalho, DataGridView Grid,
                List<ItemNaoEncontradoSalvo> Itens)>();
            string empresaCarregada = "";
            int mesCarregado = 0, anoCarregado = 0;

            void OrganizarFornecedores()
            {
                int y = fornecedores.AutoScrollPosition.Y;
                foreach (var grupo in grupos)
                {
                    grupo.Painel.SetBounds(0, y, Math.Max(1, fornecedores.ClientSize.Width -
                        SystemInformation.VerticalScrollBarWidth - 2),
                        grupo.Grid.Visible ? grupo.Grid.Height + 58 : 58);
                    y += grupo.Painel.Height + 12;
                }
            }
            fornecedores.Resize += (s, e) => OrganizarFornecedores();

            DataGridView CriarGridFornecedor()
            {
                DataGridView dgv =
                    CriarGrid();

                dgv.Dock = DockStyle.Fill;

                dgv.Width =
                    870;

                dgv.Height =
                    430;

                dgv.AllowUserToAddRows =
                    false;

                dgv.AllowUserToDeleteRows =
                    false;

                dgv.ReadOnly =
                    false;

                dgv.SelectionMode =
                    DataGridViewSelectionMode.FullRowSelect;

                dgv.MultiSelect =
                    true;

                dgv.ScrollBars =
                    ScrollBars.Vertical;

                dgv.AutoSizeColumnsMode =
                    DataGridViewAutoSizeColumnsMode.None;

                // =========================================================
                // COLUNAS
                // =========================================================

                dgv.Columns.Add(
                    "Fornecedor",
                    "Fornecedor"
                );

                dgv.Columns["Fornecedor"]!.Width =
                    170;

                dgv.Columns.Add(
                    "Insumo",
                    "Insumo"
                );

                dgv.Columns["Insumo"]!.Width =
                    300;

                dgv.Columns.Add(
                    "Preco",
                    "Preço normalizado"
                );

                dgv.Columns["Preco"]!.Width =
                    150;

                dgv.Columns.Add(
                    "Tipo",
                    "Tipo"
                );

                dgv.Columns["Tipo"]!.Width =
                    110;

                dgv.Columns.Add(
                    "Data",
                    "Data"
                );

                dgv.Columns["Data"]!.Width =
                    90;

                DataGridViewCheckBoxColumn colunaSelecionar =
                    new DataGridViewCheckBoxColumn
                    {
                        Name =
                            "Selecionar",

                        HeaderText =
                            "",

                        Width =
                            38,

                        MinimumWidth =
                            38,

                        AutoSizeMode =
                            DataGridViewAutoSizeColumnMode.None
                    };

                dgv.Columns.Add(
                    colunaSelecionar
                );

                bool alterandoSelecao = false;
                CheckState EstadoSelecao()
                {
                    int marcados = dgv.Rows.Cast<DataGridViewRow>()
                        .Count(linha => Equals(linha.Cells["Selecionar"].Value, true));
                    return marcados == 0 ? CheckState.Unchecked
                        : marcados == dgv.Rows.Count ? CheckState.Checked
                        : CheckState.Indeterminate;
                }

                dgv.CellPainting += (s, e) =>
                {
                    if (e.RowIndex != -1 || e.ColumnIndex != colunaSelecionar.Index) return;
                    e.PaintBackground(e.ClipBounds, true);
                    e.Paint(e.ClipBounds, DataGridViewPaintParts.Border);
                    CheckState estado = EstadoSelecao();
                    var estadoVisual = estado == CheckState.Checked
                        ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal
                        : estado == CheckState.Indeterminate
                            ? System.Windows.Forms.VisualStyles.CheckBoxState.MixedNormal
                            : System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal;
                    Size tamanho = CheckBoxRenderer.GetGlyphSize(e.Graphics!, estadoVisual);
                    var posicao = new Point(e.CellBounds.Left + (e.CellBounds.Width - tamanho.Width) / 2,
                        e.CellBounds.Top + (e.CellBounds.Height - tamanho.Height) / 2);
                    CheckBoxRenderer.DrawCheckBox(e.Graphics!, posicao, estadoVisual);
                    e.Handled = true;
                };
                dgv.ColumnHeaderMouseClick += (s, e) =>
                {
                    if (e.ColumnIndex != colunaSelecionar.Index || e.Button != MouseButtons.Left) return;
                    dgv.EndEdit();
                    bool selecionar = EstadoSelecao() != CheckState.Checked;
                    alterandoSelecao = true;
                    try
                    {
                        foreach (DataGridViewRow linha in dgv.Rows)
                            linha.Cells["Selecionar"].Value = selecionar;
                    }
                    finally
                    {
                        alterandoSelecao = false;
                        dgv.InvalidateCell(colunaSelecionar.Index, -1);
                    }
                };
                dgv.CurrentCellDirtyStateChanged += (s, e) =>
                {
                    if (dgv.IsCurrentCellDirty && dgv.CurrentCell?.OwningColumn == colunaSelecionar)
                        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
                };
                dgv.CellValueChanged += (s, e) =>
                {
                    if (!alterandoSelecao && e.ColumnIndex == colunaSelecionar.Index)
                        dgv.InvalidateCell(colunaSelecionar.Index, -1);
                };
                colunaSelecionar.HeaderCell.ToolTipText = "Selecionar ou desmarcar todos os itens deste fornecedor";

                // Somente checkbox editável
                foreach (
                    DataGridViewColumn coluna
                    in dgv.Columns
                )
                {
                    coluna.ReadOnly =
                        coluna.Name != "Selecionar";
                }

                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.Columns["Insumo"]!.FillWeight = 180;
                dgv.Columns["Fornecedor"]!.FillWeight = 110;
                dgv.Columns["Preco"]!.FillWeight = 95;
                dgv.Columns["Tipo"]!.FillWeight = 70;
                dgv.Columns["Data"]!.FillWeight = 65;
            return dgv;
            }

            // =========================================================
            // BOTÃO EXCLUIR
            // =========================================================

            Button btnExcluir =
                CriarBotao(
                    "Excluir selecionados",
                    180,
                    38
                );

            btnExcluir.Location =
                new Point(700, 640);

            // =========================================================
            // CARREGAR ITENS
            // =========================================================

            void CarregarItens()
            {
                var abertos = grupos.Where(g => g.Grid.Visible)
                    .Select(g => g.Itens[0].Fornecedor).ToHashSet(StringComparer.CurrentCultureIgnoreCase);
                foreach (var grupo in grupos) grupo.Painel.Dispose();
                grupos.Clear();
                fornecedores.Controls.Clear();
                fornecedores.AutoScrollPosition = Point.Empty;

                string empresa =
                    cmbEmpresa.SelectedItem?
                        .ToString()
                    ?? "";

                int mes =
                    cmbMes.SelectedIndex + 1;

                int ano =
                    (int)numAno.Value;

                lojaConsultaSelecionada =
                    empresa;

                try
                {
                    var itens = naoEncontradosRepository.Consultar(empresa, mes, ano);
                    empresaCarregada = empresa;
                    mesCarregado = mes;
                    anoCarregado = ano;
                    foreach (var fornecedor in itens
                        .GroupBy(i => i.Fornecedor, StringComparer.CurrentCultureIgnoreCase)
                        .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase))
                    {
                        var ordenados = fornecedor.OrderBy(i => i.Insumo,
                            StringComparer.Create(System.Globalization.CultureInfo.GetCultureInfo("pt-BR"), true))
                            .ThenBy(i => i.Id).ToList();
                        var dgv = CriarGridFornecedor();
                        foreach (var item in ordenados)
                        {
                            int indice = dgv.Rows.Add(item.Fornecedor, item.Insumo,
                                item.PrecoNormalizado > 0 ? item.PrecoNormalizado.ToString("0.############################") : "Conferir PDF",
                                item.TipoPreco == "NAO_IDENTIFICADO" || string.IsNullOrWhiteSpace(item.TipoPreco) ? "Conferir unidade" : item.TipoPreco,
                                item.DataRegistro.ToString("dd/MM"), false);
                            dgv.Rows[indice].Tag = item.Id;
                        }
                        // Keep the initial alphabetical order when displaying each supplier.
                        foreach (DataGridViewColumn coluna in dgv.Columns)
                            coluna.SortMode = DataGridViewColumnSortMode.NotSortable;
                        dgv.Dock = DockStyle.None;
                        dgv.Location = new Point(0, 58);
                        dgv.Height = Math.Min(430, dgv.ColumnHeadersHeight + dgv.Rows.Count * dgv.RowTemplate.Height + 4);
                        dgv.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                        var painel = new Panel();
                        var cabecalho = CriarBotao("", 870, 54);
                        cabecalho.Dock = DockStyle.Top;
                        cabecalho.TextAlign = ContentAlignment.MiddleLeft;
                        cabecalho.Padding = new Padding(16, 0, 12, 0);
                        cabecalho.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
                        cabecalho.BackColor = Color.FromArgb(26, 32, 41);
                        cabecalho.FlatAppearance.BorderColor = Color.FromArgb(48, 61, 78);
                        cabecalho.FlatAppearance.MouseOverBackColor = Color.FromArgb(34, 46, 62);
                        dgv.BackgroundColor = Color.FromArgb(20, 24, 30);
                        dgv.BorderStyle = BorderStyle.None;
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(32, 44, 60);
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.WhiteSmoke;
                        dgv.ColumnHeadersHeight = 36;
                        dgv.DefaultCellStyle.Padding = new Padding(6, 3, 6, 3);
                        dgv.RowTemplate.Height = 32;
                        foreach (DataGridViewRow linha in dgv.Rows) linha.Height = 32;
                        dgv.Height = Math.Min(430, 40 + dgv.Rows.Count * 32);
                        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(26, 30, 37);
                        bool aberto = abertos.Contains(fornecedor.Key);
                        void AtualizarCabecalho() => cabecalho.Text =
                            $"{(dgv.Visible ? "▼" : "▶")}  {fornecedor.Key}   •   {ordenados.Count} insumos";
                        painel.Controls.Add(dgv);
                        painel.Controls.Add(cabecalho);
                        fornecedores.Controls.Add(painel);
                        grupos.Add((painel, cabecalho, dgv, ordenados));
                        painel.Resize += (s, e) => dgv.Width = painel.ClientSize.Width;
                        dgv.Visible = aberto;
                        AtualizarCabecalho();
                        cabecalho.Click += (s, e) =>
                        {
                            dgv.Visible = !dgv.Visible;
                            AtualizarCabecalho();
                            OrganizarFornecedores();
                        };
                    }
                    quantidade.Text = $"Fornecedores ({grupos.Count})  •  {itens.Count} insumos";
                    orientacao.Text = itens.Count == 0
                        ? "Nenhum insumo não encontrado neste período."
                        : "Abra os fornecedores que deseja incluir no PDF.";
                    OrganizarFornecedores();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Não foi possível consultar os insumos.\n\n" + ex.Message,
                        "Erro na consulta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // =========================================================
            // BUSCAR
            // =========================================================

            btnBuscar.Click +=
                (s, e) =>
                {
                    CarregarItens();
                };

            // =========================================================
            // EXCLUIR
            // =========================================================

            btnExcluir.Click +=
                (s, e) =>
                {
                    foreach (var grupo in grupos) grupo.Grid.EndEdit();
                    var itensParaExcluir = grupos.SelectMany(g => g.Grid.Rows.Cast<DataGridViewRow>())
                        .Where(linha => Convert.ToBoolean(linha.Cells["Selecionar"].Value ?? false))
                        .Select(linha => (long)linha.Tag!).ToList();
                    if (itensParaExcluir.Count == 0)
                    {
                        MessageBox.Show(
                            "Selecione pelo menos um item para excluir.",
                            "Atenção",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    DialogResult resposta =
                        MessageBox.Show(
                            $"Deseja excluir {itensParaExcluir.Count} item(ns) selecionado(s)?",
                            "Confirmar exclusão",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                    if (
                        resposta !=
                        DialogResult.Yes
                    )
                    {
                        return;
                    }

                    try
                    {
                        string empresaExcluida = empresaCarregada;
                        int mesExcluido = mesCarregado, anoExcluido = anoCarregado;
                        var removidos = naoEncontradosRepository.Consultar(empresaExcluida, mesExcluido, anoExcluido)
                            .Where(i => itensParaExcluir.Contains(i.Id)).ToList();
                        naoEncontradosRepository.Excluir(itensParaExcluir);
                        RegistrarDesfazer(() =>
                        {
                            naoEncontradosRepository.Restaurar(removidos);
                            lojaSelecionada = empresaExcluida; mesSelecionado = mesExcluido; anoSelecionado = anoExcluido;
                            AbrirTelaNaoEncontrados();
                        }, () =>
                        {
                            naoEncontradosRepository.Excluir(removidos.Select(i => i.Id));
                            lojaSelecionada = empresaExcluida; mesSelecionado = mesExcluido; anoSelecionado = anoExcluido;
                            AbrirTelaNaoEncontrados();
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Não foi possível excluir os insumos.\n\n" + ex.Message,
                            "Erro na exclusão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    CarregarItens();
                };

            // =========================================================
            // EMPRESA
            // =========================================================

            cmbEmpresa.SelectedIndexChanged +=
                (s, e) =>
                {
                    if (
                        cmbEmpresa.SelectedItem != null
                    )
                    {
                        lojaConsultaSelecionada =
                            cmbEmpresa.SelectedItem
                                .ToString()
                            ?? "";
                    }
                };

            // =========================================================
            // CONTROLES
            // =========================================================

            AjustarContainerATela(tela);
            Button exportar = CriarBotao("Exportar PDF", 140, 36);
            exportar.BackColor = corAzul;
            exportar.Click += (s, e) =>
            {
                if (!grupos.Any(g => g.Grid.Visible))
                {
                    MessageBox.Show("Abra pelo menos um fornecedor para exportar sua tabela.", "Exportar PDF");
                    return;
                }
                string empresa = empresaCarregada;
                int mes = mesCarregado, ano = anoCarregado;
                ExecutarExportacaoPdf("Nao_encontrados", empresa, mes, ano, () =>
                {
                    var itens = grupos.Where(g => g.Grid.Visible).SelectMany(g => g.Itens);
                    return (new[] { "Fornecedor", "Insumo", "Preço normalizado", "Tipo", "Data" },
                        itens.Select(i => new[] { i.Fornecedor, i.Insumo,
                            i.PrecoNormalizado > 0 ? i.PrecoNormalizado.ToString("0.############################") : "Conferir PDF",
                            i.TipoPreco == "NAO_IDENTIFICADO" || string.IsNullOrWhiteSpace(i.TipoPreco) ? "Conferir unidade" : i.TipoPreco, i.DataRegistro.ToString("dd/MM/yyyy") }).ToList());
                });
            };
            Button abrirPlanilha = CriarBotao("Abrir planilha do período", 220, 38);
            abrirPlanilha.BackColor = Color.FromArgb(35, 59, 87);
            abrirPlanilha.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            abrirPlanilha.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(empresaCarregada)) return;
                lojaConsultaSelecionada = empresaCarregada;
                mesSelecionado = mesCarregado;
                anoSelecionado = anoCarregado;
                anoFiltroConsulta = anoCarregado;
                AbrirTabelaCotacao(empresaCarregada, mesCarregado, anoCarregado);
            };
            tela.Controls.Add(abrirPlanilha);
            Panel filtros = CriarCard(850, 108);
            filtros.BorderStyle = BorderStyle.None;
            filtros.BackColor = Color.FromArgb(23, 28, 35);
            filtros.Location = new Point(10, 94);
            tela.Controls.AddRange([filtros, quantidade, orientacao, exportar]);
            exportar.Size = new Size(148, 40);
            exportar.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            btnExcluir.ForeColor = Color.FromArgb(246, 143, 143);
            btnExcluir.FlatAppearance.BorderColor = Color.FromArgb(105, 52, 58);
            btnExcluir.BackColor = Color.FromArgb(43, 27, 32);
            btnBuscar.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            void AjustarLayoutNaoEncontrados()
            {
                int largura = Math.Max(340, tela.ClientSize.Width - 20);
                abrirPlanilha.Location = new Point(Math.Max(10,
                    tela.ClientSize.Width - abrirPlanilha.Width - 10), 5);
                filtros.Width = largura;
                bool compacto = largura < 780;
                filtros.Height = compacto ? 160 : 108;
                int coluna = Math.Max(100, (largura - 320) / 3);
                lblEmpresa.Location = new Point(18, 14);
                cmbEmpresa.SetBounds(18, 44, coluna, 30);
                lblMes.Location = new Point(cmbEmpresa.Right + 18, 14);
                cmbMes.SetBounds(lblMes.Left, 44, coluna, 30);
                lblAno.Location = new Point(cmbMes.Right + 18, 14);
                numAno.SetBounds(lblAno.Left, 44, compacto ? 90 : 110, 30);
                btnBuscar.Location = compacto ? new Point(18, 102) : new Point(numAno.Right + 18, 39);
                quantidade.Location = new Point(10, filtros.Bottom + 22);
                orientacao.Location = new Point(10, quantidade.Bottom + 8);
                exportar.Location = new Point(Math.Max(10, largura - exportar.Width + 10),
                    compacto ? orientacao.Bottom + 12 : quantidade.Top);
                int topoLista = Math.Max(orientacao.Bottom, exportar.Bottom) + 18;
                int acoesY = Math.Max(topoLista + 116,
                    tela.ClientSize.Height - btnExcluir.Height - 12);
                btnExcluir.Location = new Point(Math.Max(10, tela.ClientSize.Width - 10 - btnExcluir.Width), acoesY);
                fornecedores.SetBounds(10, topoLista, largura,
                    Math.Max(100, acoesY - topoLista - 16));
            }
            tela.Resize += (s, e) => AjustarLayoutNaoEncontrados();
            AjustarLayoutNaoEncontrados();
            tela.Controls.Add(titulo);
            tela.Controls.Add(subtitulo);

            tela.Controls.Add(lblEmpresa);
            tela.Controls.Add(cmbEmpresa);

            tela.Controls.Add(lblMes);
            tela.Controls.Add(cmbMes);

            tela.Controls.Add(lblAno);
            tela.Controls.Add(numAno);

            tela.Controls.Add(btnBuscar);

            tela.Controls.Add(fornecedores);

            tela.Controls.Add(
                btnExcluir
            );

            painelConteudo.Controls.Add(
                tela
            );

            filtros.Controls.AddRange([lblEmpresa, cmbEmpresa, lblMes, cmbMes, lblAno, numAno, btnBuscar]);
            AjustarLayoutNaoEncontrados();
            CarregarItens();
        }

        // =========================================================
        // CONFIGURAÇÕES
        // =========================================================


        private void AbrirTelaConfiguracoes()
        {
            telaAtual = TelaAtual.Configuracoes;

            LimparConteudo();
            DestacarMenu(btnConfiguracoes);

            Panel tela = CriarContainer(1090);
            painelConteudo.AutoScroll = true;

            var lblBanco = CriarLabelSecao("Local do banco de dados");
            lblBanco.Location = new Point(10, 695);
            var txtPastaBanco = new TextBox
            {
                Location = new Point(10, 730), Width = 530,
                Text = LocalBanco.CarregarPasta(),
                PlaceholderText = "Vazio: usar a pasta padrão atual"
            };
            var btnPastaBanco = CriarBotao("Selecionar pasta", 200, 32);
            btnPastaBanco.Location = new Point(550, 727);
            btnPastaBanco.Click += (s, e) =>
            {
                using var seletor = new FolderBrowserDialog
                {
                    Description = "Selecione a pasta para o banco de dados",
                    UseDescriptionForTitle = true,
                    SelectedPath = Directory.Exists(txtPastaBanco.Text) ? txtPastaBanco.Text : ""
                };
                if (seletor.ShowDialog(this) == DialogResult.OK)
                    txtPastaBanco.Text = seletor.SelectedPath;
            };
            var lblLocalAtual = CriarTexto("Banco em uso: " + LocalBanco.CaminhoAtual);
            lblLocalAtual.Location = new Point(10, 770);
            lblLocalAtual.AutoSize = true;
            var btnSalvarLocal = CriarBotaoAzul("Salvar local do banco", 200, 38);
            btnSalvarLocal.Location = new Point(550, 810);
            btnSalvarLocal.Click += (s, e) =>
            {
                try
                {
                    LocalBanco.SalvarPasta(txtPastaBanco.Text);
                    txtPastaBanco.Text = LocalBanco.CarregarPasta();
                    MessageBox.Show(this,
                        "Local salvo. Feche e abra o programa para aplicar.\n\n" +
                        "Se o destino já tiver um banco, ele será utilizado. Caso contrário, o banco atual será copiado.",
                        "Local do banco", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Não foi possível salvar o local do banco.\n\n" + ex.Message,
                        "Local do banco", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            tela.Controls.AddRange([lblBanco, txtPastaBanco, btnPastaBanco, lblLocalAtual, btnSalvarLocal]);

            var lblPastaCotacoes = CriarLabelSecao("Pasta raiz das cotações");
            lblPastaCotacoes.Location = new Point(10, 880);
            var txtPastaCotacoes = new TextBox
            {
                Location = new Point(10, 915), Width = 530,
                Text = pastaRaizCotacoes,
                PlaceholderText = "Vazio: usar Documentos"
            };
            var btnSelecionarCotacoes = CriarBotao("Selecionar pasta", 200, 32);
            btnSelecionarCotacoes.Location = new Point(550, 912);
            btnSelecionarCotacoes.Click += (s, e) =>
            {
                using var seletor = new FolderBrowserDialog
                {
                    Description = "Selecione a pasta raiz das cotações",
                    UseDescriptionForTitle = true,
                    SelectedPath = Directory.Exists(txtPastaCotacoes.Text) ? txtPastaCotacoes.Text : ObterPastaRaiz()
                };
                if (seletor.ShowDialog(this) == DialogResult.OK)
                    txtPastaCotacoes.Text = seletor.SelectedPath;
            };
            var lblRaizAtual = CriarTexto("Pasta em uso: " + ObterPastaRaiz());
            lblRaizAtual.Location = new Point(10, 955);
            lblRaizAtual.AutoSize = true;
            var lblEstrutura = CriarTexto("Arquivos organizados por empresa e ano-mês. Arquivos existentes não são movidos.");
            lblEstrutura.Location = new Point(10, 985);
            lblEstrutura.AutoSize = true;
            var btnSalvarCotacoes = CriarBotaoAzul("Salvar pasta das cotações", 250, 38);
            btnSalvarCotacoes.Location = new Point(500, 1025);
            btnSalvarCotacoes.Click += (s, e) =>
            {
                try
                {
                    configuracoesIA.SalvarPastaCotacoes(txtPastaCotacoes.Text);
                    pastaRaizCotacoes = configuracoesIA.CarregarPastaCotacoes();
                    txtPastaCotacoes.Text = pastaRaizCotacoes;
                    lblRaizAtual.Text = "Pasta em uso: " + ObterPastaRaiz();
                    MessageBox.Show(this, "Pasta salva. As próximas criações e consultas de cotações usarão esse local.",
                        "Pasta das cotações", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Não foi possível salvar a pasta das cotações.\n\n" + ex.Message,
                        "Pasta das cotações", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            tela.Controls.AddRange([lblPastaCotacoes, txtPastaCotacoes, btnSelecionarCotacoes,
                lblRaizAtual, lblEstrutura, btnSalvarCotacoes]);

            Label titulo = CriarTitulo("Configurações");
            titulo.Location = new Point(10, 5);

            Label lblSecao = CriarLabelSecao("Agente de IA");
            lblSecao.Font = new Font(
                "Segoe UI",
                11F,
                FontStyle.Bold
            );
            lblSecao.Location = new Point(10, 65);

            // =========================================================
            // PROVEDOR
            // =========================================================

            Label lblProvedor = CriarLabelSecao("Provedor");
            lblProvedor.Location = new Point(10, 115);

            ComboBox cmbProvedor = CriarCombo(350);
            cmbProvedor.Location = new Point(10, 145);

            cmbProvedor.Items.AddRange(
                new object[]
                {
            "OpenAI",
            "Claude (Anthropic)",
            "Gemini (Google)"
                }
            );

            // =========================================================
            // MODELO
            // =========================================================

            Label lblModelo = CriarLabelSecao("Modelo");
            lblModelo.Location = new Point(390, 115);

            ComboBox cmbModelo = CriarCombo(360);
            cmbModelo.Location = new Point(390, 145);

            // =========================================================
            // CHAVE API
            // =========================================================

            Label lblApi = CriarLabelSecao("Chave da API");
            lblApi.Location = new Point(10, 205);

            TextBox txtApi = new TextBox
            {
                Location = new Point(10, 235),
                Width = 680,
                Height = 30,

                BackColor = Color.FromArgb(
                    25,
                    25,
                    25
                ),

                ForeColor = Color.White,

                BorderStyle =
                    BorderStyle.FixedSingle,

                UseSystemPasswordChar =
                    true
            };

            Button btnMostrar = CriarBotao(
                "👁",
                50,
                30
            );

            btnMostrar.Location =
                new Point(700, 235);

            btnMostrar.Click +=
                (s, e) =>
                {
                    txtApi.UseSystemPasswordChar =
                        !txtApi.UseSystemPasswordChar;
                };

            // =========================================================
            // ENDPOINT
            // =========================================================

            Label lblEndpoint = CriarLabelSecao("Endpoint");
            lblEndpoint.Location = new Point(10, 295);

            TextBox txtEndpoint = new TextBox
            {
                Location = new Point(10, 325),

                Width = 740,

                ReadOnly = true,

                BackColor = Color.FromArgb(
                    25,
                    25,
                    25
                ),

                ForeColor = Color.Gray,

                BorderStyle =
                    BorderStyle.FixedSingle
            };

            // =========================================================
            // TIMEOUT
            // =========================================================

            Label lblTimeout =
                CriarLabelSecao(
                    "Timeout por processamento"
                );

            lblTimeout.Location =
                new Point(10, 380);

            NumericUpDown numTimeout =
    new NumericUpDown
    {
        Location =
            new Point(10, 410),

        Width =
            110,

        Minimum =
            30,

        Maximum =
            3600,

        Value =
            3600,

        BackColor =
            Color.FromArgb(
                25,
                25,
                25
            ),

        ForeColor =
            Color.White
    };

            Label lblSegundos =
                CriarTexto(
                    "segundos"
                );

            lblSegundos.Location =
                new Point(130, 414);

            // =========================================================
            // COMPORTAMENTO DA IA
            // =========================================================

            Label lblComportamento =
                CriarLabelSecao(
                    "Comportamento da IA"
                );

            lblComportamento.Location =
                new Point(10, 465);

            CheckBox chkFornecedor =
                new CheckBox
                {
                    Text =
                        "Identificar automaticamente o fornecedor",

                    Checked = true,
                    AutoSize = true,
                    ForeColor = Color.White,

                    Location =
                        new Point(10, 500)
                };

            CheckBox chkComparacao =
                new CheckBox
                {
                    Text =
                        "Comparar nomes semelhantes de insumos",

                    Checked = true,
                    AutoSize = true,
                    ForeColor = Color.White,

                    Location =
                        new Point(10, 530)
                };

            CheckBox chkPreco =
                new CheckBox
                {
                    Text =
                        "Normalizar preço por grama ou por unidade",

                    Checked = true,
                    AutoSize = true,
                    ForeColor = Color.White,

                    Location =
                        new Point(10, 560)
                };

            CheckBox chkFornecedorNovo =
                new CheckBox
                {
                    Text =
                        "Criar coluna quando o fornecedor não existir",

                    Checked = true,
                    AutoSize = true,
                    ForeColor = Color.White,

                    Location =
                        new Point(390, 500)
                };

            CheckBox chkNaoEncontrado =
                new CheckBox
                {
                    Text =
                        "Registrar insumos não encontrados",

                    Checked = true,
                    AutoSize = true,
                    ForeColor = Color.White,

                    Location =
                        new Point(390, 530)
                };

            var chkGravarJson = new CheckBox
            {
                Text = "Ligar gravação arquivo JSON de retorno IA",
                Checked = gravarJsonIA,
                AutoSize = true,
                ForeColor = Color.White,
                Location = new Point(10, 595)
            };
            bool restaurandoGravacaoJson = false;
            chkGravarJson.CheckedChanged += (s, e) =>
            {
                if (restaurandoGravacaoJson) return;
                try
                {
                    configuracoesIA.SalvarGravacaoJson(chkGravarJson.Checked);
                    gravarJsonIA = chkGravarJson.Checked;
                }
                catch (Exception)
                {
                    restaurandoGravacaoJson = true;
                    chkGravarJson.Checked = gravarJsonIA;
                    restaurandoGravacaoJson = false;
                    MessageBox.Show("Não foi possível salvar a opção de gravação JSON.",
                        "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            tela.Controls.Add(chkGravarJson);

            tela.Controls.Add(new Label
            {
                Text = "Similaridade mínima dos insumos (%)",
                AutoSize = true,
                ForeColor = Color.White,
                Location = new Point(390, 560)
            });
            var numSimilaridade = new NumericUpDown
            {
                Minimum = 0, Maximum = 100, DecimalPlaces = 2,
                Value = percentualMinimoMatch,
                Location = new Point(390, 590), Width = 100
            };
            bool restaurandoSimilaridade = false;
            numSimilaridade.ValueChanged += (s, e) =>
            {
                if (restaurandoSimilaridade) return;
                try
                {
                    configuracoesIA.SalvarPercentualMinimoMatch(numSimilaridade.Value);
                    percentualMinimoMatch = numSimilaridade.Value;
                }
                catch (Exception)
                {
                    restaurandoSimilaridade = true;
                    numSimilaridade.Value = percentualMinimoMatch;
                    restaurandoSimilaridade = false;
                    MessageBox.Show("Não foi possível salvar a similaridade mínima dos insumos.",
                        "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            tela.Controls.Add(numSimilaridade);

            // =========================================================
            // STATUS
            // =========================================================

            Label lblStatus =
                new Label
                {
                    Text =
                        "Não testado",

                    AutoSize = true,

                    ForeColor =
                        Color.Gray,

                    Font =
                        new Font(
                            "Segoe UI",
                            10F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(200, 638)
                };

            // =========================================================
            // FUNÇÃO LOCAL: CARREGAR PROVEDOR
            // =========================================================

            var catalogo = new CatalogoModelosIA();
            CancellationTokenSource? consultaModelos = null;
            bool carregandoProvedor = false;
            Label lblModelos = CriarTexto("Informe a chave para consultar os modelos.");
            lblModelos.Location = new Point(390, 176);
            lblModelos.AutoSize = false;
            lblModelos.Size = new Size(360, 28);
            lblModelos.Font = new Font("Segoe UI", 8F);
            Button btnAtualizarModelos = CriarBotao("Atualizar modelos", 165, 30);
            btnAtualizarModelos.Location = new Point(585, 275);
            tela.Controls.Add(lblModelos);
            tela.Controls.Add(btnAtualizarModelos);

            void ExibirModelos(string[] modelos, string? selecionado)
            {
                cmbModelo.Items.Clear();
                cmbModelo.Items.AddRange(modelos.Cast<object>().ToArray());
                if (!string.IsNullOrEmpty(selecionado))
                {
                    if (!cmbModelo.Items.Contains(selecionado)) cmbModelo.Items.Add(selecionado);
                    cmbModelo.SelectedItem = selecionado;
                }
            }

            async Task AtualizarModelosAsync(bool aguardarDigitacao = false)
            {
                consultaModelos?.Cancel();
                var consulta = new CancellationTokenSource();
                consultaModelos = consulta;
                string provedor = cmbProvedor.Text;
                string chave = txtApi.Text.Trim();
                string? selecionado = cmbModelo.SelectedItem?.ToString();
                var cache = catalogo.LerCache(provedor, chave);
                ExibirModelos(cache, selecionado);
                if (chave.Length == 0)
                {
                    lblModelos.Text = "Informe a chave para consultar os modelos.";
                    lblModelos.ForeColor = Color.Gray;
                    consultaModelos = null;
                    consulta.Dispose();
                    return;
                }
                lblModelos.Text = "Consultando modelos...";
                lblModelos.ForeColor = Color.Gray;
                try
                {
                    if (aguardarDigitacao) await Task.Delay(800, consulta.Token);
                    var modelos = await catalogo.ConsultarAsync(provedor, chave, consulta.Token);
                    if (consulta.IsCancellationRequested || tela.IsDisposed) return;
                    selecionado = cmbModelo.SelectedItem?.ToString();
                    ExibirModelos(modelos, selecionado);
                    lblModelos.Text = selecionado != null && !modelos.Contains(selecionado)
                        ? "Modelo salvo ausente do catálogo. Selecione e teste outro."
                        : $"{modelos.Length} modelos disponíveis. Selecione e teste a conexão.";
                }
                catch (OperationCanceledException) when (consulta.IsCancellationRequested) { }
                catch (Exception)
                {
                    if (consulta.IsCancellationRequested || tela.IsDisposed) return;
                    lblModelos.Text = cache.Length > 0
                        ? "Consulta falhou. Exibindo a última lista salva."
                        : "Consulta falhou. Confira a chave e tente atualizar.";
                    lblModelos.ForeColor = Color.Orange;
                }
                finally
                {
                    if (ReferenceEquals(consultaModelos, consulta)) consultaModelos = null;
                    consulta.Dispose();
                }
            }
            tela.Disposed += (s, e) => consultaModelos?.Cancel();
            btnAtualizarModelos.Click += async (s, e) => await AtualizarModelosAsync();
            txtApi.TextChanged += async (s, e) =>
            {
                if (!carregandoProvedor) await AtualizarModelosAsync(true);
            };

            void CarregarProvedor()
            {
                carregandoProvedor = true;
                consultaModelos?.Cancel();
                string provedor = cmbProvedor.SelectedItem?.ToString() ?? "";
                cmbModelo.Items.Clear();
                txtEndpoint.Text = provedor switch
                {
                    "OpenAI" => "https://api.openai.com/v1",
                    "Claude (Anthropic)" => "https://api.anthropic.com",
                    "Gemini (Google)" => "https://generativelanguage.googleapis.com",
                    _ => ""
                };
                // Cada provedor possui sua própria chave.
                txtApi.Text = "";
                try
                {
                    ConfiguracaoIA? salva = configuracoesIA.Carregar(provedor);
                    if (salva != null)
                    {
                        if (!cmbModelo.Items.Contains(salva.Modelo))
                            cmbModelo.Items.Add(salva.Modelo);
                        cmbModelo.SelectedItem = salva.Modelo;
                        txtApi.Text = salva.ChaveApi;
                        numTimeout.Value = Math.Clamp(salva.TimeoutSegundos,
                            (int)numTimeout.Minimum, (int)numTimeout.Maximum);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Não foi possível carregar as configurações deste provedor.",
                        "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                lblStatus.Text =
                    "Não testado";

                lblStatus.ForeColor =
                    Color.Gray;
                carregandoProvedor = false;
                _ = AtualizarModelosAsync();
            }

            // =========================================================
            // ALTERAR PROVEDOR
            // =========================================================

            cmbProvedor.SelectedIndexChanged +=
                (s, e) =>
                {
                    CarregarProvedor();
                };

            // =========================================================
            // SELEÇÃO INICIAL
            // =========================================================

            if (
                cmbProvedor.Items.Contains(
                    provedorIA
                )
            )
            {
                cmbProvedor.SelectedItem =
                    provedorIA;
            }
            else
            {
                cmbProvedor.SelectedIndex = 0;
            }

            // =========================================================
            // TESTAR
            // =========================================================

            Button btnTestar =
                CriarBotao(
                    "Testar conexão",
                    170,
                    42
                );

            btnTestar.Location =
                new Point(10, 625);

            btnTestar.Click +=
    async (s, e) =>
    {
        string provedor =
            cmbProvedor.SelectedItem?
                .ToString()
            ?? "";

        string modelo =
            cmbModelo.SelectedItem?
                .ToString()
            ?? "";

        string chave =
            txtApi.Text.Trim();

        // =====================================================
        // VALIDAÇÕES
        // =====================================================

        if (
            string.IsNullOrWhiteSpace(
                provedor
            )
        )
        {
            MessageBox.Show(
                "Selecione um provedor de IA.",
                "Configuração da IA",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        if (
            string.IsNullOrWhiteSpace(
                modelo
            )
        )
        {
            MessageBox.Show(
                "Selecione um modelo.",
                "Configuração da IA",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        if (
            string.IsNullOrWhiteSpace(
                chave
            )
        )
        {
            MessageBox.Show(
                "Informe a chave da API.",
                "Configuração da IA",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        // =====================================================
        // INICIAR TESTE
        // =====================================================

        lblStatus.Text =
            "Testando conexão...";

        lblStatus.ForeColor =
            Color.Gold;

        btnTestar.Enabled =
            false;

        cmbProvedor.Enabled =
            false;

        cmbModelo.Enabled =
            false;

        txtApi.Enabled =
            false;

        try
        {
            // =================================================
            // CRIA O AGENTE CORRETO
            // =================================================

            IAgenteIA agente =
                AgenteIAFactory.Criar(
                    provedor,
                    chave,
                    modelo,
                    (int)numTimeout.Value
                );

            // =================================================
            // TESTE REAL DA API
            // =================================================

            bool conectado =
                await agente
                    .TestarConexaoAsync();

            // =================================================
            // SUCESSO
            // =================================================

            if (conectado)
            {
                lblStatus.Text =
                    "● Conectado";

                lblStatus.ForeColor =
                    Color.LightGreen;

                MessageBox.Show(
                    "Conexão realizada com sucesso.\n\n" +
                    $"Provedor: {provedor}\n" +
                    $"Modelo: {modelo}",
                    "Conexão com IA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }

            // =================================================
            // SEM RESPOSTA
            // =================================================

            else
            {
                lblStatus.Text =
                    "● Sem resposta";

                lblStatus.ForeColor =
                    Color.Orange;

                MessageBox.Show(
                    "O provedor foi acessado, mas não retornou uma resposta válida.",
                    "Conexão com IA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        // =====================================================
        // ERRO
        // =====================================================

        catch (Exception ex)
        {
            lblStatus.Text =
                "● Falha na conexão";

            lblStatus.ForeColor =
                Color.LightCoral;

            ModalErroIA.Exibir(this, ex, provedor, modelo);

        }

        // =====================================================
        // REATIVAR CAMPOS
        // =====================================================

        finally
        {
            btnTestar.Enabled =
                true;

            cmbProvedor.Enabled =
                true;

            cmbModelo.Enabled =
                true;

            txtApi.Enabled =
                true;
        }
    };

            // =========================================================
            // SALVAR
            // =========================================================

            Button btnSalvar =
                CriarBotaoAzul(
                    "Salvar configurações",
                    200,
                    42
                );

            btnSalvar.Location =
                new Point(550, 625);

            btnSalvar.Click +=
                (s, e) =>
                {
                    string provedor =
                        cmbProvedor.SelectedItem?
                            .ToString()
                        ?? "";

                    string modelo =
                        cmbModelo.SelectedItem?
                            .ToString()
                        ?? "";

                    string chave =
                        txtApi.Text.Trim();

                    if (
                        string.IsNullOrWhiteSpace(
                            provedor
                        )
                    )
                    {
                        MessageBox.Show(
                            "Selecione um provedor.",
                            "Configuração",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    if (
                        string.IsNullOrWhiteSpace(
                            modelo
                        )
                    )
                    {
                        MessageBox.Show(
                            "Selecione um modelo.",
                            "Configuração",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    if (
                        string.IsNullOrWhiteSpace(
                            chave
                        )
                    )
                    {
                        MessageBox.Show(
                            "Informe a chave da API.",
                            "Configuração",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    try
                    {
                        var restaurarConfiguracao = configuracoesIA.CapturarRestauracao();
                        var anterior = configuracoesIA.Carregar();
                        configuracoesIA.Salvar(new ConfiguracaoIA(provedor, modelo, chave,
                            (int)numTimeout.Value));
                        var refazerConfiguracao = configuracoesIA.CapturarRestauracao();
                        int timeoutNovo = (int)numTimeout.Value;
                        RegistrarDesfazer(() =>
                        {
                            restaurarConfiguracao();
                            provedorIA = anterior?.Provedor ?? "OpenAI";
                            modeloIA = anterior?.Modelo ?? "gpt-5.6-sol";
                            apiKeyIA = anterior?.ChaveApi ?? "";
                            timeoutIA = anterior?.TimeoutSegundos ?? 3600;
                            configuracaoIASalva = anterior != null;
                            AbrirTelaConfiguracoes();
                        }, () =>
                        {
                            refazerConfiguracao();
                            provedorIA = provedor; modeloIA = modelo; apiKeyIA = chave; timeoutIA = timeoutNovo;
                            configuracaoIASalva = true;
                            AbrirTelaConfiguracoes();
                        });
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Não foi possível salvar as configurações no banco. Tente novamente.",
                            "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    provedorIA =
                        provedor;

                    modeloIA =
                        modelo;

                    apiKeyIA =
                        chave;

                    timeoutIA =
                        (int)numTimeout.Value;

                    configuracaoIASalva =
                        true;

                    lblStatus.Text =
                        "● Configuração salva";

                    lblStatus.ForeColor =
                        Color.LightGreen;

                    MessageBox.Show(
                        "Configuração salva.\n\n" +
                        $"Provedor: {provedorIA}\n" +
                        $"Modelo: {modeloIA}",
                        "Configuração",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                };

            // =========================================================
            // CONTROLES
            // =========================================================

            tela.Controls.Add(titulo);

            tela.Controls.Add(lblSecao);

            tela.Controls.Add(lblProvedor);
            tela.Controls.Add(cmbProvedor);

            tela.Controls.Add(lblModelo);
            tela.Controls.Add(cmbModelo);

            tela.Controls.Add(lblApi);
            tela.Controls.Add(txtApi);
            tela.Controls.Add(btnMostrar);

            tela.Controls.Add(lblEndpoint);
            tela.Controls.Add(txtEndpoint);

            tela.Controls.Add(lblTimeout);
            tela.Controls.Add(numTimeout);
            tela.Controls.Add(lblSegundos);

            tela.Controls.Add(lblComportamento);

            tela.Controls.Add(chkFornecedor);
            tela.Controls.Add(chkComparacao);
            tela.Controls.Add(chkPreco);
            tela.Controls.Add(chkFornecedorNovo);
            tela.Controls.Add(chkNaoEncontrado);

            tela.Controls.Add(btnTestar);
            tela.Controls.Add(lblStatus);
            tela.Controls.Add(btnSalvar);

            painelConteudo.Controls.Add(
                tela
            );
        }

        // =========================================================
        // AJUDA
        // =========================================================

        private void AbrirTelaAjuda()
        {
            telaAtual =
                TelaAtual.Ajuda;

            LimparConteudo();
            DestacarMenu(btnAjuda);

            Panel tela =
                CriarContainer(650);

            Label titulo =
                CriarTitulo("Ajuda");

            titulo.Location =
                new Point(10, 5);

            Label texto =
                new Label
                {
                    Location =
                        new Point(10, 65),

                    Width = 850,
                    Height = 400,

                    ForeColor =
                        Color.WhiteSmoke,

                    Font =
                        new Font(
                            "Segoe UI",
                            11F
                        ),

                    Text =
                        "1. Escolha Orlando ou Drugstore.\r\n\r\n" +

                        "2. Informe o mês e o ano.\r\n\r\n" +

                        "3. Se a cotação do período ainda não existir, selecione uma planilha-base.\r\n\r\n" +

                        "4. A cotação será criada ou atualizada na pasta da unidade e do ano-mês.\r\n\r\n" +

                        "5. Adicione apenas os novos PDFs dos fornecedores.\r\n\r\n" +

                        "6. Clique em Processar cotações para extrair os insumos dos PDFs.\r\n\r\n" +
                        "7. A cotação será preenchida com os preços tratados. Se a gravação JSON estiver ligada nas Configurações, será salvo apenas o JSON _campos por PDF na pasta do período."
                };

            tela.Controls.Add(titulo);
            tela.Controls.Add(texto);

            painelConteudo.Controls.Add(tela);
        }

        private void AbrirTabelaCotacao(
    string empresa,
    int mes,
    int ano)
        {
            string caminho =
                ObterCaminhoPlanilhaPeriodo(
                    empresa,
                    mes,
                    ano
                );

            if (!File.Exists(caminho))
            {
                MessageBox.Show(
                    "A planilha desta cotação não foi encontrada.",
                    "Arquivo não encontrado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }

            AbrirTabelaCotacaoArquivo(caminho, empresa, mes, ano);
        }

        private void AbrirTabelaCotacaoArquivo(string caminho, string empresa, int mes, int ano)
        {
            // =========================================================
            // ESCONDER MENU LATERAL E TOPO
            // =========================================================

            painelMenu.Visible = false;
            painelTopo.Visible = false;

            painelConteudo.Controls.Clear();
            painelConteudo.AutoScroll = false;

            // =========================================================
            // ESTRUTURA PRINCIPAL
            // =========================================================

            TableLayoutPanel estrutura =
                new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    BackColor = corFundo,

                    ColumnCount = 1,
                    RowCount = 2,

                    Margin = new Padding(0),
                    Padding = new Padding(0)
                };

            estrutura.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    100F
                )
            );

            estrutura.RowStyles.Add(
                new RowStyle(
                    SizeType.Absolute,
                    82F
                )
            );

            estrutura.RowStyles.Add(
                new RowStyle(
                    SizeType.Percent,
                    100F
                )
            );

            painelConteudo.Controls.Add(
                estrutura
            );

            // =========================================================
            // TOPO
            // =========================================================

            Panel topo =
                new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = corTopo
                };

            estrutura.Controls.Add(
                topo,
                0,
                0
            );

            Button btnVoltar =
                CriarBotao(
                    "← Voltar",
                    110,
                    38
                );

            btnVoltar.Location =
                new Point(15, 20);

            Button btnSalvar =
                CriarBotaoAzul(
                    "Salvar alterações",
                    160,
                    38
                );

            btnSalvar.Location =
                new Point(140, 20);

            Label lblTitulo =
                new Label
                {
                    Text =
                        $"{empresa} - {ObterNomeMes(mes)}/{ano}",

                    AutoSize = true,

                    ForeColor = Color.White,

                    Font =
                        new Font(
                            "Segoe UI",
                            13F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(325, 14)
                };

            Label lblArquivo =
                new Label
                {
                    Text =
                        Path.GetFileName(caminho),

                    AutoSize = true,

                    ForeColor =
                        corTextoSecundario,

                    Font =
                        new Font(
                            "Segoe UI",
                            9F
                        ),

                    Location =
                        new Point(325, 43)
                };

            Label lblAlteracoes =
                new Label
                {
                    Text = "",

                    AutoSize = true,

                    ForeColor = Color.Gold,

                    Font =
                        new Font(
                            "Segoe UI",
                            9F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(630, 30)
                };

            topo.Controls.Add(
                btnVoltar
            );

            topo.Controls.Add(
                btnSalvar
            );

            topo.Controls.Add(
                lblTitulo
            );

            topo.Controls.Add(
                lblArquivo
            );

            topo.Controls.Add(
                lblAlteracoes
            );

            // =========================================================
            // GRID
            // =========================================================

            DataGridView dgv =
                new DataGridView
                {
                    Name = "tabelaCotacao",
                    Dock = DockStyle.Fill,

                    BackgroundColor =
                        Color.FromArgb(
                            18,
                            18,
                            18
                        ),

                    BorderStyle =
                        BorderStyle.None,

                    RowHeadersVisible =
                        false,

                    RowHeadersWidth = 24,
                    RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing,
                    AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,

                    AllowUserToAddRows =
                        false,

                    AllowUserToDeleteRows =
                        false,

                    AllowUserToOrderColumns = true,

                    AllowUserToResizeColumns =
                        true,

                    AllowUserToResizeRows =
                        true,

                    SelectionMode =
                        DataGridViewSelectionMode.ColumnHeaderSelect,

                    MultiSelect =
                        true,

                    ReadOnly =
                        false,

                    EditMode =
                        DataGridViewEditMode.EditOnKeystrokeOrF2,

                    AutoSizeColumnsMode =
                        DataGridViewAutoSizeColumnsMode.None,

                    ScrollBars =
                        ScrollBars.Both,

                    EnableHeadersVisualStyles =
                        false,

                    GridColor =
                        Color.FromArgb(
                            55,
                            55,
                            55
                        ),

                    CellBorderStyle =
                        DataGridViewCellBorderStyle.Single
                };

            dgv.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(
                    18,
                    93,
                    155
                );

            dgv.ColumnHeadersDefaultCellStyle.ForeColor =
                Color.White;

            dgv.ColumnHeadersDefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    8.5F,
                    FontStyle.Bold
                );

            dgv.ColumnHeadersDefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleCenter;

            dgv.ColumnHeadersDefaultCellStyle.WrapMode =
                DataGridViewTriState.True;

            dgv.ColumnHeadersHeight =
                65;

            dgv.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            dgv.DefaultCellStyle.BackColor =
                Color.FromArgb(
                    24,
                    24,
                    24
                );

            dgv.DefaultCellStyle.ForeColor =
                Color.White;

            dgv.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(
                    48,
                    91,
                    130
                );

            dgv.DefaultCellStyle.SelectionForeColor =
                Color.White;

            dgv.DefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    9F
                );

            dgv.DefaultCellStyle.NullValue =
                "";

            dgv.RowTemplate.Height =
                28;

            estrutura.Controls.Add(
                dgv,
                0,
                1
            );

            // Força cálculo inicial do layout
            estrutura.PerformLayout();
            dgv.PerformLayout();

            // =========================================================
            // ALTERAÇÕES
            // =========================================================

            bool temAlteracoes = false;

            bool carregandoTabela = true;

            bool atualizandoAutomaticamente =
                false;

            HistoricoTabela? historico = null;
            void MarcarAlteracao()
            {
                if (carregandoTabela || historico?.Restaurando == true)
                    return;
                historico?.Agendar();

                temAlteracoes = true;

                lblAlteracoes.Text =
                    "● Alterações não salvas";

                lblAlteracoes.ForeColor =
                    Color.Gold;
            }

            // =========================================================
            // CARREGAR EXCEL
            // =========================================================

            try
            {
                using XLWorkbook workbook =
                    ArquivoPlanilhaExcel.Abrir(caminho);

                IXLWorksheet planilha =
                    workbook.Worksheets.First();

                IXLRange? range =
                    planilha.RangeUsed();

                if (range == null)
                {
                    MessageBox.Show(
                        "A planilha está vazia.",
                        "Planilha",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );

                    painelMenu.Visible = true;
                    painelTopo.Visible = true;

                    AbrirTelaConsultar();

                    return;
                }

                int ultimaLinha =
                    range
                        .LastRow()
                        .RowNumber();

                int ultimaColuna =
                    range
                        .LastColumn()
                        .ColumnNumber();

                var quantidadeLinhasSalva = workbook.CustomProperties
                    .FirstOrDefault(p => p.Name == "CotacaoGridRowCount");
                if (int.TryParse(Convert.ToString(quantidadeLinhasSalva?.Value), out int linhasSalvas) && linhasSalvas >= 0)
                    ultimaLinha = linhasSalvas + 1;

                // Somente cores registradas pelo editor representam pinturas do usuário.
                var coresUsuario = new Dictionary<string, int>();
                var propriedadeCores = workbook.CustomProperties
                    .FirstOrDefault(p => p.Name == "CotacaoGridUserColors");
                if (propriedadeCores?.Value is string coresJson)
                {
                    try
                    {
                        coresUsuario = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, int>>(coresJson)
                            ?? new Dictionary<string, int>();
                    }
                    catch (System.Text.Json.JsonException) { }
                }
                string[]? papeisColunas = null;
                var propriedadePapeis = workbook.CustomProperties
                    .FirstOrDefault(propriedade => propriedade.Name == "CotacaoGridRoles");
                if (propriedadePapeis?.Value is string papeisJson)
                {
                    try
                    {
                        papeisColunas = System.Text.Json.JsonSerializer.Deserialize<string[]>(papeisJson);
                    }
                    catch (System.Text.Json.JsonException)
                    {
                        // Planilhas sem metadados válidos usam os títulos para identificar colunas.
                    }
                }


                // =====================================================
                // CALCULAR LARGURAS
                // =====================================================

                // Até 26 colunas no total cabem na área disponível.
                int larguraInsumo = 210;
                int larguraFornecedorAnterior = 90;
                int larguraQuantidade = 90;
                int larguraMaisBarato = 110;
                int larguraFornecedor = Math.Max(2,
                    (dgv.ClientSize.Width - 500 - SystemInformation.VerticalScrollBarWidth - 2) / Math.Max(1, Math.Min(ultimaColuna, 26) - 4));

                // =====================================================
                // CABEÇALHOS
                // =====================================================

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
                            .GetFormattedString();

                    if (
                        string.IsNullOrWhiteSpace(
                            cabecalho
                        )
                    )
                    {
                        cabecalho =
                            $"Coluna {coluna}";
                    }

                    DataGridViewTextBoxColumn novaColuna =
                        new DataGridViewTextBoxColumn
                        {
                            Name =
                                $"COL_{coluna}",

                            HeaderText =
                                cabecalho,

                            Tag = papeisColunas != null && papeisColunas.Length == ultimaColuna
                                ? papeisColunas[coluna - 1] : NormalizarCabecalhoGrid(cabecalho),
                            MinimumWidth = 2,

                            SortMode =
                                DataGridViewColumnSortMode.NotSortable
                        };

                    // =================================================
                    // LARGURAS
                    // =================================================

                    if (coluna == 1)
                    {
                        // Nome do insumo
                        novaColuna.Width =
                            larguraInsumo;
                    }
                    else if (coluna == 2)
                    {
                        // Fornecedor anterior
                        novaColuna.Width =
                            larguraFornecedorAnterior;
                    }
                    else if (coluna == 3)
                    {
                        // Quantidade
                        novaColuna.Width =
                            larguraQuantidade;
                    }
                    else if (coluna == 4)
                    {
                        // Fornecedor mais barato
                        novaColuna.Width =
                            larguraMaisBarato;
                    }
                    else
                    {
                        // Fornecedores
                        novaColuna.Width =
                            larguraFornecedor;
                    }

                    dgv.Columns.Add(
                        novaColuna
                    );
                }

                // =====================================================
                // CARREGAR LINHAS
                // =====================================================

                for (
                    int linhaExcel = 2;
                    linhaExcel <= ultimaLinha;
                    linhaExcel++)
                {
                    int indiceLinha =
                        dgv.Rows.Add();

                    for (
                        int coluna = 1;
                        coluna <= ultimaColuna;
                        coluna++)
                    {
                        IXLCell celulaExcel =
                            planilha.Cell(
                                linhaExcel,
                                coluna
                            );

                        DataGridViewCell celulaGrid =
                            dgv.Rows[indiceLinha]
                               .Cells[coluna - 1];

                        // =============================================
                        // VALOR NUMÉRICO
                        // =============================================

                        if (
                            celulaExcel.DataType ==
                            XLDataType.Number
                        )
                        {
                            try
                            {
                                decimal valor =
                                    celulaExcel
                                        .GetValue<decimal>();

                                // Somente fornecedores
                                // devem usar 4 casas decimais
                                if (EhColunaFornecedorGrid(dgv.Columns[coluna - 1]))
                                {
                                    valor =
                                        Math.Round(
                                            valor,
                                            4,
                                            MidpointRounding.AwayFromZero
                                        );

                                    celulaGrid.Value =
                                        valor;

                                    celulaGrid.Style.Format =
                                        "0.0000";
                                }
                                else
                                {
                                    celulaGrid.Value =
                                        valor;
                                }
                            }
                            catch
                            {
                                celulaGrid.Value =
                                    celulaExcel
                                        .GetFormattedString();
                            }
                        }
                        else
                        {
                            celulaGrid.Value =
                                celulaExcel
                                    .GetFormattedString();
                        }

                        // =============================================
                        // COR
                        // =============================================

                        if (coresUsuario.TryGetValue($"{linhaExcel}:{coluna}", out int argb))
                        {
                            Color cor = Color.FromArgb(argb);
                            celulaGrid.Style.BackColor = cor;
                            celulaGrid.Style.ForeColor = ObterCorTexto(cor);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Não foi possível abrir a planilha.\n\n"
                    + ex.Message,
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                painelMenu.Visible = true;
                painelTopo.Visible = true;

                AbrirTelaConsultar();

                return;
            }

            // =========================================================
            // LOCALIZAR COLUNA "FORNECEDOR MAIS EM CONTA"
            // =========================================================

            int ObterColunaFornecedorMaisBarato()
            {
                foreach (
                    DataGridViewColumn coluna
                    in dgv.Columns
                )
                {
                    string cabecalho =
                        NormalizarCabecalhoGrid(
                            coluna.Tag as string ?? coluna.HeaderText
                        );

                    if (
                        cabecalho.Contains(
                            "FORNECEDOR MAIS EM CONTA"
                        )
                    )
                    {
                        return coluna.Index;
                    }
                }

                return -1;
            }

            // =========================================================
            // RECALCULAR AO ABRIR
            // =========================================================

            int colunaMaisBarato =
                ObterColunaFornecedorMaisBarato();

            if (colunaMaisBarato >= 0)
            {
                for (
                    int linha = 0;
                    linha < dgv.Rows.Count;
                    linha++)
                {
                    AtualizarFornecedorMaisBaratoGrid(
                        dgv,
                        linha
                    );
                }
            }

            carregandoTabela = false;

            temAlteracoes = false;

            lblAlteracoes.Text = "";

            // =========================================================
            // ALTERAÇÃO MANUAL
            // =========================================================

            dgv.CellValueChanged +=
                (s, e) =>
                {
                    if (
                        carregandoTabela || historico?.Restaurando == true ||
                        atualizandoAutomaticamente
                    )
                    {
                        return;
                    }

                    if (
                        e.RowIndex < 0 ||
                        e.ColumnIndex < 0
                    )
                    {
                        return;
                    }

                    int colunaFornecedorMaisBarato =
                        ObterColunaFornecedorMaisBarato();

                    // Identifica fornecedores pelo papel da coluna, mesmo após mover ou renomear.
                    if (
                        EhColunaFornecedorGrid(dgv.Columns[e.ColumnIndex])
                    )
                    {
                        try
                        {
                            atualizandoAutomaticamente =
                                true;

                            FormatarValorGridQuatroCasas(
                                dgv,
                                e.RowIndex,
                                e.ColumnIndex
                            );

                            AtualizarFornecedorMaisBaratoGrid(
                                dgv,
                                e.RowIndex
                            );
                        }
                        finally
                        {
                            atualizandoAutomaticamente =
                                false;
                        }
                    }

                    MarcarAlteracao();
                };

            // =========================================================
            // DELETE
            // =========================================================

            // Resize rows directly at the cell boundary without a row-header gutter.
            int linhaRedimensionando = -1, inicioArrasteY = 0, alturaInicialLinha = 0;
            int LinhaNaDivisoria(int x, int y)
            {
                var hit = dgv.HitTest(x, y);
                if (hit.RowIndex < 0 || hit.ColumnIndex < 0) return -1;
                Rectangle area = dgv.GetCellDisplayRectangle(hit.ColumnIndex, hit.RowIndex, false);
                if (Math.Abs(y - area.Bottom) <= 4) return hit.RowIndex;
                if (hit.RowIndex > 0 && Math.Abs(y - area.Top) <= 4) return hit.RowIndex - 1;
                return -1;
            }
            dgv.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Left) return;
                int linha = LinhaNaDivisoria(e.X, e.Y);
                if (linha < 0) return;
                dgv.EndEdit();
                linhaRedimensionando = linha;
                inicioArrasteY = e.Y;
                alturaInicialLinha = dgv.Rows[linha].Height;
                dgv.Capture = true;
                dgv.Cursor = Cursors.SizeNS;
            };
            dgv.MouseMove += (s, e) =>
            {
                if (linhaRedimensionando >= 0)
                {
                    var linha = dgv.Rows[linhaRedimensionando];
                    linha.Height = Math.Clamp(alturaInicialLinha + e.Y - inicioArrasteY, linha.MinimumHeight, 2000);
                    dgv.Cursor = Cursors.SizeNS;
                }
                else dgv.Cursor = LinhaNaDivisoria(e.X, e.Y) >= 0 ? Cursors.SizeNS : Cursors.Default;
            };
            dgv.MouseUp += (s, e) =>
            {
                if (linhaRedimensionando < 0) return;
                linhaRedimensionando = -1;
                historico?.Agendar();
                dgv.Capture = false;
                dgv.Cursor = Cursors.Default;
            };
            dgv.MouseCaptureChanged += (s, e) =>
            {
                if (!dgv.Capture) { linhaRedimensionando = -1; dgv.Cursor = Cursors.Default; }
            };

            bool ajustandoLarguras = false;
            bool largurasPersonalizadas = false;

            void AjustarLargurasFornecedores()
            {
                // O grid mostra a barra automaticamente quando a largura total não cabe.
                dgv.ScrollBars = ScrollBars.Both;
                if (largurasPersonalizadas) return;
                try
                {
                    ajustandoLarguras = true;
                    var fornecedores = dgv.Columns.Cast<DataGridViewColumn>()
                        .Where(EhColunaFornecedorGrid).ToArray();
                    foreach (DataGridViewColumn coluna in dgv.Columns)
                    {
                        if (EhColunaFornecedorGrid(coluna)) continue;
                        string papel = NormalizarCabecalhoGrid(coluna.Tag as string ?? coluna.HeaderText);
                        coluna.Width = papel.StartsWith("INSUMO", StringComparison.Ordinal) ? 210
                            : papel.Contains("FORNECEDOR MAIS EM CONTA") ? 110 : 90;
                    }
                    var fixas = dgv.Columns.Cast<DataGridViewColumn>()
                        .Where(coluna => coluna.Visible && !EhColunaFornecedorGrid(coluna)).ToArray();
                    fornecedores = fornecedores.Where(coluna => coluna.Visible)
                        .OrderBy(coluna => coluna.DisplayIndex).ToArray();
                    int quantidadeVisivel = fixas.Length + fornecedores.Length;
                    // As 26 colunas incluem as colunas de identificação e os fornecedores.
                    // As excedentes conservam a largura e ficam acessíveis pela barra horizontal.
                    const int limiteColunasNaTela = 26;
                    int fornecedoresNaTela = Math.Max(1,
                        Math.Min(limiteColunasNaTela, quantidadeVisivel) - fixas.Length);
                    int larguraFixa = fixas.Sum(coluna => coluna.Width);
                    int disponivel = Math.Max(fornecedoresNaTela * 2, dgv.ClientSize.Width - larguraFixa
                        - SystemInformation.VerticalScrollBarWidth - 2);
                    int largura = Math.Max(2, disponivel / fornecedoresNaTela);
                    int sobra = disponivel % fornecedoresNaTela;
                    for (int i = 0; i < fornecedores.Length; i++)
                        fornecedores[i].Width = largura + (i < sobra ? 1 : 0);
                }
                finally
                {
                    ajustandoLarguras = false;
                }
            }

            void RecalcularFornecedores()
            {
                try
                {
                    atualizandoAutomaticamente = true;
                    for (int linha = 0; linha < dgv.Rows.Count; linha++)
                        AtualizarFornecedorMaisBaratoGrid(dgv, linha);
                }
                finally
                {
                    atualizandoAutomaticamente = false;
                }
            }

            void ApagarColunasSelecionadas()
            {
                dgv.EndEdit();
                var selecionadas = dgv.SelectedColumns.Cast<DataGridViewColumn>().ToArray();
                if (selecionadas.Length == 0) return;
                foreach (var coluna in selecionadas)
                    dgv.Columns.Remove(coluna);
                RecalcularFornecedores();
                AjustarLargurasFornecedores();
                MarcarAlteracao();
            }

            DataGridViewColumn? colunaMenu = null;
            ContextMenuStrip menuCabecalho = new ContextMenuStrip();
            menuCabecalho.Items.Add("Adicionar coluna", null, (s, e) =>
            {
                using Form dialogo = new Form
                {
                    Text = "Adicionar coluna", ClientSize = new Size(420, 155),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MinimizeBox = false, MaximizeBox = false
                };
                Label legenda = new() { Text = "Título da nova coluna", AutoSize = true,
                    Location = new Point(15, 15) };
                TextBox titulo = new() { Location = new Point(15, 42), Width = 390 };
                Button adicionar = new() { Text = "Adicionar", Location = new Point(215, 100),
                    Size = new Size(90, 32) };
                Button cancelar = new() { Text = "Cancelar", Location = new Point(315, 100),
                    Size = new Size(90, 32), DialogResult = DialogResult.Cancel };
                adicionar.Click += (sender, args) =>
                {
                    if (string.IsNullOrWhiteSpace(titulo.Text))
                    {
                        MessageBox.Show(dialogo, "Digite o título da nova coluna.", "Adicionar coluna",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        titulo.Focus();
                        return;
                    }
                    dialogo.DialogResult = DialogResult.OK;
                };
                dialogo.Controls.AddRange([legenda, titulo, adicionar, cancelar]);
                dialogo.AcceptButton = adicionar;
                dialogo.CancelButton = cancelar;
                dialogo.Shown += (sender, args) => titulo.Focus();
                if (dialogo.ShowDialog(this) != DialogResult.OK) return;
                dgv.EndEdit();
                historico?.Confirmar();
                var novaColuna = new DataGridViewTextBoxColumn
                {
                    Name = $"COL_{Guid.NewGuid():N}", HeaderText = titulo.Text.Trim(),
                    Tag = "FORNECEDOR", MinimumWidth = 2,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    ReadOnly = false
                };
                novaColuna.HeaderCell.ContextMenuStrip = menuCabecalho;
                dgv.Columns.Add(novaColuna);
                AjustarLargurasFornecedores();
                dgv.ClearSelection();
                novaColuna.Selected = true;
                dgv.FirstDisplayedScrollingColumnIndex = novaColuna.Index;
                MarcarAlteracao();
            });

            menuCabecalho.Items.Add("Editar título", null, (s, e) =>
            {
                if (colunaMenu == null || colunaMenu.DataGridView != dgv) return;
                using Form dialogo = new Form
                {
                    Text = "Editar título da coluna", ClientSize = new Size(420, 120),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MinimizeBox = false, MaximizeBox = false
                };
                TextBox titulo = new TextBox
                {
                    Text = colunaMenu.HeaderText, Location = new Point(15, 15), Width = 390
                };
                Button salvarTitulo = new Button
                {
                    Text = "Salvar", Location = new Point(230, 65), DialogResult = DialogResult.OK
                };
                Button cancelar = new Button
                {
                    Text = "Cancelar", Location = new Point(320, 65), DialogResult = DialogResult.Cancel
                };
                dialogo.Controls.AddRange(new Control[] { titulo, salvarTitulo, cancelar });
                dialogo.AcceptButton = salvarTitulo;
                dialogo.CancelButton = cancelar;
                dialogo.Shown += (sender, args) => { titulo.Focus(); titulo.SelectAll(); };
                salvarTitulo.Click += (sender, args) =>
                {
                    if (string.IsNullOrWhiteSpace(titulo.Text))
                    {
                        dialogo.DialogResult = DialogResult.None;
                        titulo.Focus();
                    }
                };
                if (dialogo.ShowDialog(this) != DialogResult.OK) return;
                colunaMenu.HeaderText = titulo.Text.Trim();
                RecalcularFornecedores();
                MarcarAlteracao();
            });
            var largurasAntesDoAjuste = new Dictionary<DataGridViewColumn, int>();
            menuCabecalho.Items.Add("Ajustar ao nome da coluna", null, (s, e) =>
            {
                if (colunaMenu == null || colunaMenu.DataGridView != dgv) return;
                largurasPersonalizadas = true;
                if (largurasAntesDoAjuste.Remove(colunaMenu, out int larguraAnterior))
                {
                    colunaMenu.Width = larguraAnterior;
                }
                else
                {
                    largurasAntesDoAjuste[colunaMenu] = colunaMenu.Width;
                    var fonte = dgv.ColumnHeadersDefaultCellStyle.Font ?? dgv.Font;
                    colunaMenu.Width = Math.Max(colunaMenu.MinimumWidth,
                        TextRenderer.MeasureText(colunaMenu.HeaderText, fonte).Width + 24);
                }
                dgv.ScrollBars = ScrollBars.Both;
            });
            menuCabecalho.Items.Add("Apagar coluna", null, (s, e) => ApagarColunasSelecionadas());
            ToolStripMenuItem trocarColuna = new ToolStripMenuItem("Trocar de lugar com");
            menuCabecalho.Items.Add(trocarColuna);
            menuCabecalho.Opening += (s, e) =>
            {
                foreach (ToolStripItem item in trocarColuna.DropDownItems.Cast<ToolStripItem>().ToArray())
                    item.Dispose();
                trocarColuna.DropDownItems.Clear();
                if (colunaMenu == null || colunaMenu.DataGridView != dgv)
                {
                    e.Cancel = true;
                    return;
                }
                var origem = colunaMenu;
                foreach (var destino in dgv.Columns.Cast<DataGridViewColumn>()
                    .OrderBy(coluna => coluna.DisplayIndex).Where(coluna => coluna != origem))
                {
                    var opcao = new ToolStripMenuItem($"{destino.DisplayIndex + 1}. {destino.HeaderText}");
                    opcao.Click += (sender, args) =>
                    {
                        if (origem.DataGridView != dgv || destino.DataGridView != dgv) return;
                        dgv.EndEdit();
                        int posicaoOrigem = origem.DisplayIndex;
                        int posicaoDestino = destino.DisplayIndex;
                        origem.DisplayIndex = posicaoDestino;
                        destino.DisplayIndex = posicaoOrigem;
                        dgv.ClearSelection();
                        origem.Selected = true;
                        MarcarAlteracao();
                    };
                    trocarColuna.DropDownItems.Add(opcao);
                }
                trocarColuna.Enabled = trocarColuna.DropDownItems.Count > 0;
            };

            foreach (DataGridViewColumn coluna in dgv.Columns)
                coluna.HeaderCell.ContextMenuStrip = menuCabecalho;
            dgv.CellMouseDown += (s, e) =>
            {
                if (e.RowIndex != -1 || e.ColumnIndex < 0) return;
                if (e.Button == MouseButtons.Right)
                {
                    dgv.EndEdit();
                    dgv.ClearSelection();
                    colunaMenu = dgv.Columns[e.ColumnIndex];
                    colunaMenu.Selected = true;
                    dgv.Focus();
                }
            };
            dgv.ColumnDisplayIndexChanged += (s, e) => MarcarAlteracao();
            dgv.ColumnWidthChanged += (s, e) =>
            {
                if (!ajustandoLarguras && historico?.Restaurando != true)
                {
                    largurasPersonalizadas = true;
                    historico?.Agendar();
                }
            };
            dgv.ColumnAdded += (s, e) =>
            {
                if (historico?.Restaurando != true) AjustarLargurasFornecedores();
            };
            dgv.SizeChanged += (s, e) => { AjustarLargurasFornecedores(); historico?.Sincronizar(); };
            AjustarLargurasFornecedores();

            dgv.KeyDown += (s, e) =>
            {
                if ((e.KeyCode != Keys.Delete && e.KeyCode != Keys.Back) || dgv.IsCurrentCellInEditMode) return;
                if (dgv.SelectedColumns.Count > 0)
                    ApagarColunasSelecionadas();
                else
                {
                    foreach (DataGridViewCell celula in dgv.SelectedCells)
                        if (!celula.ReadOnly) celula.Value = null;
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
            };

            // =========================================================
            // MENU DE CORES
            // =========================================================

            ContextMenuStrip menu =
                new ContextMenuStrip();

            menu.Items.Add("Localizar... (Ctrl + L)", null,
                (s, e) => LocalizarNaCotacao.Abrir(this, dgv));
            menu.Items.Add("Copiar (Ctrl + C)", null,
                (s, e) => ExecutarAreaTransferencia(dgv, Keys.C));
            menu.Items.Add("Recortar (Ctrl + X)", null,
                (s, e) => ExecutarAreaTransferencia(dgv, Keys.X));
            menu.Items.Add("Colar (Ctrl + V)", null,
                (s, e) => ExecutarAreaTransferencia(dgv, Keys.V));

            ToolStripMenuItem pintar =
                new ToolStripMenuItem(
                    "Pintar células"
                );

            var cores =
                new[]
                {
            new
            {
                Nome = "Vermelho",
                Cor = Color.Red
            },

            new
            {
                Nome = "Azul",
                Cor = Color.Blue
            },

            new
            {
                Nome = "Amarelo",
                Cor = Color.Yellow
            },

            new
            {
                Nome = "Verde",
                Cor = Color.Green
            },

            new
            {
                Nome = "Laranja",
                Cor = Color.Orange
            },

            new
            {
                Nome = "Roxo",
                Cor = Color.Purple
            },

            new
            {
                Nome = "Ciano",
                Cor = Color.Cyan
            },

            new
            {
                Nome = "Magenta",
                Cor = Color.Magenta
            },

            new
            {
                Nome = "Azul claro",
                Cor = Color.LightBlue
            },

            new
            {
                Nome = "Verde claro",
                Cor = Color.LightGreen
            },

            new
            {
                Nome = "Amarelo claro",
                Cor = Color.LightYellow
            },

            new
            {
                Nome = "Rosa claro",
                Cor = Color.LightPink
            },

            new
            {
                Nome = "Cinza claro",
                Cor = Color.LightGray
            },

            new
            {
                Nome = "Branco",
                Cor = Color.White
            },

            new
            {
                Nome = "Preto",
                Cor = Color.Black
            }
                };

            foreach (
                var itemCor
                in cores
            )
            {
                ToolStripMenuItem item =
                    new ToolStripMenuItem(
                        itemCor.Nome
                    );

                item.Tag =
                    itemCor.Cor;

                Bitmap quadradoCor =
                    new Bitmap(
                        18,
                        18
                    );

                using (
                    Graphics g =
                        Graphics.FromImage(
                            quadradoCor
                        )
                )
                {
                    g.Clear(
                        itemCor.Cor
                    );

                    g.DrawRectangle(
                        Pens.Gray,
                        0,
                        0,
                        17,
                        17
                    );
                }

                item.Image =
                    quadradoCor;

                item.Click +=
                    (s, e) =>
                    {
                        if (
                            s is ToolStripMenuItem menuItem &&
                            menuItem.Tag is Color cor
                        )
                        {
                            PintarSelecaoGrid(
                                dgv,
                                cor
                            );

                            MarcarAlteracao();
                        }
                    };

                pintar.DropDownItems.Add(
                    item
                );
            }

            ToolStripMenuItem removerCor =
                new ToolStripMenuItem(
                    "Remover cor"
                );

            removerCor.Click +=
                (s, e) =>
                {
                    RemoverCorSelecaoGrid(
                        dgv
                    );

                    MarcarAlteracao();
                };

            menu.Items.Add(
                pintar
            );

            menu.Items.Add(
                new ToolStripSeparator()
            );

            menu.Items.Add(
                removerCor
            );

            // O cabeçalho usa a mesma paleta, aplicada à coluna selecionada.
            ToolStripMenuItem pintarColuna = new ToolStripMenuItem("Pintar coluna");
            foreach (ToolStripMenuItem opcaoCor in pintar.DropDownItems)
            {
                Color cor = (Color)opcaoCor.Tag!;
                ToolStripMenuItem opcao = new ToolStripMenuItem(opcaoCor.Text)
                {
                    Image = opcaoCor.Image == null ? null : (Image)opcaoCor.Image.Clone()
                };
                opcao.Click += (s, e) =>
                {
                    PintarSelecaoGrid(dgv, cor);
                    MarcarAlteracao();
                };
                pintarColuna.DropDownItems.Add(opcao);
            }
            menuCabecalho.Items.Add(new ToolStripSeparator());
            menuCabecalho.Items.Add(pintarColuna);
            menuCabecalho.Items.Add("Remover cor", null, (s, e) =>
            {
                RemoverCorSelecaoGrid(dgv);
                MarcarAlteracao();
            });

            DataGridViewRow? linhaMenu = null;
            void InserirLinha(bool abaixo)
            {
                dgv.EndEdit();
                historico?.Confirmar();
                int indice = linhaMenu?.DataGridView == dgv ? linhaMenu.Index
                    : dgv.CurrentCell?.RowIndex ?? dgv.Rows.Count;
                if (abaixo && indice < dgv.Rows.Count) indice++;
                indice = Math.Clamp(indice, 0, dgv.Rows.Count);
                dgv.Rows.Insert(indice, 1);
                dgv.ClearSelection();
                if (dgv.Columns.Count > 0) dgv.CurrentCell = dgv.Rows[indice].Cells[0];
                MarcarAlteracao();
            }
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Inserir linha acima", null, (s, e) => InserirLinha(false));
            menu.Items.Add("Inserir linha abaixo", null, (s, e) => InserirLinha(true));
            var apagarLinha = menu.Items.Add("Apagar linha", null, (s, e) =>
            {
                var linha = linhaMenu?.DataGridView == dgv ? linhaMenu : dgv.CurrentCell?.OwningRow;
                if (linha == null || linha.IsNewRow) return;
                dgv.EndEdit();
                historico?.Confirmar();
                dgv.Rows.Remove(linha);
                linhaMenu = null;
                MarcarAlteracao();
            });
            menu.Opening += (s, e) => apagarLinha.Enabled = dgv.Rows.Count > 0;

            dgv.ContextMenuStrip =
                menu;

            // =========================================================
            // CLIQUE DIREITO
            // =========================================================

            dgv.CellClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var coluna = dgv.Columns[e.ColumnIndex];
                string papel = NormalizarCabecalhoGrid(coluna.Tag as string ?? coluna.HeaderText);
                if (!papel.StartsWith("INSUMO", StringComparison.Ordinal)) return;
                var celula = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (string.IsNullOrWhiteSpace(Convert.ToString(celula.FormattedValue))) return;
                celula.Style.WrapMode = DataGridViewTriState.True;
                int alturaAnterior = dgv.Rows[e.RowIndex].Height;
                dgv.AutoResizeRow(e.RowIndex, DataGridViewAutoSizeRowMode.AllCellsExceptHeader);
                dgv.Rows[e.RowIndex].Height = Math.Max(alturaAnterior, dgv.Rows[e.RowIndex].Height);
                dgv.InvalidateRow(e.RowIndex);
            };

            dgv.CellMouseDown +=
                (s, e) =>
                {
                    if (
                        e.Button !=
                        MouseButtons.Right
                        ||
                        e.RowIndex < 0
                        ||
                        e.ColumnIndex < 0
                    )
                    {
                        return;
                    }

                    linhaMenu = dgv.Rows[e.RowIndex];
                    DataGridViewCell clicada =
                        dgv.Rows[e.RowIndex]
                           .Cells[e.ColumnIndex];

                    if (!clicada.Selected)
                    {
                        dgv.ClearSelection();

                        clicada.Selected =
                            true;

                        dgv.CurrentCell =
                            clicada;
                    }
                };

            // =========================================================
            // SALVAR
            // =========================================================

            historico = new HistoricoTabela(dgv, RegistrarDesfazer, () =>
            {
                if (estrutura.Parent != painelConteudo)
                {
                    LimparConteudo();
                    painelMenu.Visible = false; painelTopo.Visible = false;
                    painelConteudo.AutoScroll = false;
                    painelConteudo.Controls.Add(estrutura);
                }
                historicoTabelaAtivo = historico;
            }, () =>
            {
                largurasPersonalizadas = true;
                largurasAntesDoAjuste.Clear();
                temAlteracoes = true;
                lblAlteracoes.Text = "● Alterações não salvas";
                lblAlteracoes.ForeColor = Color.Gold;
            });
            historicoTabelaAtivo = historico;
            indicarAlteracaoTabela = () => { temAlteracoes = true; lblAlteracoes.Text = "● Alterações não salvas"; lblAlteracoes.ForeColor = Color.Gold; };

            btnSalvar.Click +=
                (s, e) =>
                {
                    try
                    {
                        SalvarTabelaNoExcel(
                            dgv,
                            caminho
                        );

                        temAlteracoes = false;

                        lblAlteracoes.Text =
                            "✓ Alterações salvas";

                        lblAlteracoes.ForeColor =
                            Color.LightGreen;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "Não foi possível salvar a planilha.\n\n"
                            + ex.Message,
                            "Erro ao salvar",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                };

            // =========================================================
            // VOLTAR
            // =========================================================

            void SairDaTabela(Action destino)
                {
                    dgv.EndEdit();
                    if (temAlteracoes)
                    {
                        DialogResult resposta =
                            MessageBox.Show(
                                "Existem alterações que ainda não foram salvas.\n\n"
                                + "Deseja salvar antes de sair da planilha?",
                                "Alterações não salvas",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question
                            );

                        if (
                            resposta ==
                            DialogResult.Cancel
                        )
                        {
                            return;
                        }

                        if (
                            resposta ==
                            DialogResult.Yes
                        )
                        {
                            try
                            {
                                SalvarTabelaNoExcel(
                                    dgv,
                                    caminho
                                );
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(
                                    "Não foi possível salvar a planilha.\n\n"
                                    + ex.Message,
                                    "Erro ao salvar",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );

                                return;
                            }
                        }
                    }

                    painelMenu.Visible = true;
                    painelTopo.Visible = true;

                    painelConteudo.AutoScroll =
                        true;

                    destino();
                }
            btnVoltar.Click += (s, e) => SairDaTabela(AbrirTelaConsultar);
            Button btnNaoLocalizados = CriarBotao("Insumos não encontrados", 210, 38);
            btnNaoLocalizados.BackColor = Color.FromArgb(35, 59, 87);
            btnNaoLocalizados.ForeColor = Color.FromArgb(245, 190, 85);
            topo.Controls.Add(btnNaoLocalizados);
            btnNaoLocalizados.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            void AjustarAtalhoTabela()
            {
                btnNaoLocalizados.Location = new Point(Math.Max(325,
                    topo.ClientSize.Width - btnNaoLocalizados.Width - 16), 20);
                int larguraTitulo = Math.Max(1, btnNaoLocalizados.Left - lblTitulo.Left - 16);
                lblTitulo.AutoSize = false;
                lblTitulo.AutoEllipsis = true;
                lblTitulo.Size = new Size(larguraTitulo, 28);
                lblArquivo.AutoSize = false;
                lblArquivo.AutoEllipsis = true;
                lblArquivo.Size = new Size(larguraTitulo, 20);
                lblAlteracoes.Location = new Point(325, 65);
            }
            topo.Resize += (s, e) => AjustarAtalhoTabela();
            estrutura.RowStyles[0].Height = 100F;
            AjustarAtalhoTabela();
            btnNaoLocalizados.Click += (s, e) => SairDaTabela(() =>
            {
                lojaConsultaSelecionada = empresa;
                mesSelecionado = mes;
                anoSelecionado = ano;
                anoFiltroConsulta = ano;
                AbrirTelaNaoEncontrados();
            });
        }

        private void AdicionarCorMenu(ToolStripMenuItem menuPai,DataGridView dgv,string nome,Color cor,Action aoAlterar)
        {
            ToolStripMenuItem item =
                new ToolStripMenuItem();

            item.Text =
                nome;

            item.Tag =
                cor;

            Bitmap imagem =
                new Bitmap(
                    18,
                    18
                );

            using (
                Graphics g =
                    Graphics.FromImage(
                        imagem
                    )
            )
            {
                g.Clear(cor);

                g.DrawRectangle(
                    Pens.Gray,
                    0,
                    0,
                    17,
                    17
                );
            }

            item.Image =
                imagem;

            item.Click +=
                (s, e) =>
                {
                    if (
                        s is not ToolStripMenuItem itemClicado
                    )
                        return;

                    if (
                        itemClicado.Tag
                        is not Color corSelecionada
                    )
                        return;

                    PintarSelecaoGrid(
                        dgv,
                        corSelecionada
                    );

                    aoAlterar();
                };

            menuPai.DropDownItems.Add(
                item
            );
        }

        private void PintarSelecaoGrid(
     DataGridView dgv,
     Color cor)
        {
            foreach (
                DataGridViewCell celula
                in dgv.SelectedCells
            )
            {
                celula.Style.BackColor =
                    cor;

                celula.Style.ForeColor =
                    ObterCorTexto(
                        cor
                    );
            }

            dgv.Refresh();
        }

        private void RemoverCorSelecaoGrid(
    DataGridView dgv)
        {
            foreach (
                DataGridViewCell celula
                in dgv.SelectedCells
            )
            {
                celula.Style.BackColor =
                    Color.Empty;

                celula.Style.ForeColor =
                    Color.Empty;
            }

            dgv.Refresh();
        }

        private Color ObterCorTexto(
     Color fundo)
        {
            double brilho =
                (
                    fundo.R * 299 +
                    fundo.G * 587 +
                    fundo.B * 114
                ) / 1000.0;

            if (brilho > 150)
                return Color.Black;

            return Color.White;
        }

        private void SalvarTabelaNoExcel(
    DataGridView dgv,
    string caminho)
        {
            if (string.IsNullOrWhiteSpace(caminho))
            {
                throw new Exception(
                    "O caminho da planilha não foi informado."
                );
            }

            if (!File.Exists(caminho))
            {
                throw new FileNotFoundException(
                    "A planilha não foi encontrada.",
                    caminho
                );
            }

            historicoTabelaAtivo?.Confirmar();
            byte[] arquivoAnterior = File.ReadAllBytes(caminho);
            var indicarAlteracao = indicarAlteracaoTabela;
            string? pasta =
                Path.GetDirectoryName(caminho);

            if (string.IsNullOrWhiteSpace(pasta))
            {
                throw new Exception(
                    "Não foi possível identificar a pasta da planilha."
                );
            }

            string caminhoTemporario =
                Path.Combine(
                    pasta,
                    $"temp_{Guid.NewGuid():N}.xlsx"
                );

            try
            {
                // =====================================================
                // SE O ARQUIVO ESTIVER COMO SOMENTE LEITURA,
                // REMOVE O ATRIBUTO
                // =====================================================

                FileAttributes atributos =
                    File.GetAttributes(caminho);

                if (
                    atributos.HasFlag(
                        FileAttributes.ReadOnly
                    )
                )
                {
                    File.SetAttributes(
                        caminho,
                        atributos &
                        ~FileAttributes.ReadOnly
                    );
                }

                // =====================================================
                // ABRIR O EXCEL
                // =====================================================

                using (
                    XLWorkbook workbook =
                        ArquivoPlanilhaExcel.Abrir(caminho)
                )
                {
                    IXLWorksheet planilha =
                        workbook.Worksheets.First();

                    dgv.EndEdit();
                    var colunasOrdenadas = dgv.Columns.Cast<DataGridViewColumn>()
                        .OrderBy(coluna => coluna.DisplayIndex).ToArray();
                    workbook.CustomProperties.Delete("CotacaoGridRoles");
                    workbook.CustomProperties.Add("CotacaoGridRoles",
                        System.Text.Json.JsonSerializer.Serialize(colunasOrdenadas
                            .Select(coluna => coluna.Tag as string ?? NormalizarCabecalhoGrid(coluna.HeaderText))
                            .ToArray()));
                    workbook.CustomProperties.Delete("CotacaoGridRowCount");
                    workbook.CustomProperties.Add("CotacaoGridRowCount", dgv.Rows.Count.ToString());
                    var coresUsuario = new Dictionary<string, int>();
                    for (int linhaCor = 0; linhaCor < dgv.Rows.Count; linhaCor++)
                        for (int colunaCor = 0; colunaCor < colunasOrdenadas.Length; colunaCor++)
                        {
                            Color corUsuario = dgv.Rows[linhaCor].Cells[colunasOrdenadas[colunaCor].Index].Style.BackColor;
                            if (!corUsuario.IsEmpty)
                                coresUsuario[$"{linhaCor + 2}:{colunaCor + 1}"] = corUsuario.ToArgb();
                        }
                    workbook.CustomProperties.Delete("CotacaoGridUserColors");
                    workbook.CustomProperties.Add("CotacaoGridUserColors",
                        System.Text.Json.JsonSerializer.Serialize(coresUsuario));
                    int ultimaLinhaAnterior = planilha.LastRowUsed()?.RowNumber() ?? 1;
                    if (ultimaLinhaAnterior > dgv.Rows.Count + 1)
                        planilha.Rows(dgv.Rows.Count + 2, ultimaLinhaAnterior).Delete();
                    int ultimaColunaAnterior = planilha.LastColumnUsed()?.ColumnNumber() ?? 0;
                    if (ultimaColunaAnterior > colunasOrdenadas.Length)
                        planilha.Columns(colunasOrdenadas.Length + 1, ultimaColunaAnterior).Delete();
                    // =================================================
                    // CABEÇALHOS
                    // =================================================

                    for (
                        int coluna = 0;
                        coluna < dgv.Columns.Count;
                        coluna++
                    )
                    {
                        planilha
                            .Cell(
                                1,
                                coluna + 1
                            )
                            .Value =
                            colunasOrdenadas[coluna].HeaderText;
                    }

                    // =================================================
                    // DADOS
                    // =================================================

                    for (
                        int linha = 0;
                        linha < dgv.Rows.Count;
                        linha++
                    )
                    {
                        for (
                            int coluna = 0;
                            coluna < dgv.Columns.Count;
                            coluna++
                        )
                        {
                            DataGridViewCell celulaGrid =
                                dgv.Rows[linha]
                                   .Cells[colunasOrdenadas[coluna].Index];

                            IXLCell celulaExcel =
                                planilha.Cell(
                                    linha + 2,
                                    coluna + 1
                                );

                            object? valor =
                                celulaGrid.Value;

                            // =========================================
                            // VAZIO
                            // =========================================

                            if (
                                valor == null ||
                                string.IsNullOrWhiteSpace(
                                    valor.ToString()
                                )
                            )
                            {
                                celulaExcel.Clear(
                                    XLClearOptions.Contents
                                );
                            }

                            // =========================================
                            // NUMÉRICO
                            // =========================================

                            else if (
                                decimal.TryParse(
                                    valor.ToString(),
                                    System.Globalization.NumberStyles.Any,
                                    new System.Globalization.CultureInfo(
                                        "pt-BR"
                                    ),
                                    out decimal numero
                                )
                                ||
                                decimal.TryParse(
                                    valor.ToString(),
                                    System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    out numero
                                )
                            )
                            {
                                numero =
                                    Math.Round(
                                        numero,
                                        4,
                                        MidpointRounding.AwayFromZero
                                    );

                                celulaExcel.Value =
                                    numero;

                                if (EhColunaFornecedorGrid(colunasOrdenadas[coluna]))
                                {
                                    celulaExcel
                                        .Style
                                        .NumberFormat
                                        .Format =
                                        "0.0000";
                                }
                            }

                            // =========================================
                            // TEXTO
                            // =========================================

                            else
                            {
                                celulaExcel.Value =
                                    valor.ToString();
                            }

                            // =========================================
                            // COR
                            // =========================================

                            Color cor =
                                celulaGrid
                                    .Style
                                    .BackColor;

                            if (
                                cor != Color.Empty &&
                                cor !=
                                    Color.FromArgb(
                                        24,
                                        24,
                                        24
                                    )
                            )
                            {
                                celulaExcel
                                    .Style
                                    .Fill
                                    .BackgroundColor =
                                    XLColor.FromColor(
                                        cor
                                    );
                            }
                            else
                            {
                                celulaExcel
                                    .Style
                                    .Fill
                                    .BackgroundColor =
                                    XLColor.NoColor;
                            }
                        }
                    }

                    // =================================================
                    // SALVAR EM TEMPORÁRIO
                    // =================================================

                    workbook.SaveAs(
                        caminhoTemporario
                    );
                }

                ArquivoPlanilhaExcel.ValidarArquivoTemporario(caminhoTemporario);
                // Neste ponto o ClosedXML já liberou os arquivos.

                // =====================================================
                // SUBSTITUIR O ORIGINAL
                // =====================================================

                if (File.Exists(caminho))
                {
                    File.Replace(
                        caminhoTemporario,
                        caminho,
                        null,
                        true
                    );
                }
                else
                {
                    File.Move(
                        caminhoTemporario,
                        caminho
                    );
                }
                byte[] arquivoDepois = File.ReadAllBytes(caminho);
                RegistrarDesfazer(() =>
                {
                    File.WriteAllBytes(caminho, arquivoAnterior);
                    indicarAlteracao?.Invoke();
                }, () =>
                {
                    File.WriteAllBytes(caminho, arquivoDepois);
                    indicarAlteracao?.Invoke();
                });
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new Exception(
                    "O Windows não permitiu gravar a planilha.\n\n" +
                    "O arquivo pode estar protegido como somente leitura " +
                    "ou a pasta pode estar bloqueada.\n\n" +
                    ex.Message
                );
            }
            catch (IOException ex)
            {
                throw new Exception(
                    "Não foi possível substituir a planilha.\n\n" +
                    "Se ela estiver aberta no Microsoft Excel, feche o Excel.\n\n" +
                    "A tabela pode continuar aberta normalmente dentro deste programa.\n\n" +
                    ex.Message
                );
            }
            finally
            {
                try
                {
                    if (
                        File.Exists(
                            caminhoTemporario
                        )
                    )
                    {
                        File.Delete(
                            caminhoTemporario
                        );
                    }
                }
                catch
                {
                    // Ignora erro ao apagar temporário.
                }
            }
        }
        

       

        private int CalcularLarguraColunaVisualizacao(IXLWorksheet worksheet, int coluna, string cabecalho)
        {
            int larguraMinima = 90;

            int larguraMaxima = 280;

            // largura baseada no Excel
            double larguraExcel =
                worksheet
                    .Column(coluna)
                    .Width;

            int largura =
                (int)(
                    larguraExcel * 8
                );

            // também considera o tamanho do cabeçalho
            int larguraCabecalho =
                cabecalho.Length * 8 + 30;

            if (
                larguraCabecalho >
                largura
            )
            {
                largura =
                    larguraCabecalho;
            }

            if (
                largura <
                larguraMinima
            )
            {
                largura =
                    larguraMinima;
            }

            if (
                largura >
                larguraMaxima
            )
            {
                largura =
                    larguraMaxima;
            }

            return largura;
        }
        private int CalcularLarguraColuna(IXLWorksheet planilha, int numeroColuna)
        {
            double larguraExcel =
                planilha
                    .Column(numeroColuna)
                    .Width;

            int largura =
                (int)(larguraExcel * 7);

            if (largura < 100)
                largura = 100;

            if (largura > 300)
                largura = 300;

            return largura;
        }

        private void Form1_Load(
            object sender,
            EventArgs e)
        {
        }
    }
}

