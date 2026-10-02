using ClosedXML.Excel;
using OpenAI.Responses;

#pragma warning disable OPENAI001

namespace COTACAO_INSUMO
{
    public partial class Form1 : Form
    {
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
            Processando,
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

        private int ultimoTotalInsumosPreenchidos = 0;
        private int ultimoTotalInsumosNaoEncontrados = 0;

        private bool EhColunaFornecedorGrid(DataGridViewColumn coluna)
        {
            string titulo = NormalizarCabecalhoGrid(coluna.HeaderText);
            return !string.IsNullOrWhiteSpace(titulo)
                && !titulo.StartsWith("INSUMO", StringComparison.Ordinal)
                && !titulo.Contains("FORNECEDOR ANTERIOR")
                && !titulo.Contains("FORNECEDOR MAIS EM CONTA")
                && !titulo.Contains("QUANTIDADE")
                && titulo != "QTD";
        }
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
                        coluna.HeaderText
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

        private void FormatarValorGridTresCasas(
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
                    3,
                    MidpointRounding.AwayFromZero
                );

            celula.Value =
                valor;

            celula.Style.Format =
                "0.000";
        }

        private class ItemNaoEncontrado
        {
            public string Loja { get; set; } = "";

            public int Mes { get; set; }

            public int Ano { get; set; }

            public string Fornecedor { get; set; } = "";

            public string Insumo { get; set; } = "";

            public string PrecoPorGrama { get; set; } = "";

            public string TipoPreco { get; set; } = "";

            public string UnidadeOriginal { get; set; } = "";

            public string Data { get; set; } = "";
        }

        private readonly List<ItemNaoEncontrado> itensNaoEncontrados =
            new List<ItemNaoEncontrado>
            {
                new ItemNaoEncontrado
                {
                    Loja = "Drugstore",
                    Mes = 3,
                    Ano = 2026,
                    Fornecedor = "Alfa Química",
                    Insumo = "Ácido hialurônico",
                    PrecoPorGrama = "0,4587321",
                    Data = "30/03"
                },

                new ItemNaoEncontrado
                {
                    Loja = "Drugstore",
                    Mes = 3,
                    Ano = 2026,
                    Fornecedor = "Gama Farma",
                    Insumo = "Óleo de rícino",
                    PrecoPorGrama = "0,0231800",
                    Data = "30/03"
                },

                new ItemNaoEncontrado
                {
                    Loja = "Drugstore",
                    Mes = 3,
                    Ano = 2026,
                    Fornecedor = "Beta Insumos",
                    Insumo = "Pantenol",
                    PrecoPorGrama = "0,0894412",
                    Data = "30/03"
                },

                new ItemNaoEncontrado
                {
                    Loja = "Orlando",
                    Mes = 9,
                    Ano = 2026,
                    Fornecedor = "Fornecedor Orlando",
                    Insumo = "Insumo de teste",
                    PrecoPorGrama = "0,0123456789",
                    Data = "30/09"
                }
            };

        public Form1()
        {
            InitializeComponent();

            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1250, 780);

            BackColor = corFundo;
            Font = new Font("Segoe UI", 10F);

            try
            {
                ConfiguracaoIA? salva = configuracoesIA.Carregar();
                if (salva != null)
                {
                    provedorIA = salva.Provedor;
                    modeloIA = salva.Modelo;
                    apiKeyIA = salva.ChaveApi;
                    timeoutIA = salva.TimeoutSegundos;
                    configuracaoIASalva = true;
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

        private string ObterPastaRaiz()
        {
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
            void AjustarLayoutInicio()
            {
                int larguraCard = Math.Max(280, (tela.ClientSize.Width - 40) / 2);
                int alturaCard = Math.Clamp(tela.ClientSize.Height / 4, 170, 210);
                cardNova.Size = new Size(larguraCard, alturaCard);
                cardConsulta.Size = cardNova.Size;
                cardConsulta.Left = cardNova.Right + 20;
                tituloNova.Location = new Point((cardNova.ClientSize.Width - tituloNova.PreferredSize.Width) / 2, 30);
                descNova.Location = new Point((cardNova.ClientSize.Width - descNova.PreferredSize.Width) / 2, 68);
                iniciar.Location = new Point((cardNova.ClientSize.Width - iniciar.Width) / 2, alturaCard - 64);
                tituloConsulta.Location = new Point((cardConsulta.ClientSize.Width - tituloConsulta.PreferredSize.Width) / 2, 30);
                descConsulta.Location = new Point((cardConsulta.ClientSize.Width - descConsulta.PreferredSize.Width) / 2, 68);
                consultar.Location = new Point((cardConsulta.ClientSize.Width - consultar.Width) / 2, alturaCard - 64);
                acesso.Top = cardNova.Bottom + 30;
                orlando.Bounds = new Rectangle(10, acesso.Bottom + 15, larguraCard, 48);
                drugstore.Bounds = new Rectangle(cardConsulta.Left, orlando.Top, larguraCard, 48);
                naoEncontrados.Bounds = new Rectangle(10, orlando.Bottom + 20,
                    Math.Max(280, larguraCard), 48);
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
                        new Padding(8, 0, 0, 0)
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
                        $"Pasta {anoSelecionado}-{mesSelecionado:00} já existe. Será complementada.";

                    lblAviso.ForeColor =
                        Color.FromArgb(255, 183, 0);

                    lblAviso.BackColor =
                        Color.FromArgb(58, 38, 0);
                }
                else
                {
                    lblAviso.Text =
                        $"Nova tabela de cotação para {mesSelecionado:00}/{anoSelecionado} será criada.";

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
             "Cotação existente: não selecione Excel.\n" +
             "O sistema utilizará a planilha já salva.",

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

            processar.Click +=
    async (s, e) =>
    {
        // =====================================================
        // EMPRESA
        // =====================================================

        if (
            string.IsNullOrWhiteSpace(
                lojaSelecionada
            )
        )
        {
            MessageBox.Show(
                "Selecione Orlando ou Drugstore.",
                "Atenção",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        // =====================================================
        // PDF
        // =====================================================

        if (
            pdfsSelecionados.Count == 0
        )
        {
            MessageBox.Show(
                "Selecione pelo menos um PDF de fornecedor.",
                "Atenção",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        // =====================================================
        // CONFIGURAÇÃO DA IA
        // =====================================================

        if (
            !configuracaoIASalva ||
            string.IsNullOrWhiteSpace(
                apiKeyIA
            )
        )
        {
            MessageBox.Show(
                "Configure a OpenAI antes de processar a cotação.\n\n" +
                "Acesse Configurações > Agente de IA.",
                "IA não configurada",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        bool cotacaoExiste =
            PlanilhaPeriodoExiste(
                lojaSelecionada,
                mesSelecionado,
                anoSelecionado
            );

        // =====================================================
        // NOVA COTAÇÃO PRECISA DO EXCEL
        // =====================================================

        if (
            !cotacaoExiste &&
            string.IsNullOrWhiteSpace(
                caminhoExcelSelecionado
            )
        )
        {
            MessageBox.Show(
                "Esta é uma nova cotação.\n\n" +
                "Selecione a planilha Excel-base.",
                "Planilha necessária",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            return;
        }

        // =====================================================
        // PROCESSAR
        // =====================================================

        await ProcessarCotacaoComIAAsync();
    };

            AjustarContainerATela(tela);
            void AjustarLayoutNovaCotacao()
            {
                int largura = Math.Max(1, tela.ClientSize.Width - 20);
                cardLoja.Width = Math.Max(410, (largura - 10) / 2);
                cardExcel.Left = cardLoja.Right + 10;
                cardExcel.Width = Math.Max(460, largura - cardLoja.Width - 10);
                cardPdfs.Width = largura;
                cardPdfs.Height = Math.Max(160, tela.ClientSize.Height - cardPdfs.Top - 76);
                dgvPdfs.Width = Math.Max(1, cardPdfs.ClientSize.Width - 20);
                dgvPdfs.Height = Math.Max(45, cardPdfs.ClientSize.Height - 130);
                selecionarPdfs.Top = dgvPdfs.Bottom + 15;
                processar.Location = new Point(Math.Max(10, cardPdfs.Right - processar.Width),
                    cardPdfs.Bottom + 20);
            }
            tela.Resize += (s, e) => AjustarLayoutNovaCotacao();
            AjustarLayoutNovaCotacao();
            tela.Controls.Add(titulo);
            tela.Controls.Add(cardLoja);
            tela.Controls.Add(cardExcel);
            tela.Controls.Add(cardPdfs);
            tela.Controls.Add(processar);

            painelConteudo.Controls.Add(tela);
        }

        // =========================================================
        // PROCESSANDO
        // =========================================================

        private void CriarEstruturaPeriodo()
        {
            if (
                string.IsNullOrWhiteSpace(
                    lojaSelecionada
                )
            )
                return;

            string pasta =
                ObterPastaPeriodo(
                    lojaSelecionada,
                    mesSelecionado,
                    anoSelecionado
                );

            Directory.CreateDirectory(
                pasta
            );

            Directory.CreateDirectory(
                Path.Combine(
                    pasta,
                    "NaoEncontrados"
                )
            );
        }

        private void CriarArquivoPlanilhaPeriodo()
        {
            if (
                string.IsNullOrWhiteSpace(
                    caminhoExcelSelecionado
                )
            )
                return;

            string destino =
                ObterCaminhoPlanilhaPeriodo(
                    lojaSelecionada,
                    mesSelecionado,
                    anoSelecionado
                );

            string pasta =
                Path.GetDirectoryName(destino)!;

            Directory.CreateDirectory(pasta);

            File.Copy(
                caminhoExcelSelecionado,
                destino,
                true
            );
        }

        // =========================================================
        // CONCLUÍDO
        // =========================================================

        private void AbrirTelaConcluido()
        {
            telaAtual =
                TelaAtual.Concluido;

            LimparConteudo();

            Panel tela =
                CriarContainer(700);

            Label titulo =
                CriarTitulo(
                    "Processamento concluído"
                );

            titulo.Location =
                new Point(65, 15);

            Label subtitulo =
                CriarTexto(
                    "A cotação foi processada e salva com sucesso."
                );

            subtitulo.Location =
                new Point(65, 53);

            Label iconeSucesso =
                new Label
                {
                    Text = "✓",
                    AutoSize = false,
                    Width = 38,
                    Height = 38,

                    Location =
                        new Point(15, 16),

                    ForeColor =
                        Color.LimeGreen,

                    Font =
                        new Font(
                            "Segoe UI",
                            22F,
                            FontStyle.Bold
                        ),

                    TextAlign =
                        ContentAlignment.MiddleCenter
                };

            Panel informacoes =
                CriarCard(850, 120);

            informacoes.Location =
                new Point(10, 105);

            Label pasta =
                new Label
                {
                    Text =
                        $"Pasta interna: {lojaSelecionada}  >  " +
                        $"{anoSelecionado}-{mesSelecionado:00}",

                    AutoSize = true,

                    ForeColor =
                        Color.WhiteSmoke,

                    Font =
                        new Font(
                            "Segoe UI",
                            10F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(22, 27)
                };

            Label arquivo =
                new Label
                {
                    Text =
                        $"Arquivo: Cotacao_{lojaSelecionada}_" +
                        $"{anoSelecionado}-{mesSelecionado:00}.xlsx",

                    AutoSize = true,

                    ForeColor =
                        corTextoSecundario,

                    Location =
                        new Point(22, 68)
                };

            informacoes.Controls.Add(pasta);
            informacoes.Controls.Add(arquivo);

            // -----------------------------------------------------
            // CARD PREENCHIDOS
            // -----------------------------------------------------

            Panel cardPreenchidos =
                new Panel
                {
                    Width = 405,
                    Height = 115,

                    Location =
                        new Point(10, 250),

                    BackColor =
                        Color.FromArgb(
                            8,
                            47,
                            88
                        )
                };

            Label lblPreenchidosTitulo =
                new Label
                {
                    Text = "Insumos preenchidos",
                    AutoSize = true,

                    ForeColor =
                        Color.FromArgb(
                            90,
                            175,
                            255
                        ),

                    Font =
                        new Font(
                            "Segoe UI",
                            11F
                        ),

                    Location =
                        new Point(18, 16)
                };

            Label lblPreenchidosNumero =
                new Label
                {
                    Text =
                        ultimoTotalInsumosPreenchidos
                        .ToString(),

                    AutoSize = true,

                    ForeColor =
                        Color.FromArgb(
                            115,
                            190,
                            255
                        ),

                    Font =
                        new Font(
                            "Segoe UI",
                            28F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(18, 43)
                };

            cardPreenchidos.Controls.Add(
                lblPreenchidosTitulo
            );

            cardPreenchidos.Controls.Add(
                lblPreenchidosNumero
            );

            // -----------------------------------------------------
            // CARD NÃO ENCONTRADOS
            // -----------------------------------------------------

            Panel cardNaoEncontrados =
                new Panel
                {
                    Width = 405,
                    Height = 115,

                    Location =
                        new Point(430, 250),

                    BackColor =
                        Color.FromArgb(
                            65,
                            39,
                            0
                        )
                };

            Label lblNaoTitulo =
                new Label
                {
                    Text =
                        "Insumos não localizados",

                    AutoSize = true,

                    ForeColor =
                        Color.FromArgb(
                            255,
                            185,
                            0
                        ),

                    Font =
                        new Font(
                            "Segoe UI",
                            11F
                        ),

                    Location =
                        new Point(18, 16)
                };

            Label lblNaoNumero =
                new Label
                {
                    Text =
                        ultimoTotalInsumosNaoEncontrados
                        .ToString(),

                    AutoSize = true,

                    ForeColor =
                        Color.FromArgb(
                            255,
                            190,
                            0
                        ),

                    Font =
                        new Font(
                            "Segoe UI",
                            28F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(18, 43)
                };

            cardNaoEncontrados.Controls.Add(
                lblNaoTitulo
            );

            cardNaoEncontrados.Controls.Add(
                lblNaoNumero
            );

            // -----------------------------------------------------
            // BOTÕES
            // -----------------------------------------------------

            Button consultar =
                CriarBotao(
                    "Consultar cotação",
                    200,
                    44
                );

            consultar.Location =
                new Point(10, 400);

            consultar.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        lojaSelecionada;

                    AbrirTelaConsultar();
                };

            Button verNaoEncontrados =
                CriarBotao(
                    "Ver não localizados",
                    200,
                    44
                );

            verNaoEncontrados.Location =
                new Point(225, 400);

            verNaoEncontrados.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        lojaSelecionada;

                    AbrirTelaNaoEncontrados();
                };

            Button adicionarMaisPdfs =
                CriarBotaoAzul(
                    "Adicionar mais PDFs",
                    200,
                    44
                );

            adicionarMaisPdfs.Location =
                new Point(440, 400);

            adicionarMaisPdfs.Click +=
                (s, e) =>
                {
                    PrepararComplementoCotacao();

                    AbrirTelaNovaCotacao();
                };

            Button novaCotacao =
                CriarBotao(
                    "Nova cotação",
                    180,
                    44
                );

            novaCotacao.Location =
                new Point(655, 400);

            novaCotacao.Click +=
                (s, e) =>
                {
                    PrepararNovaCotacao();

                    AbrirTelaNovaCotacao();
                };

            tela.Controls.Add(iconeSucesso);
            tela.Controls.Add(titulo);
            tela.Controls.Add(subtitulo);
            tela.Controls.Add(informacoes);
            tela.Controls.Add(cardPreenchidos);
            tela.Controls.Add(cardNaoEncontrados);
            tela.Controls.Add(consultar);
            tela.Controls.Add(verNaoEncontrados);
            tela.Controls.Add(adicionarMaisPdfs);
            tela.Controls.Add(novaCotacao);

            painelConteudo.Controls.Add(tela);
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

            ultimoTotalInsumosPreenchidos = 0;
            ultimoTotalInsumosNaoEncontrados = 0;
        }

        private void PrepararComplementoCotacao()
        {
            // mantém loja, mês e ano
            // porque vamos complementar a cotação existente

            pdfsSelecionados.Clear();

            // Excel antigo também desaparece
            caminhoExcelSelecionado = "";

            ultimoTotalInsumosPreenchidos = 0;
            ultimoTotalInsumosNaoEncontrados = 0;
        }

        // =========================================================
        // CONSULTAR
        // =========================================================

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
                        if (Directory.Exists(destino))
                            Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(destino,
                                Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                                Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin,
                                Microsoft.VisualBasic.FileIO.UICancelOption.ThrowException);
                    }
                    catch (Exception ex)
                    {
                        erros.Add($"{linha.Cells["Mes"].Value}/{linha.Cells["Ano"].Value}: {ex.Message}");
                    }
                }
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
            void AjustarLayoutConsulta()
            {
                dgv.Width = Math.Max(1, tela.ClientSize.Width - 20);
                dgv.Height = Math.Max(80, tela.ClientSize.Height - dgv.Top - 100);
                excluir.Location = new Point(dgv.Right - excluir.Width, dgv.Bottom + 15);
                selecionarTudo.Location = new Point(excluir.Left - selecionarTudo.Width - 15, excluir.Top);
                vazio.Top = selecionarTudo.Bottom + 12;
            }
            tela.Resize += (s, e) => AjustarLayoutConsulta();
            AjustarLayoutConsulta();
            tela.Controls.AddRange(new Control[] { titulo, orlando, drugstore,
                lblAno, campoAno, buscar, dgv, selecionarTudo, excluir, vazio });
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

            DataGridView dgv =
                CriarGrid();

            dgv.Location =
                new Point(10, 190);

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

            dgv.Columns["Fornecedor"].Width =
                170;

            dgv.Columns.Add(
                "Insumo",
                "Insumo"
            );

            dgv.Columns["Insumo"].Width =
                300;

            dgv.Columns.Add(
                "Preco",
                "Preço normalizado"
            );

            dgv.Columns["Preco"].Width =
                150;

            dgv.Columns.Add(
                "Tipo",
                "Tipo"
            );

            dgv.Columns["Tipo"].Width =
                110;

            dgv.Columns.Add(
                "Data",
                "Data"
            );

            dgv.Columns["Data"].Width =
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

            // Somente checkbox editável
            foreach (
                DataGridViewColumn coluna
                in dgv.Columns
            )
            {
                coluna.ReadOnly =
                    coluna.Name != "Selecionar";
            }

            Button btnSelecionarTudo =
    CriarBotao(
        "Selecionar tudo",
        130,
        38
    );

            btnSelecionarTudo.Location =
                new Point(555, 640);

            bool todosSelecionados =
                false;

            btnSelecionarTudo.Click +=
                (s, e) =>
                {
                    todosSelecionados =
                        !todosSelecionados;

                    foreach (
                        DataGridViewRow linha
                        in dgv.Rows
                    )
                    {
                        linha
                            .Cells["Selecionar"]
                            .Value =
                            todosSelecionados;
                    }

                    btnSelecionarTudo.Text =
                        todosSelecionados
                            ? "Desmarcar tudo"
                            : "Selecionar tudo";
                };

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
                dgv.Rows.Clear();

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

                List<ItemNaoEncontrado> filtrados =
                    itensNaoEncontrados
                        .Where(
                            x =>
                                x.Loja.Equals(
                                    empresa,
                                    StringComparison.OrdinalIgnoreCase
                                )
                                &&
                                x.Mes == mes
                                &&
                                x.Ano == ano
                        )
                        .ToList();

                foreach (
                    ItemNaoEncontrado item
                    in filtrados
                )
                {
                    dgv.Rows.Add(
                        item.Fornecedor,
                        item.Insumo,
                        item.PrecoPorGrama,
                        item.TipoPreco,
                        item.Data,
                        false
                    );
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
                    List<ItemNaoEncontrado>
                        itensParaExcluir =
                        new List<ItemNaoEncontrado>();

                    string empresa =
                        cmbEmpresa.SelectedItem?
                            .ToString()
                        ?? "";

                    int mes =
                        cmbMes.SelectedIndex + 1;

                    int ano =
                        (int)numAno.Value;

                    foreach (
                        DataGridViewRow linha
                        in dgv.Rows
                    )
                    {
                        bool marcado =
                            Convert.ToBoolean(
                                linha
                                    .Cells["Selecionar"]
                                    .Value
                                ?? false
                            );

                        if (!marcado)
                            continue;

                        string fornecedor =
                            linha
                                .Cells["Fornecedor"]
                                .Value?
                                .ToString()
                            ?? "";

                        string insumo =
                            linha
                                .Cells["Insumo"]
                                .Value?
                                .ToString()
                            ?? "";

                        ItemNaoEncontrado? encontrado =
                            itensNaoEncontrados
                                .FirstOrDefault(
                                    x =>
                                        x.Loja.Equals(
                                            empresa,
                                            StringComparison.OrdinalIgnoreCase
                                        )
                                        &&
                                        x.Mes == mes
                                        &&
                                        x.Ano == ano
                                        &&
                                        x.Fornecedor.Equals(
                                            fornecedor,
                                            StringComparison.OrdinalIgnoreCase
                                        )
                                        &&
                                        x.Insumo.Equals(
                                            insumo,
                                            StringComparison.OrdinalIgnoreCase
                                        )
                                );

                        if (encontrado != null)
                        {
                            itensParaExcluir.Add(
                                encontrado
                            );
                        }
                    }

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

                    foreach (
                        ItemNaoEncontrado item
                        in itensParaExcluir
                    )
                    {
                        itensNaoEncontrados.Remove(
                            item
                        );
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
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Columns["Insumo"]!.FillWeight = 180;
            dgv.Columns["Fornecedor"]!.FillWeight = 110;
            dgv.Columns["Preco"]!.FillWeight = 95;
            dgv.Columns["Tipo"]!.FillWeight = 70;
            dgv.Columns["Data"]!.FillWeight = 65;
            void AjustarLayoutNaoEncontrados()
            {
                dgv.Width = Math.Max(1, tela.ClientSize.Width - 20);
                dgv.Height = Math.Max(80, tela.ClientSize.Height - dgv.Top - 76);
                btnExcluir.Location = new Point(Math.Max(10, dgv.Right - btnExcluir.Width),
                    dgv.Bottom + 20);
                btnSelecionarTudo.Location = new Point(Math.Max(10,
                    btnExcluir.Left - btnSelecionarTudo.Width - 15), btnExcluir.Top);
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

            tela.Controls.Add(dgv);

            tela.Controls.Add(
                btnSelecionarTudo
            );

            tela.Controls.Add(
                btnExcluir
            );

            painelConteudo.Controls.Add(
                tela
            );

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

            Panel tela = CriarContainer(720);

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

            void CarregarProvedor()
            {
                string provedor =
                    cmbProvedor.SelectedItem?
                        .ToString()
                    ?? "";

                cmbModelo.Items.Clear();

                // -----------------------------------------------------
                // OPENAI
                // -----------------------------------------------------

                if (provedor == "OpenAI")
                {
                    cmbModelo.Items.AddRange(
                        new object[]
                        {
                    "gpt-5.6-sol",
                    "gpt-6-luna",
                    "gpt-6-sol"
                        }
                    );

                    txtEndpoint.Text =
                        "https://api.openai.com/v1";
                }

                // -----------------------------------------------------
                // CLAUDE
                // -----------------------------------------------------

                else if (
                    provedor ==
                    "Claude (Anthropic)"
                )
                {
                    cmbModelo.Items.AddRange(
                        new object[]
                        {
                    "claude-sonnet",
                    "claude-opus",
                    "claude-haiku"
                        }
                    );

                    txtEndpoint.Text =
                        "https://api.anthropic.com";
                }

                // -----------------------------------------------------
                // GEMINI
                // -----------------------------------------------------

                else if (
                    provedor ==
                    "Gemini (Google)"
                )
                {
                    cmbModelo.Items.AddRange(
                        new object[]
                        {
                            "gemini-3.5-flash-lite",
                            "gemini-3.8-flash"
                        }
                    );

                    txtEndpoint.Text =
                        "https://generativelanguage.googleapis.com";
                }

                if (
                    cmbModelo.Items.Count > 0
                )
                {
                    cmbModelo.SelectedIndex = 0;
                }

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

            string mensagemErro =
                ex.Message;

            // =============================================
            // OPENAI - SEM CRÉDITOS
            // =============================================

            if (
                mensagemErro.Contains(
                    "credit_balance_exhausted",
                    StringComparison.OrdinalIgnoreCase
                )
                ||
                mensagemErro.Contains(
                    "insufficient_quota",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                MessageBox.Show(
                    "A conexão com a OpenAI foi realizada, " +
                    "mas a conta da API está sem créditos disponíveis.\n\n" +
                    "Adicione créditos no faturamento da OpenAI " +
                    "e tente novamente.",
                    "OpenAI sem créditos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }

            // =============================================
            // CHAVE INVÁLIDA
            // =============================================

            if (
                mensagemErro.Contains(
                    "401"
                )
                ||
                mensagemErro.Contains(
                    "Unauthorized",
                    StringComparison.OrdinalIgnoreCase
                )
                ||
                mensagemErro.Contains(
                    "API_KEY_INVALID",
                    StringComparison.OrdinalIgnoreCase
                )
                ||
                mensagemErro.Contains(
                    "invalid api key",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                MessageBox.Show(
                    "A chave da API parece ser inválida.\n\n" +
                    "Confira a chave informada e tente novamente.",
                    "Chave inválida",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                return;
            }

            // =============================================
            // MODELO NÃO ENCONTRADO
            // =============================================

            if (
                mensagemErro.Contains(
                    "404"
                )
                ||
                mensagemErro.Contains(
                    "model",
                    StringComparison.OrdinalIgnoreCase
                )
                &&
                mensagemErro.Contains(
                    "not found",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                MessageBox.Show(
                    "O modelo selecionado não foi encontrado ou " +
                    "não está disponível para essa chave.\n\n" +
                    $"Modelo: {modelo}",
                    "Modelo indisponível",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }

            // =============================================
            // LIMITE / RATE LIMIT
            // =============================================

            if (
                mensagemErro.Contains(
                    "429"
                )
                ||
                mensagemErro.Contains(
                    "rate limit",
                    StringComparison.OrdinalIgnoreCase
                )
                ||
                mensagemErro.Contains(
                    "RESOURCE_EXHAUSTED",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                MessageBox.Show(
                    "O limite de uso da API foi atingido temporariamente.\n\n" +
                    "Aguarde alguns instantes e tente novamente.",
                    "Limite da API",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }

            // =============================================
            // TIMEOUT
            // =============================================

            if (
                ex is TaskCanceledException
                ||
                mensagemErro.Contains(
                    "timeout",
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                MessageBox.Show(
                    "A requisição demorou mais do que o tempo configurado.\n\n" +
                    $"Timeout atual: {numTimeout.Value} segundos.",
                    "Tempo excedido",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }

            // =============================================
            // ERRO GENÉRICO
            // =============================================

            MessageBox.Show(
                "Não foi possível conectar ao provedor selecionado.\n\n" +
                $"Provedor: {provedor}\n" +
                $"Modelo: {modelo}\n\n" +
                mensagemErro,
                "Erro de conexão",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
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
                        configuracoesIA.Salvar(new ConfiguracaoIA(provedor, modelo, chave,
                            (int)numTimeout.Value));
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

                        "3. Em uma cotação nova, selecione o Excel-base.\r\n\r\n" +

                        "4. Se o período já existir, não é necessário selecionar o Excel novamente.\r\n\r\n" +

                        "5. Adicione apenas os novos PDFs dos fornecedores.\r\n\r\n" +

                        "6. Clique em Processar cotações.\r\n\r\n" +

                        "7. Depois do processamento, os PDFs e o Excel selecionado são removidos da tela.\r\n\r\n" +

                        "8. A planilha existente será reutilizada automaticamente quando o período já existir."
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

            // =========================================================
            // ESCONDER MENU E TOPO
            // =========================================================

            painelMenu.Visible = false;
            painelTopo.Visible = false;

            painelConteudo.Controls.Clear();
            painelConteudo.AutoScrollPosition = Point.Empty;
            painelConteudo.AutoScroll = true;

            painelConteudo.AutoScroll = false;

            // =========================================================
            // ESTRUTURA PRINCIPAL
            // =========================================================

            TableLayoutPanel estrutura =
                new TableLayoutPanel
                {
                    Dock =
                        DockStyle.Fill,

                    BackColor =
                        corFundo,

                    ColumnCount =
                        1,

                    RowCount =
                        2,

                    Margin =
                        new Padding(0),

                    Padding =
                        new Padding(0)
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
            // TOPO DA TABELA
            // =========================================================

            Panel topo =
                new Panel
                {
                    Dock =
                        DockStyle.Fill,

                    BackColor =
                        corTopo
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

                    AutoSize =
                        true,

                    ForeColor =
                        Color.White,

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

                    AutoSize =
                        true,

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
                    Text =
                        "",

                    AutoSize =
                        true,

                    ForeColor =
                        Color.Gold,

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
                    Dock =
                        DockStyle.Fill,

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

                    AllowUserToAddRows =
                        false,

                    AllowUserToDeleteRows =
                        false,

                    AllowUserToResizeColumns =
                        true,

                    AllowUserToResizeRows =
                        false,

                    SelectionMode =
                        DataGridViewSelectionMode.CellSelect,

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

            // =========================================================
            // CONTROLE DE ALTERAÇÕES
            // =========================================================

            bool temAlteracoes =
                false;

            bool carregandoTabela =
                true;

            bool atualizandoAutomaticamente =
                false;

            void MarcarAlteracao()
            {
                if (carregandoTabela)
                    return;

                temAlteracoes =
                    true;

                lblAlteracoes.Text =
                    "● Alterações não salvas";
            }

            // =========================================================
            // CARREGAR EXCEL
            // =========================================================

            try
            {
                using XLWorkbook workbook =
                    new XLWorkbook(caminho);

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

                    painelMenu.Visible =
                        true;

                    painelTopo.Visible =
                        true;

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

                            SortMode =
                                DataGridViewColumnSortMode.NotSortable
                        };

                    // =============================================
                    // LARGURAS
                    // =============================================

                    if (coluna == 1)
                    {
                        novaColuna.Width =
                            260;
                    }
                    else if (coluna == 2)
                    {
                        novaColuna.Width =
                            130;
                    }
                    else if (coluna == 3)
                    {
                        novaColuna.Width =
                            125;
                    }
                    else if (coluna == 4)
                    {
                        novaColuna.Width =
                            150;
                    }
                    else
                    {
                        novaColuna.Width =
                            95;
                    }

                    dgv.Columns.Add(
                        novaColuna
                    );
                }

                // =====================================================
                // LINHAS
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

                        // =========================================
                        // VALOR
                        // =========================================

                        if (
                            celulaExcel.DataType ==
                            XLDataType.Number
                        )
                        {
                            decimal valor;

                            try
                            {
                                valor =
                                    celulaExcel
                                        .GetValue<decimal>();

                                valor =
                                    Math.Round(
                                        valor,
                                        3,
                                        MidpointRounding.AwayFromZero
                                    );

                                celulaGrid.Value =
                                    valor;

                                // Fornecedores geralmente começam
                                // depois das colunas fixas
                                if (coluna >= 5)
                                {
                                    celulaGrid.Style.Format =
                                        "0.000";
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

                        // =========================================
                        // COR DO EXCEL
                        // =========================================

                        try
                        {
                            XLColor corExcel =
                                celulaExcel
                                    .Style
                                    .Fill
                                    .BackgroundColor;

                            if (
                                corExcel.ColorType ==
                                XLColorType.Color
                            )
                            {
                                Color cor =
                                    corExcel.Color;

                                celulaGrid.Style.BackColor =
                                    cor;

                                celulaGrid.Style.ForeColor =
                                    ObterCorTexto(
                                        cor
                                    );
                            }
                        }
                        catch
                        {
                            // mantém tema padrão
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Não foi possível abrir a planilha.\n\n" +
                    ex.Message,
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                painelMenu.Visible =
                    true;

                painelTopo.Visible =
                    true;

                AbrirTelaConsultar();

                return;
            }

            // =========================================================
            // AJUSTAR COLUNA FORNECEDOR MAIS EM CONTA
            // =========================================================

            int ObterColunaFornecedorMaisBarato()
            {
                foreach (
                    DataGridViewColumn coluna
                    in dgv.Columns
                )
                {
                    string cabecalho =
                        coluna.HeaderText?
                            .Trim()
                        ?? "";

                    if (
                        cabecalho.Equals(
                            "Fornecedor mais em conta",
                            StringComparison.OrdinalIgnoreCase
                        )
                    )
                    {
                        return coluna.Index;
                    }
                }

                return -1;
            }

            // =========================================================
            // RECALCULAR TODAS AS LINHAS AO ABRIR
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

            // Terminou o carregamento inicial
            carregandoTabela =
                false;

            // Como o recálculo inicial não deve deixar a tela
            // marcada como modificada
            temAlteracoes =
                false;

            lblAlteracoes.Text =
                "";

            // =========================================================
            // EDIÇÃO MANUAL
            // =========================================================

            dgv.CellValueChanged +=
                (s, e) =>
                {
                    if (
                        carregandoTabela ||
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

                    if (
                        colunaFornecedorMaisBarato >= 0 &&
                        EhColunaFornecedorGrid(dgv.Columns[e.ColumnIndex])
                    )
                    {
                        try
                        {
                            atualizandoAutomaticamente =
                                true;

                            // =====================================
                            // FORMATAR VALOR COM 3 CASAS
                            // =====================================

                            FormatarValorGridTresCasas(
                                dgv,
                                e.RowIndex,
                                e.ColumnIndex
                            );

                            // =====================================
                            // RECALCULAR FORNECEDOR MAIS BARATO
                            // =====================================

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
            // Seleção explícita pelo cabeçalho distingue excluir coluna de limpar células.
            DataGridViewColumn? colunaSelecionada = null;
            dgv.AllowUserToOrderColumns = true;

            void SelecionarColuna(DataGridViewColumn coluna)
            {
                dgv.EndEdit();
                dgv.ClearSelection();
                colunaSelecionada = coluna;
                foreach (DataGridViewRow linha in dgv.Rows)
                    linha.Cells[coluna.Index].Selected = true;
                dgv.Focus();
            }

            void ExcluirColunaSelecionada()
            {
                if (colunaSelecionada == null || dgv.Columns.Count <= 1)
                    return;

                dgv.EndEdit();
                DataGridViewColumn removida = colunaSelecionada;
                colunaSelecionada = null;
                atualizandoAutomaticamente = true;
                try
                {
                    dgv.Columns.Remove(removida);
                    dgv.ClearSelection();
                    for (int linha = 0; linha < dgv.Rows.Count; linha++)
                        AtualizarFornecedorMaisBaratoGrid(dgv, linha);
                }
                finally
                {
                    atualizandoAutomaticamente = false;
                }
                MarcarAlteracao();
            }

            dgv.ColumnHeaderMouseClick += (s, e) =>
            {
                if (e.ColumnIndex >= 0)
                    SelecionarColuna(dgv.Columns[e.ColumnIndex]);
            };
            dgv.CellMouseDown += (s, e) =>
            {
                if (e.ColumnIndex < 0)
                    return;
                if (e.RowIndex < 0 && e.Button == MouseButtons.Right)
                    SelecionarColuna(dgv.Columns[e.ColumnIndex]);
                else if (e.RowIndex >= 0 &&
                    (e.Button == MouseButtons.Left ||
                     colunaSelecionada?.Index != e.ColumnIndex))
                    colunaSelecionada = null;
            };
            dgv.ColumnDisplayIndexChanged += (s, e) => MarcarAlteracao();
            // DELETE EXCLUI A COLUNA SELECIONADA OU LIMPA CÉLULAS
            // =========================================================

            dgv.KeyDown +=
                (s, e) =>
                {
                    if (
                        e.KeyCode ==
                        Keys.Delete
                    )
                    {
                        if (colunaSelecionada != null)
                        {
                            ExcluirColunaSelecionada();
                            e.Handled = true;
                            e.SuppressKeyPress = true;
                            return;
                        }
                        foreach (
                            DataGridViewCell celula
                            in dgv.SelectedCells
                        )
                        {
                            if (!celula.ReadOnly)
                            {
                                celula.Value =
                                    null;
                            }
                        }

                        e.Handled =
                            true;
                    }
                };

            // =========================================================
            // MENU DE CONTEXTO / PINTAR
            // =========================================================

            ContextMenuStrip menu =
                new ContextMenuStrip();

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

            ToolStripMenuItem editarTitulo = new ToolStripMenuItem("Editar título da coluna");
            editarTitulo.Click += (s, e) =>
            {
                DataGridViewColumn? coluna = colunaSelecionada;
                if (coluna == null)
                    return;

                using Form dialogo = new Form
                {
                    Text = "Editar título da coluna",
                    ClientSize = new Size(440, 145),
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    StartPosition = FormStartPosition.CenterParent,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    ShowInTaskbar = false
                };
                Label rotulo = new Label
                {
                    Text = "Título da coluna:",
                    AutoSize = true,
                    Location = new Point(15, 15)
                };
                TextBox campoTitulo = new TextBox
                {
                    Text = coluna.HeaderText,
                    Location = new Point(15, 40),
                    Width = 410
                };
                Button confirmar = new Button
                {
                    Text = "Confirmar",
                    Location = new Point(230, 95),
                    Size = new Size(95, 30)
                };
                Button cancelar = new Button
                {
                    Text = "Cancelar",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(330, 95),
                    Size = new Size(95, 30)
                };
                confirmar.Click += (sender, args) =>
                {
                    if (string.IsNullOrWhiteSpace(campoTitulo.Text))
                    {
                        MessageBox.Show(dialogo, "Informe um título para a coluna.",
                            "Título da coluna", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        campoTitulo.Focus();
                        return;
                    }
                    dialogo.DialogResult = DialogResult.OK;
                };
                dialogo.Controls.AddRange(new Control[] { rotulo, campoTitulo, confirmar, cancelar });
                dialogo.AcceptButton = confirmar;
                dialogo.CancelButton = cancelar;
                dialogo.Shown += (sender, args) =>
                {
                    campoTitulo.Focus();
                    campoTitulo.SelectAll();
                };
                if (dialogo.ShowDialog(dgv.FindForm()) == DialogResult.OK)
                {
                    string titulo = campoTitulo.Text.Trim();
                    if (titulo != coluna.HeaderText)
                    {
                        coluna.HeaderText = titulo;
                        MarcarAlteracao();
                    }
                }
            };
            menu.Items.Add(editarTitulo);
            menu.Opening += (s, e) => editarTitulo.Enabled = colunaSelecionada != null;
            ToolStripMenuItem excluirColuna = new ToolStripMenuItem("Excluir coluna");
            excluirColuna.Click += (s, e) => ExcluirColunaSelecionada();
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(excluirColuna);
            menu.Opening += (s, e) =>
                excluirColuna.Enabled = colunaSelecionada != null && dgv.Columns.Count > 1;
            dgv.ContextMenuStrip =
                menu;

            // =========================================================
            // CLIQUE DIREITO
            // =========================================================

            dgv.CellMouseDown +=
                (s, e) =>
                {
                    if (
                        e.Button !=
                            MouseButtons.Right ||
                        e.RowIndex < 0 ||
                        e.ColumnIndex < 0
                    )
                    {
                        return;
                    }

                    DataGridViewCell clicada =
                        dgv.Rows[e.RowIndex]
                           .Cells[e.ColumnIndex];

                    // Se a célula clicada não estiver entre
                    // as selecionadas, seleciona apenas ela.
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

            btnSalvar.Click +=
                (s, e) =>
                {
                    try
                    {
                        SalvarTabelaNoExcel(
                            dgv,
                            caminho
                        );

                        temAlteracoes =
                            false;

                        lblAlteracoes.Text =
                            "✓ Alterações salvas";

                        lblAlteracoes.ForeColor =
                            Color.LightGreen;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "Não foi possível salvar a planilha.\n\n" +
                            ex.Message,
                            "Erro ao salvar",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                };

            // =========================================================
            // VOLTAR
            // =========================================================

            btnVoltar.Click +=
                (s, e) =>
                {
                    if (temAlteracoes)
                    {
                        DialogResult resposta =
                            MessageBox.Show(
                                "Existem alterações que ainda não foram salvas.\n\n" +
                                "Deseja salvar antes de voltar?",
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
                                    "Não foi possível salvar a planilha.\n\n" +
                                    ex.Message,
                                    "Erro ao salvar",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );

                                return;
                            }
                        }
                    }

                    painelMenu.Visible =
                        true;

                    painelTopo.Visible =
                        true;

                    painelConteudo.AutoScroll =
                        true;

                    AbrirTelaConsultar();
                };
        }

        private void AdicionarCorMenu(
    ToolStripMenuItem menuPai,
    DataGridView dgv,
    string nome,
    Color cor,
    Action aoAlterar)
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
            Color fundoPadrao =
                Color.FromArgb(
                    24,
                    24,
                    24
                );

            foreach (
                DataGridViewCell celula
                in dgv.SelectedCells
            )
            {
                celula.Style.BackColor =
                    fundoPadrao;

                celula.Style.ForeColor =
                    Color.White;
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
                        new XLWorkbook(caminho)
                )
                {
                    IXLWorksheet planilha =
                        workbook.Worksheets.First();

                    dgv.EndEdit();
                    var colunasOrdenadas = dgv.Columns.Cast<DataGridViewColumn>()
                        .OrderBy(coluna => coluna.DisplayIndex).ToArray();
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
                                        3,
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
                                        "0.000";
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

        private async Task ProcessarCotacaoComIAAsync()
        {
            // Faz cópia para evitar modificação da lista
            // enquanto estamos processando.
            List<string> pdfsDoProcessamento =
                pdfsSelecionados.ToList();

            string excelOriginal =
                caminhoExcelSelecionado;

            string caminhoPlanilhaDestino =
                ObterCaminhoPlanilhaPeriodo(
                    lojaSelecionada,
                    mesSelecionado,
                    anoSelecionado
                );

            // =========================================================
            // TELA DE PROCESSAMENTO
            // =========================================================

            LimparConteudo();

            Panel tela =
                CriarContainer(
                    650
                );

            Label titulo =
                CriarTitulo(
                    "Processando cotações..."
                );

            titulo.Location =
                new Point(10, 10);

            Label descricao =
                CriarTexto(
                    "Aguarde enquanto a IA analisa os PDFs e atualiza a planilha."
                );

            descricao.Location =
                new Point(10, 50);

            ProgressBar progresso =
                new ProgressBar
                {
                    Location =
                        new Point(
                            10,
                            90
                        ),

                    Width =
                        800,

                    Height =
                        20,

                    Style =
                        ProgressBarStyle.Marquee,

                    MarqueeAnimationSpeed =
                        30
                };

            Panel card =
                CriarCard(
                    800,
                    230
                );

            card.Location =
                new Point(
                    10,
                    135
                );

            Label lblStatus =
                new Label
                {
                    Text =
                        "Preparando arquivos...",

                    AutoSize =
                        false,

                    Width =
                        750,

                    Height =
                        180,

                    ForeColor =
                        Color.White,

                    Font =
                        new Font(
                            "Segoe UI",
                            11F
                        ),

                    Location =
                        new Point(
                            20,
                            20
                        )
                };

            card.Controls.Add(
                lblStatus
            );

            tela.Controls.Add(
                titulo
            );

            tela.Controls.Add(
                descricao
            );

            tela.Controls.Add(
                progresso
            );

            tela.Controls.Add(
                card
            );

            painelConteudo.Controls.Add(
                tela
            );

            try
            {
                // =====================================================
                // 1. PREPARAR PLANILHA
                // =====================================================

                lblStatus.Text =
                    "1/5  Preparando a planilha...";

                await Task.Yield();

                CriarEstruturaPeriodo();

                bool planilhaJaExiste =
                    File.Exists(
                        caminhoPlanilhaDestino
                    );

                if (!planilhaJaExiste)
                {
                    if (
                        string.IsNullOrWhiteSpace(
                            excelOriginal
                        ) ||
                        !File.Exists(
                            excelOriginal
                        )
                    )
                    {
                        throw new Exception(
                            "O Excel-base da nova cotação não foi encontrado."
                        );
                    }

                    File.Copy(
                        excelOriginal,
                        caminhoPlanilhaDestino,
                        true
                    );
                }

                // =====================================================
                // 2. IA
                // =====================================================

                lblStatus.Text =
                    "2/5  Conectando à OpenAI...\n\n" +
                    $"Modelo: {modeloIA}";

                await Task.Yield();

                IAgenteIA iaService =
                    AgenteIAFactory.Criar(
                    provedorIA,
                    apiKeyIA,
                    modeloIA,
                    timeoutIA
                    );

                CotacaoProcessor processor =
                    new CotacaoProcessor(
                        iaService
                    );

                // =====================================================
                // 3. ANALISAR PDFs
                // =====================================================

                lblStatus.Text =
                    "3/5  Analisando os PDFs dos fornecedores...\n\n" +
                    $"{pdfsDoProcessamento.Count} arquivo(s) serão analisados.\n\n" +
                    "A IA está identificando fornecedores, produtos e correspondências.";

                await Task.Yield();

                ResultadoProcessamento resultado =
                    await processor.ProcessarAsync(
                        caminhoPlanilhaDestino,
                        pdfsDoProcessamento
                    );

                // =====================================================
                // 4. CONTADORES
                // =====================================================

                lblStatus.Text =
                    "4/5  Registrando os resultados...";

                await Task.Yield();

                ultimoTotalInsumosPreenchidos =
                    resultado.TotalPreenchidos;

                ultimoTotalInsumosNaoEncontrados =
                    resultado.NaoEncontrados.Count;

                // =====================================================
                // NÃO ENCONTRADOS
                // =====================================================

                RegistrarNaoEncontrados(
                    resultado
                );

                // =====================================================
                // 5. FINAL
                // =====================================================

                lblStatus.Text =
                    "5/5  Finalizando a cotação...\n\n" +
                    $"Preenchidos: {ultimoTotalInsumosPreenchidos}\n" +
                    $"Não encontrados: {ultimoTotalInsumosNaoEncontrados}";

                await Task.Delay(
                    500
                );

                // =====================================================
                // LIMPAR SOMENTE DEPOIS DE SUCESSO
                // =====================================================

                pdfsSelecionados.Clear();

                caminhoExcelSelecionado =
                    "";

                // =====================================================
                // CONCLUÍDO
                // =====================================================

                AbrirTelaConcluido();
            }
            catch (Exception ex)
            {
                progresso.Style =
                    ProgressBarStyle.Blocks;

                progresso.Value =
                    0;

                lblStatus.Text =
                    "O processamento foi interrompido.";

                MessageBox.Show(
                    "Não foi possível processar a cotação.\n\n" +
                    ex.Message +
                    "\n\n" +
                    "Os PDFs selecionados foram mantidos para você tentar novamente.",
                    "Erro no processamento",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                // IMPORTANTE:
                // não limpamos PDF/Excel quando dá erro.
                AbrirTelaNovaCotacao();
            }
        }

        private void RegistrarNaoEncontrados(ResultadoProcessamento resultado)
        {
            foreach (
                ItemNaoEncontradoProcessado item
                in resultado.NaoEncontrados
            )
            {
                ItemNaoEncontrado novo =
                    new ItemNaoEncontrado
                    {
                        Loja =
                            lojaSelecionada,

                        Mes =
                            mesSelecionado,

                        Ano =
                            anoSelecionado,

                        Fornecedor =
                            item.Fornecedor,

                        Insumo =
                            item.Insumo,

                        PrecoPorGrama =
                            item.PrecoNormalizado
                                .ToString(
                                    "0.############################"
                                ),

                        TipoPreco =
                            item.TipoPreco,

                        UnidadeOriginal =
                            item.UnidadeOriginal,

                        Data =
                            DateTime.Now.ToString(
                                "dd/MM"
                            )
                    };

                bool jaExiste =
                    itensNaoEncontrados.Any(
                        x =>
                            x.Loja ==
                                novo.Loja &&

                            x.Mes ==
                                novo.Mes &&

                            x.Ano ==
                                novo.Ano &&

                            x.Fornecedor ==
                                novo.Fornecedor &&

                            x.Insumo ==
                                novo.Insumo
                    );

                if (!jaExiste)
                {
                    itensNaoEncontrados.Add(
                        novo
                    );
                }
            }
        }

        private void Form1_Load(
            object sender,
            EventArgs e)
        {
        }
    }
}