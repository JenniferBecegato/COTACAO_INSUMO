using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace COTACAO_INSUMO
{
    public partial class Form1 : Form
    {
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

        private int mesSelecionado = DateTime.Now.Month;
        private int anoSelecionado = DateTime.Now.Year;

        private int ultimoTotalInsumosPreenchidos = 0;
        private int ultimoTotalInsumosNaoEncontrados = 0;

        private class ItemNaoEncontrado
        {
            public string Loja { get; set; } = "";
            public int Mes { get; set; }
            public int Ano { get; set; }
            public string Fornecedor { get; set; } = "";
            public string Insumo { get; set; } = "";
            public string PrecoPorGrama { get; set; } = "";
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
                ColumnHeadersHeight = 34
            };

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
                CriarLabelSecao("1. Loja *");

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
                CriarLabelSecao("2. Mês e ano *");

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
                CriarTexto(
                    "Cotação existente: não selecione Excel."+"\n"+"O sistema utilizará a planilha já salva."
                );

            obs1.Location =
                new Point(18, 120);

            Label obs2 =
                CriarTexto(
                    "Cotação nova: selecione o Excel-base para criar a nova planilha."
                );

            obs2.Location =
                new Point(18, 145);

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
                (s, e) =>
                {
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

                    if (pdfsSelecionados.Count == 0)
                    {
                        MessageBox.Show(
                            "Selecione pelo menos um PDF de fornecedor.",
                            "Atenção",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    bool cotacaoJaExiste =
                        PeriodoExiste(
                            lojaSelecionada,
                            mesSelecionado,
                            anoSelecionado
                        );

                    // ---------------------------------------------
                    // NOVA COTAÇÃO
                    // ---------------------------------------------

                    if (
                        !cotacaoJaExiste &&
                        string.IsNullOrWhiteSpace(
                            caminhoExcelSelecionado
                        )
                    )
                    {
                        MessageBox.Show(
                            "Esta é uma nova cotação.\n\n" +
                            "Selecione a planilha Excel-base antes de processar.",
                            "Planilha necessária",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );

                        return;
                    }

                    // ---------------------------------------------
                    // COTAÇÃO EXISTENTE
                    // ---------------------------------------------

                    if (
                        cotacaoJaExiste &&
                        !string.IsNullOrWhiteSpace(
                            caminhoExcelSelecionado
                        )
                    )
                    {
                        DialogResult resposta =
                            MessageBox.Show(
                                "Já existe uma cotação para este mês e ano.\n\n" +
                                "O sistema pode complementar a planilha já existente sem importar outro Excel.\n\n" +
                                "Deseja utilizar a planilha existente?",
                                "Cotação existente",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question
                            );

                        if (
                            resposta ==
                            DialogResult.Yes
                        )
                        {
                            caminhoExcelSelecionado = "";
                        }
                        else
                        {
                            return;
                        }
                    }

                    AbrirTelaProcessando();
                };

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

        private void AbrirTelaProcessando()
        {
            telaAtual =
                TelaAtual.Processando;

            LimparConteudo();

            Panel tela =
                CriarContainer(600);

            Label titulo =
                CriarTitulo(
                    "Processando cotações..."
                );

            titulo.Location =
                new Point(10, 10);

            Label descricao =
                CriarTexto(
                    "Aguarde enquanto os PDFs são analisados e a cotação é preenchida."
                );

            descricao.Location =
                new Point(10, 48);

            ProgressBar barra =
                new ProgressBar
                {
                    Location =
                        new Point(10, 85),

                    Width = 800,
                    Height = 18,
                    Value = 70
                };

            Panel card =
                CriarCard(800, 190);

            card.Location =
                new Point(10, 125);

            Label texto =
                CriarTexto(
                    "✓ Lendo os PDFs dos fornecedores...\n\n" +
                    "✓ Identificando fornecedores...\n\n" +
                    "✓ Comparando os insumos...\n\n" +
                    "◌ Calculando o preço por grama...\n\n" +
                    "◌ Preenchendo a planilha..."
                );

            texto.Location =
                new Point(20, 20);

            card.Controls.Add(texto);

            Button cancelar =
                CriarBotao(
                    "Cancelar",
                    130,
                    42
                );

            cancelar.Location =
                new Point(680, 345);

            cancelar.ForeColor =
                Color.LightCoral;

            cancelar.Click +=
                (s, e) =>
                    AbrirTelaNovaCotacao();

            tela.Controls.Add(titulo);
            tela.Controls.Add(descricao);
            tela.Controls.Add(barra);
            tela.Controls.Add(card);
            tela.Controls.Add(cancelar);

            painelConteudo.Controls.Add(tela);

            System.Windows.Forms.Timer timer =
                new System.Windows.Forms.Timer
                {
                    Interval = 1500
                };

            timer.Tick +=
                (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();

                    bool periodoJaExistia =
                        PeriodoExiste(
                            lojaSelecionada,
                            mesSelecionado,
                            anoSelecionado
                        );

                    CriarEstruturaPeriodo();

                    // =================================================
                    // NÚMEROS TEMPORÁRIOS
                    // =================================================
                    // Depois serão substituídos pelo resultado
                    // real do processamento com IA.

                    ultimoTotalInsumosNaoEncontrados =
                        itensNaoEncontrados.Count(
                            x =>
                                x.Loja ==
                                    lojaSelecionada &&
                                x.Mes ==
                                    mesSelecionado &&
                                x.Ano ==
                                    anoSelecionado
                        );

                    ultimoTotalInsumosPreenchidos = 0;

                    // =================================================
                    // SE FOR NOVA, CRIA A PLANILHA DO PERÍODO
                    // =================================================

                    if (!periodoJaExistia)
                    {
                        CriarArquivoPlanilhaPeriodo();
                    }

                    AbrirTelaConcluido();

                    // =================================================
                    // LIMPAR ARQUIVOS UTILIZADOS
                    // =================================================

                    pdfsSelecionados.Clear();
                    caminhoExcelSelecionado = "";
                };

            timer.Start();
        }

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
            telaAtual =
                TelaAtual.Consultar;

            LimparConteudo();
            DestacarMenu(btnConsultar);

            Panel tela =
                CriarContainer(700);

            Label titulo =
                CriarTitulo(
                    "Consultar cotações"
                );

            titulo.Location =
                new Point(10, 5);

            Button orlando =
                CriarBotao(
                    "Orlando",
                    130,
                    40
                );

            orlando.Location =
                new Point(10, 65);

            Button drugstore =
                CriarBotao(
                    "Drugstore",
                    130,
                    40
                );

            drugstore.Location =
                new Point(150, 65);

            if (
                lojaConsultaSelecionada ==
                "Orlando"
            )
                orlando.BackColor = corAzul;
            else
                drugstore.BackColor = corAzul;

            orlando.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        "Orlando";

                    AbrirTelaConsultar();
                };

            drugstore.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        "Drugstore";

                    AbrirTelaConsultar();
                };

            Label pasta =
                new Label
                {
                    Text =
                        "▣ " +
                        lojaConsultaSelecionada,

                    AutoSize = true,

                    ForeColor =
                        Color.White,

                    Font =
                        new Font(
                            "Segoe UI",
                            11F,
                            FontStyle.Bold
                        ),

                    Location =
                        new Point(800, 75)
                };

            DataGridView dgv =
                CriarGrid();

            dgv.Location =
                new Point(10, 125);

            dgv.Size =
                new Size(960, 350);

            dgv.Columns.Add(
                "Mes",
                "Mês"
            );

            dgv.Columns.Add(
                "Ano",
                "Ano"
            );

            dgv.Columns.Add(
                "Arquivo",
                "Arquivo"
            );

            dgv.Columns.Add(
                "Modificado",
                "Última modificação"
            );

            DataGridViewButtonColumn naoEncontrados =
                new DataGridViewButtonColumn
                {
                    Name = "NaoEncontrados",
                    HeaderText = "Não encontrados",
                    Text = "Visualizar",
                    UseColumnTextForButtonValue = true
                };

            dgv.Columns.Add(naoEncontrados);

            string pastaEmpresa =
                ObterPastaEmpresa(
                    lojaConsultaSelecionada
                );

            if (Directory.Exists(pastaEmpresa))
            {
                string[] pastas =
                    Directory.GetDirectories(
                        pastaEmpresa
                    );

                foreach (
                    string pastaPeriodo
                    in pastas
                        .OrderByDescending(
                            x => x
                        )
                )
                {
                    string nome =
                        Path.GetFileName(
                            pastaPeriodo
                        );

                    string[] partes =
                        nome.Split('-');

                    if (partes.Length != 2)
                        continue;

                    if (
                        !int.TryParse(
                            partes[0],
                            out int ano
                        )
                    )
                        continue;

                    if (
                        !int.TryParse(
                            partes[1],
                            out int mes
                        )
                    )
                        continue;

                    string nomeMes =
                        ObterNomeMes(
                            mes
                        );

                    string arquivoCotacao =
                        Directory
                            .GetFiles(
                                pastaPeriodo,
                                "*.xlsx"
                            )
                            .Select(
                                Path.GetFileName
                            )
                            .FirstOrDefault()
                        ??
                        "Planilha ainda não gerada";

                    DateTime modificado =
                        Directory.GetLastWriteTime(
                            pastaPeriodo
                        );

                    dgv.Rows.Add(
                        nomeMes,
                        ano,
                        arquivoCotacao,
                        modificado.ToString(
                            "dd/MM/yyyy HH:mm"
                        )
                    );
                }
            }

            Label vazio =
                CriarTexto(
                    dgv.Rows.Count == 0
                    ? "Nenhuma cotação encontrada para esta empresa."
                    : ""
                );

            vazio.Location =
                new Point(10, 500);

            dgv.CellContentClick +=
                (s, e) =>
                {
                    if (e.RowIndex < 0)
                        return;

                    if (
                        dgv.Columns[
                            e.ColumnIndex
                        ].Name ==
                        "NaoEncontrados"
                    )
                    {
                        AbrirTelaNaoEncontrados();
                    }
                };

            tela.Controls.Add(titulo);
            tela.Controls.Add(orlando);
            tela.Controls.Add(drugstore);
            tela.Controls.Add(pasta);
            tela.Controls.Add(dgv);
            tela.Controls.Add(vazio);

            painelConteudo.Controls.Add(tela);
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

            Panel tela = CriarContainer(700);

            // =========================================================
            // TÍTULO
            // =========================================================

            Label titulo =
                CriarTitulo(
                    "Insumos não encontrados"
                );

            titulo.Location =
                new Point(10, 5);

            // =========================================================
            // LOJA
            // =========================================================

            ComboBox cmbLoja =
                CriarCombo(150);

            cmbLoja.Location =
                new Point(10, 65);

            cmbLoja.Items.AddRange(
                new string[]
                {
            "Orlando",
            "Drugstore"
                }
            );

            if (
                lojaConsultaSelecionada == "Drugstore"
            )
            {
                cmbLoja.SelectedItem =
                    "Drugstore";
            }
            else
            {
                cmbLoja.SelectedItem =
                    "Orlando";
            }

            // =========================================================
            // MÊS
            // =========================================================

            ComboBox cmbMes =
                CriarCombo(200);

            cmbMes.Location =
                new Point(175, 65);

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
                DateTime.Now.Month - 1;

            // =========================================================
            // ANO
            // =========================================================

            NumericUpDown numAno =
                CriarAno(90);

            numAno.Location =
                new Point(390, 65);

            numAno.Value =
                DateTime.Now.Year;

            // =========================================================
            // BOTÃO PESQUISAR
            // =========================================================

            Button pesquisar =
                CriarBotao(
                    "Pesquisar",
                    120,
                    38
                );

            pesquisar.Location =
                new Point(495, 65);

            // =========================================================
            // GRID
            // =========================================================

            DataGridView dgv =
                CriarGrid();

            dgv.Location =
                new Point(10, 120);

            dgv.Size =
                new Size(
                    950,
                    300
                );

            // =========================================================
            // COLUNA FORNECEDOR
            // =========================================================

            dgv.Columns.Add(
                "Fornecedor",
                "Fornecedor"
            );

            // =========================================================
            // COLUNA INSUMO
            // =========================================================

            dgv.Columns.Add(
                "Insumo",
                "Insumo"
            );

            // =========================================================
            // COLUNA PREÇO
            // =========================================================

            dgv.Columns.Add(
                "Preco",
                "Preço por grama"
            );

            // =========================================================
            // COLUNA DATA
            // =========================================================

            dgv.Columns.Add(
                "Data",
                "Data"
            );

            // =========================================================
            // CHECKBOX POR ÚLTIMO
            // =========================================================

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
                        DataGridViewAutoSizeColumnMode.None,

                    FlatStyle =
                        FlatStyle.Standard
                };

            dgv.Columns.Add(
                colunaSelecionar
            );

            // =========================================================
            // AJUSTAR TAMANHOS DAS COLUNAS
            // =========================================================

            dgv.Columns["Fornecedor"].FillWeight = 24;
            dgv.Columns["Insumo"].FillWeight = 28;
            dgv.Columns["Preco"].FillWeight = 22;
            dgv.Columns["Data"].FillWeight = 12;

            dgv.Columns["Selecionar"].AutoSizeMode =
                DataGridViewAutoSizeColumnMode.None;

            dgv.Columns["Selecionar"].Width =
                38;

            // =========================================================
            // TEXTO DO RESULTADO
            // =========================================================

            Label resultado =
                CriarTexto(
                    "Selecione loja, mês e ano e clique em Pesquisar."
                );

            resultado.Location =
                new Point(10, 440);

            // =========================================================
            // MÉTODO LOCAL DE PESQUISA
            // =========================================================

            void ExecutarPesquisa()
            {
                dgv.Rows.Clear();

                if (
                    cmbLoja.SelectedIndex == -1 ||
                    cmbMes.SelectedIndex == -1
                )
                {
                    resultado.Text =
                        "Selecione loja e mês.";

                    resultado.ForeColor =
                        Color.LightCoral;

                    return;
                }

                string loja =
                    cmbLoja.SelectedItem?
                    .ToString()
                    ?? "";

                int mes =
                    cmbMes.SelectedIndex + 1;

                int ano =
                    (int)numAno.Value;

                // Mantém a loja atualmente selecionada
                lojaConsultaSelecionada =
                    loja;

                List<ItemNaoEncontrado> encontrados =
                    itensNaoEncontrados
                        .Where(
                            x =>
                                x.Loja == loja &&
                                x.Mes == mes &&
                                x.Ano == ano
                        )
                        .ToList();

                foreach (
                    ItemNaoEncontrado item
                    in encontrados
                )
                {
                    dgv.Rows.Add(
                        item.Fornecedor,
                        item.Insumo,
                        item.PrecoPorGrama,
                        item.Data,
                        false
                    );
                }

                if (
                    encontrados.Count == 0
                )
                {
                    resultado.Text =
                        "Nenhum insumo não encontrado para este período.";

                    resultado.ForeColor =
                        Color.Gray;
                }
                else if (
                    encontrados.Count == 1
                )
                {
                    resultado.Text =
                        "1 insumo encontrado.";

                    resultado.ForeColor =
                        Color.White;
                }
                else
                {
                    resultado.Text =
                        $"{encontrados.Count} insumos encontrados.";

                    resultado.ForeColor =
                        Color.White;
                }
            }

            // =========================================================
            // PESQUISAR
            // =========================================================

            pesquisar.Click +=
                (s, e) =>
                {
                    ExecutarPesquisa();
                };

            // =========================================================
            // BOTÃO EXCLUIR
            // =========================================================

            Button excluir =
                CriarBotao(
                    "Excluir selecionados",
                    220,
                    42
                );

            excluir.Location =
                new Point(740, 435);

            excluir.ForeColor =
                Color.LightCoral;

            excluir.Click +=
                (s, e) =>
                {
                    List<DataGridViewRow> linhasMarcadas =
                        new List<DataGridViewRow>();

                    // Primeiro identifica quais linhas foram marcadas
                    foreach (
                        DataGridViewRow linha
                        in dgv.Rows
                    )
                    {
                        if (
                            linha.IsNewRow
                        )
                        {
                            continue;
                        }

                        bool marcado =
                            Convert.ToBoolean(
                                linha
                                    .Cells["Selecionar"]
                                    .Value
                                ?? false
                            );

                        if (marcado)
                        {
                            linhasMarcadas.Add(
                                linha
                            );
                        }
                    }

                    // Nenhuma selecionada
                    if (
                        linhasMarcadas.Count == 0
                    )
                    {
                        MessageBox.Show(
                            "Selecione pelo menos um insumo para excluir.",
                            "Nenhum item selecionado",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );

                        return;
                    }

                    DialogResult confirmacao =
                        MessageBox.Show(
                            linhasMarcadas.Count == 1
                                ? "Deseja excluir o insumo selecionado?"
                                : $"Deseja excluir os {linhasMarcadas.Count} insumos selecionados?",
                            "Confirmar exclusão",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning
                        );

                    if (
                        confirmacao !=
                        DialogResult.Yes
                    )
                    {
                        return;
                    }

                    foreach (
                        DataGridViewRow linha
                        in linhasMarcadas
                    )
                    {
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

                        string preco =
                            linha
                                .Cells["Preco"]
                                .Value?
                                .ToString()
                            ?? "";

                        string data =
                            linha
                                .Cells["Data"]
                                .Value?
                                .ToString()
                            ?? "";

                        ItemNaoEncontrado? item =
                            itensNaoEncontrados
                                .FirstOrDefault(
                                    x =>
                                        x.Fornecedor ==
                                            fornecedor &&
                                        x.Insumo ==
                                            insumo &&
                                        x.PrecoPorGrama ==
                                            preco &&
                                        x.Data ==
                                            data
                                );

                        if (
                            item != null
                        )
                        {
                            itensNaoEncontrados.Remove(
                                item
                            );
                        }

                        dgv.Rows.Remove(
                            linha
                        );
                    }

                    // Atualiza quantidade
                    int quantidadeRestante =
                        dgv.Rows.Count;

                    if (
                        quantidadeRestante == 0
                    )
                    {
                        resultado.Text =
                            "Nenhum insumo não encontrado para este período.";

                        resultado.ForeColor =
                            Color.Gray;
                    }
                    else if (
                        quantidadeRestante == 1
                    )
                    {
                        resultado.Text =
                            "1 insumo encontrado.";

                        resultado.ForeColor =
                            Color.White;
                    }
                    else
                    {
                        resultado.Text =
                            $"{quantidadeRestante} insumos encontrados.";

                        resultado.ForeColor =
                            Color.White;
                    }
                };

            // =========================================================
            // ADICIONAR À TELA
            // =========================================================

            tela.Controls.Add(
                titulo
            );

            tela.Controls.Add(
                cmbLoja
            );

            tela.Controls.Add(
                cmbMes
            );

            tela.Controls.Add(
                numAno
            );

            tela.Controls.Add(
                pesquisar
            );

            tela.Controls.Add(
                dgv
            );

            tela.Controls.Add(
                resultado
            );

            tela.Controls.Add(
                excluir
            );

            painelConteudo.Controls.Add(
                tela
            );
        }

        // =========================================================
        // CONFIGURAÇÕES
        // =========================================================

        private void AbrirTelaConfiguracoes()
        {
            telaAtual =
                TelaAtual.Configuracoes;

            LimparConteudo();
            DestacarMenu(btnConfiguracoes);

            Panel tela =
                CriarContainer(650);

            Label titulo =
                CriarTitulo(
                    "Configurações"
                );

            titulo.Location =
                new Point(10, 5);

            Button abaIa =
                CriarBotao(
                    "Agente de IA",
                    150,
                    40
                );

            abaIa.Location =
                new Point(10, 65);

            Button abaPastas =
                CriarBotao(
                    "Pastas",
                    110,
                    40
                );

            abaPastas.Location =
                new Point(170, 65);

            Button abaGeral =
                CriarBotao(
                    "Geral",
                    100,
                    40
                );

            abaGeral.Location =
                new Point(290, 65);

            Label lblProvedor =
                CriarLabelSecao(
                    "Provedor"
                );

            lblProvedor.Location =
                new Point(10, 135);

            ComboBox provedor =
                CriarCombo(390);

            provedor.Location =
                new Point(10, 165);

            provedor.Items.AddRange(
                new string[]
                {
                    "Claude (Anthropic)",
                    "OpenAI",
                    "Gemini",
                    "OpenAI compatível"
                }
            );

            provedor.SelectedIndex = 0;

            Label lblModelo =
                CriarLabelSecao(
                    "Modelo"
                );

            lblModelo.Location =
                new Point(430, 135);

            TextBox modelo =
                new TextBox
                {
                    Location =
                        new Point(430, 165),

                    Width = 320,

                    BackColor =
                        Color.FromArgb(
                            25,
                            25,
                            25
                        ),

                    ForeColor =
                        Color.White,

                    BorderStyle =
                        BorderStyle.FixedSingle
                };

            Label lblApi =
                CriarLabelSecao(
                    "Chave da API"
                );

            lblApi.Location =
                new Point(10, 225);

            TextBox api =
                new TextBox
                {
                    Location =
                        new Point(10, 255),

                    Width = 740,

                    BackColor =
                        Color.FromArgb(
                            25,
                            25,
                            25
                        ),

                    ForeColor =
                        Color.White,

                    BorderStyle =
                        BorderStyle.FixedSingle,

                    UseSystemPasswordChar =
                        true
                };

            Button testar =
                CriarBotao(
                    "Testar conexão",
                    170,
                    42
                );

            testar.Location =
                new Point(10, 320);

            Label status =
                new Label
                {
                    Text = "Não testado",
                    AutoSize = true,
                    ForeColor = Color.Gray,
                    Location = new Point(200, 332)
                };

            testar.Click +=
                (s, e) =>
                {
                    status.Text =
                        "Conectado";

                    status.ForeColor =
                        Color.LimeGreen;
                };

            tela.Controls.Add(titulo);
            tela.Controls.Add(abaIa);
            tela.Controls.Add(abaPastas);
            tela.Controls.Add(abaGeral);
            tela.Controls.Add(lblProvedor);
            tela.Controls.Add(provedor);
            tela.Controls.Add(lblModelo);
            tela.Controls.Add(modelo);
            tela.Controls.Add(lblApi);
            tela.Controls.Add(api);
            tela.Controls.Add(testar);
            tela.Controls.Add(status);

            painelConteudo.Controls.Add(tela);
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
            string caminhoPlanilha =
                ObterCaminhoPlanilhaPeriodo(
                    empresa,
                    mes,
                    ano
                );

            if (!File.Exists(caminhoPlanilha))
            {
                MessageBox.Show(
                    "A planilha desta cotação ainda não foi encontrada.",
                    "Planilha não encontrada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return;
            }

            LimparConteudo();

            Panel tela =
                CriarContainer(800);

            // =========================================================
            // VOLTAR
            // =========================================================

            Button btnVoltar =
                CriarBotao(
                    "← Voltar",
                    120,
                    40
                );

            btnVoltar.Location =
                new Point(10, 10);

            btnVoltar.Click +=
                (s, e) =>
                {
                    lojaConsultaSelecionada =
                        empresa;

                    AbrirTelaConsultar();
                };

            // =========================================================
            // TÍTULO
            // =========================================================

            Label titulo =
                CriarTitulo(
                    $"{empresa} - {ObterNomeMes(mes)}/{ano}"
                );

            titulo.Location =
                new Point(150, 13);

            // =========================================================
            // ARQUIVO
            // =========================================================

            Label lblArquivo =
                CriarTexto(
                    Path.GetFileName(caminhoPlanilha)
                );

            lblArquivo.Location =
                new Point(150, 52);

            // =========================================================
            // TABELA
            // =========================================================

            DataGridView dgv =
                new DataGridView
                {
                    Location =
                        new Point(10, 95),

                    Size =
                        new Size(
                            Math.Max(
                                1000,
                                painelConteudo.ClientSize.Width - 90
                            ),
                            570
                        ),

                    BackgroundColor =
                        Color.FromArgb(18, 18, 18),

                    BorderStyle =
                        BorderStyle.FixedSingle,

                    GridColor =
                        Color.FromArgb(65, 65, 65),

                    RowHeadersVisible =
                        false,

                    AllowUserToAddRows =
                        false,

                    AllowUserToDeleteRows =
                        false,

                    AllowUserToResizeRows =
                        true,

                    AllowUserToResizeColumns =
                        true,

                    SelectionMode =
                        DataGridViewSelectionMode.CellSelect,

                    MultiSelect =
                        true,

                    EnableHeadersVisualStyles =
                        false,

                    AutoSizeColumnsMode =
                        DataGridViewAutoSizeColumnsMode.None,

                    ClipboardCopyMode =
                        DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
                };

            dgv.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(12, 85, 145);

            dgv.ColumnHeadersDefaultCellStyle.ForeColor =
                Color.White;

            dgv.ColumnHeadersDefaultCellStyle.Font =
                new Font(
                    "Segoe UI",
                    9.5F,
                    FontStyle.Bold
                );

            dgv.DefaultCellStyle.BackColor =
                Color.FromArgb(24, 24, 24);

            dgv.DefaultCellStyle.ForeColor =
                Color.White;

            dgv.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(45, 90, 130);

            dgv.DefaultCellStyle.SelectionForeColor =
                Color.White;

            dgv.ColumnHeadersHeight =
                36;

            dgv.RowTemplate.Height =
                30;

            // =========================================================
            // LER EXCEL
            // =========================================================

            try
            {
                using XLWorkbook workbook =
                    new XLWorkbook(
                        caminhoPlanilha
                    );

                IXLWorksheet planilha =
                    workbook.Worksheets.First();

                IXLRange? areaUsada =
                    planilha.RangeUsed();

                if (areaUsada == null)
                {
                    MessageBox.Show(
                        "A planilha está vazia.",
                        "Cotação",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );

                    return;
                }

                int ultimaLinha =
                    areaUsada.LastRow()
                        .RowNumber();

                int ultimaColuna =
                    areaUsada.LastColumn()
                        .ColumnNumber();

                // =====================================================
                // CABEÇALHOS
                // =====================================================

                for (
                    int coluna = 1;
                    coluna <= ultimaColuna;
                    coluna++
                )
                {
                    string nomeColuna =
                        planilha.Cell(1, coluna)
                            .GetFormattedString();

                    if (string.IsNullOrWhiteSpace(nomeColuna))
                    {
                        nomeColuna =
                            $"Coluna {coluna}";
                    }

                    DataGridViewTextBoxColumn colunaGrid =
                        new DataGridViewTextBoxColumn
                        {
                            Name =
                                $"COL_{coluna}",

                            HeaderText =
                                nomeColuna,

                            SortMode =
                                DataGridViewColumnSortMode.NotSortable,

                            MinimumWidth =
                                90,

                            Width =
                                CalcularLarguraColuna(
                                    planilha,
                                    coluna
                                )
                        };

                    dgv.Columns.Add(
                        colunaGrid
                    );
                }

                // =====================================================
                // DADOS
                // =====================================================

                for (
                    int linha = 2;
                    linha <= ultimaLinha;
                    linha++
                )
                {
                    object[] valores =
                        new object[ultimaColuna];

                    for (
                        int coluna = 1;
                        coluna <= ultimaColuna;
                        coluna++
                    )
                    {
                        IXLCell celula =
                            planilha.Cell(
                                linha,
                                coluna
                            );

                        valores[coluna - 1] =
                            celula.GetFormattedString();
                    }

                    int indiceLinha =
                        dgv.Rows.Add(
                            valores
                        );

                    // ================================================
                    // COPIAR ALGUNS ESTILOS DO EXCEL
                    // ================================================

                    for (
                        int coluna = 1;
                        coluna <= ultimaColuna;
                        coluna++
                    )
                    {
                        IXLCell celulaExcel =
                            planilha.Cell(
                                linha,
                                coluna
                            );

                        DataGridViewCell celulaGrid =
                            dgv.Rows[indiceLinha]
                                .Cells[coluna - 1];

                        if (celulaExcel.Style.Font.Bold)
                        {
                            celulaGrid.Style.Font =
                                new Font(
                                    "Segoe UI",
                                    9F,
                                    FontStyle.Bold
                                );
                        }

                        switch (
                            celulaExcel.Style.Alignment.Horizontal
                        )
                        {
                            case XLAlignmentHorizontalValues.Center:

                                celulaGrid.Style.Alignment =
                                    DataGridViewContentAlignment.MiddleCenter;

                                break;

                            case XLAlignmentHorizontalValues.Right:

                                celulaGrid.Style.Alignment =
                                    DataGridViewContentAlignment.MiddleRight;

                                break;

                            default:

                                celulaGrid.Style.Alignment =
                                    DataGridViewContentAlignment.MiddleLeft;

                                break;
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

                return;
            }

            // =========================================================
            // RODAPÉ
            // =========================================================

            Label lblInfo =
                CriarTexto(
                    "Esta tabela representa a planilha da cotação já preenchida com os valores encontrados nos PDFs."
                );

            lblInfo.Location =
                new Point(10, 685);

            tela.Controls.Add(
                btnVoltar
            );

            tela.Controls.Add(
                titulo
            );

            tela.Controls.Add(
                lblArquivo
            );

            tela.Controls.Add(
                dgv
            );

            tela.Controls.Add(
                lblInfo
            );

            painelConteudo.Controls.Add(
                tela
            );
        }

        private void Form1_Load(
            object sender,
            EventArgs e)
        {
        }
    }
}