namespace COTACAO_INSUMO;

public sealed class ModalErroIA : Form
{
    public ModalErroIA(Exception erro, string provedor, string modelo, bool arquivosPreservados = false, string? arquivo = null)
    {
        var mensagem = MensagemErroIA.Classificar(erro);
        Text = "Erro na comunicação com a IA";
        StartPosition = FormStartPosition.CenterParent;
        ShowInTaskbar = false;
        MinimizeBox = false;
        MaximizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 10F);
        BackColor = Color.FromArgb(18, 18, 18);
        ForeColor = Color.FromArgb(238, 238, 238);
        ClientSize = new Size(660, 480);
        MinimumSize = new Size(480, 400);

        var estrutura = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, Padding = new Padding(24), ColumnCount = 1, RowCount = 3
        };
        estrutura.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        estrutura.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        estrutura.RowStyles.Add(new RowStyle(SizeType.Absolute, 0));
        estrutura.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(estrutura);

        var conteudo = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var textos = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(0, 0, 12, 0) };
        textos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        conteudo.Controls.Add(textos);
        estrutura.Controls.Add(conteudo, 0, 0);

        void Texto(string texto, bool destaque = false, Color? cor = null)
        {
            var label = new Label
            {
                Text = texto, AutoSize = true, Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, destaque ? 10 : 18),
                ForeColor = cor ?? ForeColor,
                Font = new Font(Font.FontFamily, destaque ? 11F : 10F, destaque ? FontStyle.Bold : FontStyle.Regular)
            };
            textos.Controls.Add(label);
        }

        Texto("⚠  " + mensagem.Titulo, true, Color.FromArgb(255, 183, 90));
        string contextoArquivo = string.IsNullOrWhiteSpace(arquivo) ? "" : $"Arquivo: {Path.GetFileName(arquivo)}";
        if (contextoArquivo.Length > 0) Texto(contextoArquivo, true);
        string contexto = string.Join("   •   ", new[]
        {
            string.IsNullOrWhiteSpace(provedor) ? null : $"Provedor: {provedor}",
            string.IsNullOrWhiteSpace(modelo) ? null : $"Modelo: {modelo}",
            mensagem.StatusHttp is int status ? $"HTTP {status}" : null
        }.Where(x => x != null));
        if (contexto.Length > 0) Texto(contexto, cor: Color.FromArgb(180, 180, 180));
        Texto("O que aconteceu", true);
        Texto(mensagem.Descricao);
        Texto("O que você pode fazer", true);
        Texto(mensagem.Orientacao);
        if (arquivosPreservados) Texto("Seus PDFs e a planilha continuam selecionados para uma nova tentativa.", cor: Color.LightGreen);

        // Somente a mensagem recebida, sem stack trace ou cabeçalhos de autenticação.
        string detalhes = $"{contextoArquivo}{Environment.NewLine}{contexto}{Environment.NewLine}{Environment.NewLine}{erro.Message}";
        var caixaDetalhes = new TextBox
        {
            Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
            WordWrap = false, Text = detalhes, Visible = false,
            BackColor = Color.FromArgb(25, 25, 25), ForeColor = ForeColor,
            BorderStyle = BorderStyle.FixedSingle, AccessibleName = "Detalhes técnicos do erro"
        };
        estrutura.Controls.Add(caixaDetalhes, 0, 1);

        var acoes = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Padding = new Padding(0, 14, 0, 0) };
        estrutura.Controls.Add(acoes, 0, 2);
        Button Botao(string texto)
        {
            var botao = new Button
            {
                Text = texto, AutoSize = true, MinimumSize = new Size(100, 36),
                FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(35, 35, 35),
                ForeColor = ForeColor, Margin = new Padding(0, 0, 10, 0), Padding = new Padding(8, 3, 8, 3)
            };
            botao.FlatAppearance.BorderColor = Color.FromArgb(72, 72, 72);
            acoes.Controls.Add(botao);
            return botao;
        }
        var expandir = Botao("Ver &detalhes técnicos");
        expandir.Click += (_, _) =>
        {
            caixaDetalhes.Visible = !caixaDetalhes.Visible;
            estrutura.RowStyles[1].Height = caixaDetalhes.Visible ? 160 : 0;
            expandir.Text = caixaDetalhes.Visible ? "Ocultar &detalhes técnicos" : "Ver &detalhes técnicos";
        };
        var copiar = Botao("&Copiar detalhes");
        copiar.Click += (_, _) =>
        {
            try { Clipboard.SetText(detalhes); copiar.Text = "Detalhes copiados"; }
            catch (System.Runtime.InteropServices.ExternalException) { copiar.Text = "Tente copiar novamente"; }
        };
        var fechar = Botao("&Fechar");
        fechar.BackColor = Color.FromArgb(45, 117, 235);
        fechar.DialogResult = DialogResult.OK;
        AcceptButton = fechar;
        CancelButton = fechar;
    }

    public static void Exibir(IWin32Window owner, Exception erro, string provedor, string modelo, bool arquivosPreservados = false, string? arquivo = null)
    {
        using var modal = new ModalErroIA(erro, provedor, modelo, arquivosPreservados, arquivo);
        modal.ShowDialog(owner);
    }
}
