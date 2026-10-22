using System.Runtime.CompilerServices;

namespace COTACAO_INSUMO;

public partial class Form1
{
    private FiltroZoom? filtroZoom;
    private readonly ConditionalWeakTable<Control, EstadoZoom> estadosZoom = new();
    private sealed class EstadoZoom { public int Percentual = 100; }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        filtroZoom = new FiltroZoom(this);
        Application.AddMessageFilter(filtroZoom);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        if (filtroZoom != null) Application.RemoveMessageFilter(filtroZoom);
        base.OnFormClosed(e);
    }

    private void AplicarZoom(int direcao)
    {
        if (painelConteudo.Controls.Count == 0) return;
        Control pagina = painelConteudo.Controls[0];
        var estado = estadosZoom.GetOrCreateValue(pagina);
        int percentual = Math.Clamp(estado.Percentual + direcao * 10, 50, 200);
        if (percentual == estado.Percentual) return;
        float fator = percentual / (float)estado.Percentual;
        estado.Percentual = percentual;
        historicoTabelaAtivo?.Confirmar();
        pagina.SuspendLayout();
        try
        {
            // Capture explicit fonts before changing inherited parent fonts.
            var controles = new List<(Control Controle, Font Fonte)>();
            void Capturar(Control controle)
            {
                controles.Add((controle, controle.Font));
                foreach (Control filho in controle.Controls) Capturar(filho);
            }
            Capturar(pagina);
            pagina.Scale(new SizeF(fator, fator));
            foreach (var (controle, fonte) in controles)
            {
                controle.Font = new Font(fonte.FontFamily, fonte.Size * fator, fonte.Style, fonte.Unit);
                if (controle is not DataGridView grade) continue;
                grade.ScrollBars = ScrollBars.Both;
                grade.ColumnHeadersHeight = Math.Max(4, (int)Math.Round(grade.ColumnHeadersHeight * fator));
                grade.RowTemplate.Height = Math.Max(2, (int)Math.Round(grade.RowTemplate.Height * fator));
                foreach (DataGridViewRow linha in grade.Rows)
                    linha.Height = Math.Max(linha.MinimumHeight, (int)Math.Round(linha.Height * fator));
                foreach (DataGridViewColumn coluna in grade.Columns)
                    coluna.Width = Math.Max(coluna.MinimumWidth, (int)Math.Round(coluna.Width * fator));
                foreach (DataGridViewCellStyle estilo in new[] { grade.DefaultCellStyle,
                    grade.ColumnHeadersDefaultCellStyle, grade.RowHeadersDefaultCellStyle,
                    grade.AlternatingRowsDefaultCellStyle })
                {
                    if (estilo.Font is Font fonteEstilo)
                        estilo.Font = new Font(fonteEstilo.FontFamily, fonteEstilo.Size * fator,
                            fonteEstilo.Style, fonteEstilo.Unit);
                }
            }
            if (pagina is Panel painel && pagina.Dock == DockStyle.None)
                painel.AutoScroll = true;
        }
        finally
        {
            pagina.ResumeLayout(true);
            historicoTabelaAtivo?.Sincronizar();
        }
    }

    private sealed class FiltroZoom(Form1 janela) : IMessageFilter
    {
        public bool PreFilterMessage(ref Message mensagem)
        {
            const int WmMouseWheel = 0x020A;
            if (mensagem.Msg != WmMouseWheel || (ModifierKeys & Keys.Control) == 0 ||
                Form.ActiveForm != janela) return false;
            var alvo = Control.FromChildHandle(mensagem.HWnd);
            if (alvo?.FindForm() != janela) return false;
            int delta = unchecked((short)((mensagem.WParam.ToInt64() >> 16) & 0xffff));
            if (delta == 0) return false;
            janela.AplicarZoom(Math.Sign(delta));
            return true;
        }
    }
}
