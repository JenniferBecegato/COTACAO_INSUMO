using System.Text;
namespace COTACAO_INSUMO;

internal sealed class HistoricoTabela
{
    internal sealed class Estado
    {
        public DataGridViewColumn[] Colunas = [];
        public DataGridViewRow[] Linhas = [];
        public string Assinatura = "";
    }
    readonly DataGridView grid;
    readonly Action<Action, Action> registrar;
    readonly Action aoRestaurar;
    readonly Action ativar;
    Estado anterior;
    bool pendente;
    public bool Restaurando { get; private set; }
    public HistoricoTabela(DataGridView grid, Action<Action, Action> registrar, Action ativar, Action aoRestaurar)
    {
        this.grid = grid; this.registrar = registrar; this.ativar = ativar; this.aoRestaurar = aoRestaurar;
        anterior = Capturar();
    }
    Estado Capturar()
    {
        var estado = new Estado
        {
            Colunas = grid.Columns.Cast<DataGridViewColumn>().Select(c => { var copia = (DataGridViewColumn)c.Clone(); copia.DisplayIndex = c.DisplayIndex; copia.HeaderCell.ContextMenuStrip = c.HeaderCell.ContextMenuStrip; return copia; }).ToArray(),
            Linhas = grid.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r =>
            {
                var copia = (DataGridViewRow)r.Clone();
                for (int i = 0; i < r.Cells.Count; i++) { copia.Cells[i].Value = r.Cells[i].Value; copia.Cells[i].Style = r.Cells[i].Style.Clone(); }
                return copia;
            }).ToArray()
        };
        var texto = new StringBuilder();
        foreach (var c in estado.Colunas) texto.Append(c.Name).Append('|').Append(c.HeaderText).Append('|').Append(c.DisplayIndex).Append('|').Append(c.Width).Append('\n');
        foreach (var r in estado.Linhas)
        {
            texto.Append(r.Height).Append(':');
            foreach (DataGridViewCell c in r.Cells) texto.Append(c.Value?.ToString()?.Length ?? -1).Append(':').Append(c.Value).Append('|').Append(c.Style.BackColor.ToArgb()).Append('|').Append(c.Style.ForeColor.ToArgb()).Append(';');
            texto.Append('\n');
        }
        estado.Assinatura = texto.ToString();
        return estado;
    }
    public void Agendar()
    {
        if (Restaurando || pendente || grid.IsDisposed || !grid.IsHandleCreated) return;
        pendente = true;
        grid.BeginInvoke((Action)(() => { if (pendente) Confirmar(); }));
    }
    public void Confirmar()
    {
        pendente = false;
        if (Restaurando || grid.IsDisposed) return;
        var atual = Capturar();
        if (atual.Assinatura == anterior.Assinatura) return;
        var salvar = anterior;
        anterior = atual;
        registrar(() => Restaurar(salvar), () => Restaurar(atual));
    }
    void Restaurar(Estado estado)
    {
        ativar();
        Restaurando = true;
        grid.SuspendLayout();
        try
        {
            grid.CancelEdit();
            grid.Rows.Clear(); grid.Columns.Clear();
            foreach (var c in estado.Colunas) { var copia = (DataGridViewColumn)c.Clone(); copia.DisplayIndex = -1; copia.HeaderCell.ContextMenuStrip = c.HeaderCell.ContextMenuStrip; grid.Columns.Add(copia); }
            foreach (var c in estado.Colunas.OrderBy(c => c.DisplayIndex)) grid.Columns[c.Name]!.DisplayIndex = c.DisplayIndex;
            foreach (var r in estado.Linhas)
            {
                var copia = (DataGridViewRow)r.Clone();
                for (int i = 0; i < r.Cells.Count; i++) { copia.Cells[i].Value = r.Cells[i].Value; copia.Cells[i].Style = r.Cells[i].Style.Clone(); }
                grid.Rows.Add(copia);
            }
            anterior = Capturar();
        }
        finally { grid.ResumeLayout(); Restaurando = false; }
        aoRestaurar();
    }
    public void Sincronizar() { if (!pendente && !Restaurando) anterior = Capturar(); }
}
