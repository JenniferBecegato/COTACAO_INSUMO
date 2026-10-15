using System.Globalization;

namespace COTACAO_INSUMO;

internal static class LocalizarNaCotacao
{
    public static void Abrir(Form dono, DataGridView tabela)
    {
        if (!tabela.EndEdit()) return;

        using var janela = new Form
        {
            Text = "Localizar na cotação",
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MinimizeBox = false,
            MaximizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(470, 165),
            AutoScaleMode = AutoScaleMode.Dpi
        };
        var rotulo = new Label { Text = "Localizar texto:", AutoSize = true, Location = new Point(16, 16) };
        var termo = new TextBox { Location = new Point(16, 40), Width = 438, TabIndex = 0 };
        var resultado = new Label { Location = new Point(16, 74), Size = new Size(438, 32) };
        var proximo = new Button { Text = "Localizar próxima", Location = new Point(204, 118), Size = new Size(145, 30), TabIndex = 1 };
        var fechar = new Button { Text = "Fechar", Location = new Point(359, 118), Size = new Size(95, 30), DialogResult = DialogResult.Cancel, TabIndex = 2 };
        janela.Controls.AddRange([rotulo, termo, resultado, proximo, fechar]);
        janela.AcceptButton = proximo;
        janela.CancelButton = fechar;

        bool primeiraBusca = true;
        termo.TextChanged += (_, _) => { primeiraBusca = true; resultado.Text = ""; };
        proximo.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(termo.Text))
            {
                resultado.Text = "Digite um termo para localizar.";
                termo.Focus();
                return;
            }

            var colunas = tabela.Columns.Cast<DataGridViewColumn>()
                .Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToArray();
            var celulas = tabela.Rows.Cast<DataGridViewRow>()
                .Where(r => r.Visible && !r.IsNewRow)
                .SelectMany(r => colunas.Select(c => r.Cells[c.Index])).ToArray();
            int atual = Array.IndexOf(celulas, tabela.CurrentCell);
            int inicio = atual < 0 ? 0 : (atual + (primeiraBusca ? 0 : 1)) % Math.Max(1, celulas.Length);
            for (int passo = 0; passo < celulas.Length; passo++)
            {
                var celula = celulas[(inicio + passo) % celulas.Length];
                string texto = Convert.ToString(celula.FormattedValue, CultureInfo.CurrentCulture) ?? "";
                if (CultureInfo.CurrentCulture.CompareInfo.IndexOf(texto, termo.Text,
                    CompareOptions.IgnoreCase) < 0) continue;

                tabela.ClearSelection();
                tabela.CurrentCell = celula;
                celula.Selected = true;
                if (celula.OwningRow is { Frozen: false }) tabela.FirstDisplayedScrollingRowIndex = celula.RowIndex;
                if (celula.OwningColumn is { Frozen: false }) tabela.FirstDisplayedScrollingColumnIndex = celula.ColumnIndex;
                resultado.Text = $"Encontrado na linha {celula.RowIndex + 1}, coluna {celula.OwningColumn?.HeaderText}.";
                primeiraBusca = false;
                return;
            }
            resultado.Text = "Nenhuma ocorrência encontrada.";
        };
        janela.Shown += (_, _) => termo.Focus();
        janela.ShowDialog(dono);
        tabela.Focus();
    }
}
