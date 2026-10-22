using System.Text;

namespace COTACAO_INSUMO;

internal static class AreaTransferenciaCelulas
{
    public static void Copiar(DataGridView tabela, bool recortar)
    {
        var selecionadas = tabela.SelectedCells.Cast<DataGridViewCell>()
            .Where(c => c.Visible && c.OwningRow?.IsNewRow == false).ToHashSet();
        if (selecionadas.Count == 0) return;
        var colunas = tabela.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible)
            .OrderBy(c => c.DisplayIndex).ToList();
        int inicio = selecionadas.Min(c => colunas.FindIndex(coluna => coluna.Index == c.ColumnIndex));
        int fim = selecionadas.Max(c => colunas.FindIndex(coluna => coluna.Index == c.ColumnIndex));
        var linhas = new List<string>();
        for (int r = selecionadas.Min(c => c.RowIndex); r <= selecionadas.Max(c => c.RowIndex); r++)
        {
            if (!tabela.Rows[r].Visible) continue;
            linhas.Add(string.Join("\t", colunas.Skip(inicio).Take(fim - inicio + 1).Select(coluna =>
            {
                var celula = tabela.Rows[r].Cells[coluna.Index];
                string valor = selecionadas.Contains(celula) ? celula.FormattedValue?.ToString() ?? "" : "";
                return valor.IndexOfAny(['\t', '\r', '\n', '"']) >= 0
                    ? "\"" + valor.Replace("\"", "\"\"") + "\"" : valor;
            })));
        }
        // Só limpar a origem depois de copiar com sucesso.
        Clipboard.SetText(string.Join("\r\n", linhas));
        if (recortar)
            foreach (var celula in selecionadas)
                if (!celula.ReadOnly) celula.Value = null;
    }

    public static void Colar(DataGridView tabela)
    {
        if (tabela.ReadOnly || tabela.CurrentCell == null || !Clipboard.ContainsText()) return;
        var valores = LerTexto(Clipboard.GetText());
        var colunas = tabela.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible)
            .OrderBy(c => c.DisplayIndex).ToList();
        int inicio = colunas.FindIndex(coluna => coluna.Index == tabela.CurrentCell.ColumnIndex);
        if (inicio < 0) return;
        var linhas = tabela.Rows.Cast<DataGridViewRow>()
            .Where(r => r.Visible && !r.IsNewRow && r.Index >= tabela.CurrentCell.RowIndex).ToArray();
        for (int r = 0; r < valores.Count && r < linhas.Length; r++)
            for (int c = 0; c < valores[r].Count && inicio + c < colunas.Count; c++)
            {
                var destino = linhas[r].Cells[colunas[inicio + c].Index];
                if (!destino.ReadOnly) destino.Value = valores[r][c];
            }
    }

    internal static List<List<string>> LerTexto(string texto)
    {
        var resultado = new List<List<string>>();
        var linha = new List<string>();
        var campo = new StringBuilder();
        bool aspas = false;
        for (int i = 0; i < texto.Length; i++)
        {
            char ch = texto[i];
            if (ch == '"' && (aspas || campo.Length == 0))
            {
                if (aspas && i + 1 < texto.Length && texto[i + 1] == '"') { campo.Append('"'); i++; }
                else aspas = !aspas;
            }
            else if (!aspas && (ch == '\t' || ch == '\r' || ch == '\n'))
            {
                linha.Add(campo.ToString()); campo.Clear();
                if (ch != '\t')
                {
                    if (ch == '\r' && i + 1 < texto.Length && texto[i + 1] == '\n') i++;
                    resultado.Add(linha); linha = new();
                }
            }
            else campo.Append(ch);
        }
        if (campo.Length > 0 || linha.Count > 0 || resultado.Count == 0)
        {
            linha.Add(campo.ToString()); resultado.Add(linha);
        }
        return resultado;
    }
}
