using System.Drawing.Imaging;
using System.Text;

namespace COTACAO_INSUMO;

// A4 portrait for missing items; landscape with all columns for quotations.
internal static class RelatorioPdf
{
    public static void Salvar(string caminho, string titulo, string periodo,
        string[] cabecalhos, List<string[]> linhas, bool todasColunas = false)
    {
        if (linhas.Count == 0 || cabecalhos.Length == 0)
            throw new InvalidOperationException("Não há dados para exportar neste período.");
        var paginas = new List<byte[]>();
        int largura = todasColunas ? 1754 : 1240, altura = todasColunas ? 1240 : 1754;
        int margem = todasColunas ? 40 : 65;
        float tamanhoFonte = todasColunas ? Math.Min(10f, 105f / Math.Max(1, cabecalhos.Length - 1)) : 10f;
        using var fonte = new Font("Arial", tamanhoFonte);
        using var negrito = new Font("Arial", tamanhoFonte, FontStyle.Bold);
        using var fonteLegenda = new Font("Arial", 10);
        float padding = todasColunas ? 4 : 8;
        using var fonteTitulo = new Font("Arial", 18, FontStyle.Bold);
        using var formato = new StringFormat { Trimming = StringTrimming.None };
        // Repeat the input/material column in each group of supplier columns.
        int limiteColunas = cabecalhos.Length <= 5 ? 5 : 4;
        int porGrupo = limiteColunas - 1;
        int grupos = todasColunas || cabecalhos.Length <= limiteColunas ? 1
            : (int)Math.Ceiling((cabecalhos.Length - 1) / (double)porGrupo);
        for (int grupo = 0; grupo < grupos; grupo++)
        {
            int[] colunas = grupos == 1 ? Enumerable.Range(0, cabecalhos.Length).ToArray()
                : new[] { 0 }.Concat(Enumerable.Range(1 + grupo * porGrupo,
                    Math.Min(porGrupo, cabecalhos.Length - 1 - grupo * porGrupo))).ToArray();
            float disponivel = largura - margem * 2;
            int colunaInsumo = Array.FindIndex(colunas, c => cabecalhos[c].Equals("Insumo", StringComparison.OrdinalIgnoreCase));
            if (colunaInsumo < 0) colunaInsumo = 0;
            float maior = colunas.Length > 1 ? disponivel * (todasColunas && colunas.Length > 6 ? .22f : .40f) : disponivel;
            float restante = colunas.Length > 1 ? (disponivel - maior) / (colunas.Length - 1) : 0;
            float Largura(int c) => c == colunaInsumo ? maior : restante;
            int linha = 0;
            while (linha < linhas.Count)
            {
                using var bitmap = new Bitmap(largura, altura);
                bitmap.SetResolution(150, 150);
                using var g = Graphics.FromImage(bitmap);
                g.Clear(Color.White);
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                g.DrawString(titulo, fonteTitulo, Brushes.Black, margem, margem);
                g.DrawString(periodo + (grupos > 1 ? $" • Colunas {grupo + 1}/{grupos}" : ""),
                    fonteLegenda, Brushes.Black, margem, margem + 48);
                float y = margem + 85;
                float AlturaLinha(string[] valores, Font f)
                {
                    float h = todasColunas ? fonte.GetHeight(g) + padding * 2 : 42;
                    for (int c = 0; c < colunas.Length; c++)
                    {
                        string valor = colunas[c] < valores.Length ? valores[colunas[c]] : "";
                        h = Math.Max(h, g.MeasureString(valor, f,
                            (int)(Largura(c) - padding * 2), formato).Height + padding * 2);
                    }
                    return h;
                }
                void Desenhar(string[] valores, Font f, float h, bool header)
                {
                    float x = margem;
                    for (int c = 0; c < colunas.Length; c++)
                    {
                        float w = Largura(c);
                        if (header) g.FillRectangle(Brushes.LightGray, x, y, w, h);
                        g.DrawRectangle(Pens.Gray, x, y, w, h);
                        string valor = colunas[c] < valores.Length ? valores[colunas[c]] : "";
                        g.DrawString(valor, f, Brushes.Black, new RectangleF(x + padding, y + padding, w - padding * 2, h - padding * 2), formato);
                        x += w;
                    }
                    y += h;
                }
                Desenhar(cabecalhos, negrito, AlturaLinha(cabecalhos, negrito), true);
                int inicio = linha;
                while (linha < linhas.Count)
                {
                    float h = AlturaLinha(linhas[linha], fonte);
                    if (y + h > altura - margem - 35)
                    {
                        if (linha == inicio) throw new InvalidOperationException("Uma linha contém texto demais para uma página A4.");
                        break;
                    }
                    Desenhar(linhas[linha++], fonte, h, false);
                }
                g.DrawString($"Página {paginas.Count + 1} • Linhas {inicio + 1} a {linha}", fonteLegenda,
                    Brushes.Black, margem, altura - margem);
                using var memoria = new MemoryStream();
                bitmap.Save(memoria, ImageFormat.Jpeg);
                paginas.Add(memoria.ToArray());
            }
        }
        using var arquivo = new FileStream(caminho, FileMode.Create, FileAccess.Write);
        var offsets = new List<long> { 0 };
        void Texto(string texto) { arquivo.Write(Encoding.ASCII.GetBytes(texto)); }
        void Objeto(int id, string dados)
        {
            offsets.Add(arquivo.Position);
            Texto($"{id} 0 obj\n{dados}\nendobj\n");
        }
        Texto("%PDF-1.4\n");
        Objeto(1, "<< /Type /Catalog /Pages 2 0 R >>");
        Objeto(2, $"<< /Type /Pages /Count {paginas.Count} /Kids [" +
            string.Join(" ", Enumerable.Range(0, paginas.Count).Select(p => $"{3 + p * 3} 0 R")) + "] >>");
        int larguraPagina = todasColunas ? 842 : 595, alturaPagina = todasColunas ? 595 : 842;
        for (int p = 0; p < paginas.Count; p++)
        {
            int id = 3 + p * 3;
            Objeto(id, $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {larguraPagina} {alturaPagina}] /Resources << /XObject << /Im {id + 1} 0 R >> >> /Contents {id + 2} 0 R >>");
            offsets.Add(arquivo.Position);
            Texto($"{id + 1} 0 obj\n<< /Type /XObject /Subtype /Image /Width {largura} /Height {altura} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {paginas[p].Length} >>\nstream\n");
            arquivo.Write(paginas[p]);
            Texto("\nendstream\nendobj\n");
            string conteudo = $"q {larguraPagina} 0 0 {alturaPagina} 0 0 cm /Im Do Q\n";
            Objeto(id + 2, $"<< /Length {conteudo.Length} >>\nstream\n{conteudo}endstream");
        }
        long xref = arquivo.Position;
        Texto($"xref\n0 {offsets.Count}\n0000000000 65535 f \n");
        foreach (long offset in offsets.Skip(1)) Texto($"{offset:0000000000} 00000 n \n");
        Texto($"trailer\n<< /Size {offsets.Count} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
    }
}
