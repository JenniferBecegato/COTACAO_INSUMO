using ClosedXML.Excel;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace COTACAO_INSUMO;

public static class ArquivoPlanilhaExcel
{
    private const string NamespacePlanilha = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static XLWorkbook Abrir(string caminho) => Abrir(File.ReadAllBytes(caminho));

    public static XLWorkbook Abrir(byte[] conteudo)
    {
        // ClosedXML retém o stream e volta a consultá-lo em SaveAs.
        // MemoryStream não possui recurso nativo; permanece vivo com o workbook.
        var stream = new MemoryStream(Normalizar(conteudo));
        return new XLWorkbook(stream);
    }

    // Aplicar apenas ao temporário, antes da substituição segura do arquivo original.
    public static void ValidarArquivoTemporario(string caminho)
    {
        byte[] original = File.ReadAllBytes(caminho);
        byte[] corrigido = Normalizar(original);
        if (!original.AsSpan().SequenceEqual(corrigido)) File.WriteAllBytes(caminho, corrigido);
    }

    public static byte[] Normalizar(byte[] conteudo)
    {
        using var memoria = new MemoryStream();
        memoria.Write(conteudo);
        bool alterado = false;
        using (var zip = new ZipArchive(memoria, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entrada in zip.Entries.ToArray())
            {
                if (!entrada.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                string xml;
                using (var leitor = new StreamReader(entrada.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    xml = leitor.ReadToEnd();
                try { ValidarXml(xml); }
                catch (XmlException)
                {
                    // Compatibilidade restrita às extensões SpreadsheetML conhecidas.
                    // Não inventar namespaces para outros prefixos ou descartar extensões.
                    if (entrada.FullName is not ("xl/workbook.xml" or "xl/styles.xml") ||
                        !Regex.IsMatch(xml, @"</?x:ext\b") || Regex.IsMatch(xml, @"\bxmlns:x\s*="))
                        throw;
                    var raiz = Regex.Match(xml, @"<(?:workbook|styleSheet)\b[^>]*>");
                    if (!raiz.Success || !raiz.Value.Contains($"xmlns=\"{NamespacePlanilha}\"", StringComparison.Ordinal))
                        throw;
                    string reparado = xml.Insert(raiz.Index + raiz.Length - 1, $" xmlns:x=\"{NamespacePlanilha}\"");
                    ValidarXml(reparado);
                    using var destino = entrada.Open();
                    destino.SetLength(0);
                    using var escritor = new StreamWriter(destino, new UTF8Encoding(false));
                    escritor.Write(reparado);
                    alterado = true;
                }
            }
        }
        return alterado ? memoria.ToArray() : conteudo;
    }

    private static void ValidarXml(string texto)
    {
        using var leitor = XmlReader.Create(new StringReader(texto), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null
        });
        while (leitor.Read()) { }
    }
}
