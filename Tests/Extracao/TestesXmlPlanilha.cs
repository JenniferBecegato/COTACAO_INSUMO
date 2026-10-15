using System.IO.Compression;
using System.Text;
using System.Xml;
using COTACAO_INSUMO;

internal static class TestesXmlPlanilha
{
    public static void Executar(string pasta, string? arquivoReal)
    {
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        byte[] Pacote(bool quebrado, bool desconhecido = false)
        {
            using var stream = new MemoryStream();
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                foreach (var (nome, raiz) in new[] { ("xl/workbook.xml", "workbook"), ("xl/styles.xml", "styleSheet") })
                {
                    using var escritor = new StreamWriter(zip.CreateEntry(nome).Open(), new UTF8Encoding(false));
                    string declaracao = quebrado ? "" : $" xmlns:x=\"{ns}\"";
                    string prefixo = desconhecido ? "outro" : "x";
                    escritor.Write($"<{raiz} xmlns=\"{ns}\"{declaracao}><extLst><{prefixo}:ext uri=\"teste\"/></extLst></{raiz}>");
                }
            }
            return stream.ToArray();
        }
        void Check(bool ok, string mensagem)
        {
            if (!ok) throw new Exception(mensagem);
        }
        var valido = Pacote(false);
        Check(valido.SequenceEqual(ArquivoPlanilhaExcel.Normalizar(valido)), "Pacotes válidos devem permanecer idênticos.");
        var reparado = ArquivoPlanilhaExcel.Normalizar(Pacote(true));
        using (var zip = new ZipArchive(new MemoryStream(reparado)))
            foreach (var entrada in zip.Entries)
            {
                using var stream = entrada.Open();
                var xml = new XmlDocument();
                xml.Load(stream);
                var ext = xml.GetElementsByTagName("ext", ns);
                Check(ext.Count == 1 && ext[0]!.Attributes!["uri"]!.Value == "teste", "Preservar extensões ao declarar namespace.");
            }
        bool rejeitou = false;
        try { ArquivoPlanilhaExcel.Normalizar(Pacote(true, true)); }
        catch (XmlException) { rejeitou = true; }
        Check(rejeitou, "Não tentar corrigir prefixos desconhecidos.");

        if (arquivoReal != null)
        {
            byte[] original = File.ReadAllBytes(arquivoReal);
            string copia = Path.Combine(pasta, "compatibilidade-xml.xlsx");
            File.WriteAllBytes(copia, original);
            var insumosAntes = LeitorPlanilhaInsumos.Ler(copia);
            int abas;
            using (var livro = ArquivoPlanilhaExcel.Abrir(copia)) abas = livro.Worksheets.Count;
            for (int rodada = 0; rodada < 2; rodada++)
            {
                string temp = copia + ".tmp.xlsx";
                using (var livro = ArquivoPlanilhaExcel.Abrir(copia)) livro.SaveAs(temp);
                ArquivoPlanilhaExcel.ValidarArquivoTemporario(temp);
                File.Move(temp, copia, overwrite: true);
                Check(insumosAntes.SequenceEqual(LeitorPlanilhaInsumos.Ler(copia)), "Preservar insumos após salvar/reabrir a planilha real.");
                using var reaberto = ArquivoPlanilhaExcel.Abrir(copia);
                Check(reaberto.Worksheets.Count == abas, "Preservar abas da planilha real.");
            }
            Check(original.SequenceEqual(File.ReadAllBytes(arquivoReal)), "Diagnóstico não deve modificar arquivo do usuário.");
            Console.WriteLine("PASS: planilha real com prefixo x inválido, duas gravações/reaberturas e insumos preservados.");
        }
        Console.WriteLine("PASS: compatibilidade XML, preservação de extensões e rejeição de corrupção desconhecida.");
    }
}
