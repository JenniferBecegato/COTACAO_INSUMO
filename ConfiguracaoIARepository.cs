using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace COTACAO_INSUMO;

public sealed record ConfiguracaoIA(string Provedor, string Modelo, string ChaveApi, int TimeoutSegundos);

// SQLite nativo do Windows. Nenhuma chave é armazenada em texto simples.
public sealed class ConfiguracaoIARepository
{
    private readonly string caminho;

    public ConfiguracaoIARepository(string? caminhoBanco = null)
    {
        caminho = caminhoBanco ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "COTACAO_INSUMO", "configuracoes.db");
    }

    public ConfiguracaoIA? Carregar(string? provedor = null)
    {
        using var banco = new Banco(caminho);
        using var consulta = banco.Preparar(provedor == null
            ? "SELECT provedor, modelo, chave_protegida, timeout_segundos FROM configuracao_ia WHERE selecionado = 1;"
            : "SELECT provedor, modelo, chave_protegida, timeout_segundos FROM configuracao_ia WHERE provedor = ?1;");
        if (provedor != null)
            consulta.Vincular(1, provedor);
        if (!consulta.Ler())
            return null;
        byte[] chave = ProtectedData.Unprotect(Convert.FromBase64String(consulta.Texto(2)),
            null, DataProtectionScope.CurrentUser);
        return new ConfiguracaoIA(consulta.Texto(0), consulta.Texto(1),
            Encoding.UTF8.GetString(chave), int.Parse(consulta.Texto(3)));
    }

    public void Salvar(ConfiguracaoIA configuracao)
    {
        if (string.IsNullOrWhiteSpace(configuracao.Provedor)
            || string.IsNullOrWhiteSpace(configuracao.Modelo)
            || string.IsNullOrWhiteSpace(configuracao.ChaveApi)
            || configuracao.TimeoutSegundos <= 0)
            throw new ArgumentException("Informe provedor, modelo, chave da API e tempo limite válidos.");
        string chaveProtegida = Convert.ToBase64String(ProtectedData.Protect(
            Encoding.UTF8.GetBytes(configuracao.ChaveApi), null, DataProtectionScope.CurrentUser));
        using var banco = new Banco(caminho);
        banco.Executar("BEGIN IMMEDIATE;");
        try
        {
            banco.Executar("UPDATE configuracao_ia SET selecionado = 0;");
            using var comando = banco.Preparar("""
                INSERT INTO configuracao_ia (provedor, modelo, chave_protegida, timeout_segundos, selecionado)
                VALUES (?1, ?2, ?3, ?4, 1)
                ON CONFLICT(provedor) DO UPDATE SET modelo = excluded.modelo,
                    chave_protegida = excluded.chave_protegida,
                    timeout_segundos = excluded.timeout_segundos, selecionado = 1;
                """);
            comando.Vincular(1, configuracao.Provedor);
            comando.Vincular(2, configuracao.Modelo);
            comando.Vincular(3, chaveProtegida);
            comando.Vincular(4, configuracao.TimeoutSegundos.ToString());
            comando.Ler();
            banco.Executar("COMMIT;");
        }
        catch
        {
            banco.Executar("ROLLBACK;");
            throw;
        }
    }

    public Action CapturarRestauracao()
    {
        var registros = new List<string[]>();
        using (var banco = new Banco(caminho))
        using (var consulta = banco.Preparar("SELECT provedor, modelo, chave_protegida, timeout_segundos, selecionado FROM configuracao_ia;"))
            while (consulta.Ler()) registros.Add(Enumerable.Range(0, 5).Select(consulta.Texto).ToArray());
        return () =>
        {
            using var banco = new Banco(caminho);
            banco.Executar("BEGIN IMMEDIATE;");
            try
            {
                banco.Executar("DELETE FROM configuracao_ia;");
                foreach (var registro in registros)
                {
                    using var comando = banco.Preparar("INSERT INTO configuracao_ia (provedor, modelo, chave_protegida, timeout_segundos, selecionado) VALUES (?1, ?2, ?3, ?4, ?5);");
                    for (int i = 0; i < registro.Length; i++) comando.Vincular(i + 1, registro[i]);
                    comando.Ler();
                }
                banco.Executar("COMMIT;");
            }
            catch { banco.Executar("ROLLBACK;"); throw; }
        };
    }

    internal sealed class Banco : IDisposable
    {
        public IntPtr Handle { get; private set; }
        public Banco(string caminho)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(caminho))!);
            int resultado = Native.sqlite3_open(caminho, out IntPtr handle);
            Handle = handle;
            if (resultado != 0)
            {
                Dispose();
                throw new IOException("Não foi possível abrir o banco de configurações.");
            }
            try
            {
                Native.sqlite3_busy_timeout(Handle, 5000);
                Executar("""
                    CREATE TABLE IF NOT EXISTS configuracao_ia (
                        provedor TEXT PRIMARY KEY NOT NULL,
                        modelo TEXT NOT NULL,
                        chave_protegida TEXT NOT NULL,
                        timeout_segundos INTEGER NOT NULL CHECK(timeout_segundos > 0),
                        selecionado INTEGER NOT NULL DEFAULT 0 CHECK(selecionado IN (0, 1))
                    );
                    """);
                Executar("CREATE UNIQUE INDEX IF NOT EXISTS ux_provedor_selecionado ON configuracao_ia(selecionado) WHERE selecionado = 1;");
                Executar("PRAGMA user_version = 1;");
            }
            catch { Dispose(); throw; }
        }
        public Consulta Preparar(string sql) => new Consulta(this, sql);
        public void Executar(string sql) { using var comando = Preparar(sql); comando.Ler(); }
        public void Verificar(int resultado)
        {
            if (resultado != 0)
                throw new IOException("Falha no banco de configurações: " +
                    Marshal.PtrToStringUTF8(Native.sqlite3_errmsg(Handle)));
        }
        public void Dispose()
        {
            if (Handle != IntPtr.Zero) { Native.sqlite3_close(Handle); Handle = IntPtr.Zero; }
        }
    }
    internal sealed class Consulta : IDisposable
    {
        private readonly Banco banco;
        private IntPtr handle;
        public Consulta(Banco banco, string sql)
        {
            this.banco = banco;
            banco.Verificar(Native.sqlite3_prepare_v2(banco.Handle, sql, -1, out handle, IntPtr.Zero));
        }
        public void Vincular(int indice, string valor) => banco.Verificar(
            Native.sqlite3_bind_text(handle, indice, valor, -1, new IntPtr(-1)));
        public bool Ler()
        {
            int resultado = Native.sqlite3_step(handle);
            if (resultado == 100) return true;
            if (resultado == 101) return false;
            banco.Verificar(resultado);
            return false;
        }
        public string Texto(int indice) => Marshal.PtrToStringUTF8(Native.sqlite3_column_text(handle, indice)) ?? "";
        public void Dispose()
        {
            if (handle != IntPtr.Zero) { Native.sqlite3_finalize(handle); handle = IntPtr.Zero; }
        }
    }
    private static class Native
    {
        private const string Dll = "winsqlite3.dll";
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_open([MarshalAs(UnmanagedType.LPUTF8Str)] string path, out IntPtr db);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_close(IntPtr db);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_busy_timeout(IntPtr db, int ms);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern IntPtr sqlite3_errmsg(IntPtr db);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_prepare_v2(IntPtr db, [MarshalAs(UnmanagedType.LPUTF8Str)] string sql, int length, out IntPtr statement, IntPtr tail);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_bind_text(IntPtr statement, int index, [MarshalAs(UnmanagedType.LPUTF8Str)] string value, int length, IntPtr destructor);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_step(IntPtr statement);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern IntPtr sqlite3_column_text(IntPtr statement, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int sqlite3_finalize(IntPtr statement);
    }
}
