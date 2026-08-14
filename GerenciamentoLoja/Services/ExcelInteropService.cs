using System.IO;
using System.Runtime.InteropServices;
using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

// Mantém uma planilha do Excel aberta ao vivo (via COM Interop com late binding) e
// lê/grava diretamente nela, para que qualquer alteração feita pelo app ou pelo
// cliente na janela do Excel apareça em tempo real dos dois lados.
//
// Usamos "dynamic" (late binding) em vez das Primary Interop Assemblies fortemente
// tipadas: isso evita depender do assembly "office" (Microsoft.Office.Core), que só
// existe registrado na GAC de máquinas com o Office instalado via MSI tradicional e
// não é redistribuível como pacote NuGet — o que quebraria a instalação no cliente.
public class ExcelInteropService : IExcelWorkbookService
{
    // Valor de Excel.XlDirection.xlUp, usado em Range.End sem precisar do enum tipado.
    private const int XlUp = -4162;

    private dynamic? _app;
    private dynamic? _workbook;
    private bool _appCriadoPorNos;

    public bool EstaConectado => _workbook != null;
    public string? ArquivoAtual { get; private set; }

    public event EventHandler? ConexaoAlterada;

    public void AbrirPlanilha(string caminhoArquivo)
    {
        Desconectar();

        caminhoArquivo = Path.GetFullPath(caminhoArquivo);

        var appExistente = TentarObterExcelAtivo();
        if (appExistente != null)
        {
            var workbookExistente = ProcurarWorkbookAberto(appExistente, caminhoArquivo);
            if (workbookExistente != null)
            {
                _app = appExistente;
                _workbook = workbookExistente;
                _appCriadoPorNos = false;
                ArquivoAtual = caminhoArquivo;
                ConexaoAlterada?.Invoke(this, EventArgs.Empty);
                return;
            }

            // Reaproveita a instância ativa do Excel para abrir o arquivo do cliente,
            // em vez de subir uma segunda instância separada.
            _app = appExistente;
            _appCriadoPorNos = false;
        }
        else
        {
            var tipoExcel = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("O Microsoft Excel não está instalado nesta máquina.");
            _app = Activator.CreateInstance(tipoExcel)!;
            _app.Visible = true;
            _app.DisplayAlerts = false;
            _appCriadoPorNos = true;
        }

        _workbook = _app!.Workbooks.Open(caminhoArquivo);
        _app.Visible = true;
        ArquivoAtual = caminhoArquivo;
        ConexaoAlterada?.Invoke(this, EventArgs.Empty);
    }

    public void Desconectar()
    {
        if (_workbook != null)
        {
            Marshal.ReleaseComObject((object)_workbook);
            _workbook = null;
        }

        // Nunca fechamos (Quit) o Excel automaticamente: o cliente pode
        // continuar trabalhando na planilha mesmo com o app desconectado.
        if (_app != null)
        {
            if (_appCriadoPorNos)
            {
                Marshal.ReleaseComObject((object)_app);
            }
            _app = null;
        }

        ArquivoAtual = null;
        ConexaoAlterada?.Invoke(this, EventArgs.Empty);
    }

    public List<string> ObterNomesAbas()
    {
        GarantirConectado();
        var nomes = new List<string>();
        foreach (var aba in _workbook!.Worksheets)
        {
            nomes.Add((string)aba.Name);
            Marshal.ReleaseComObject((object)aba);
        }
        return nomes;
    }

    public List<string> ObterCabecalhos(string nomeAba, int linhaCabecalho)
    {
        GarantirConectado();
        var worksheet = ObterWorksheet(nomeAba);
        try
        {
            var usedRange = worksheet.UsedRange;
            try
            {
                int primeiraColuna = (int)usedRange.Column;
                int ultimaColuna = primeiraColuna + (int)usedRange.Columns.Count - 1;

                var cabecalhos = new List<string>();
                for (int coluna = primeiraColuna; coluna <= ultimaColuna; coluna++)
                {
                    var celula = worksheet.Cells[linhaCabecalho, coluna];
                    try
                    {
                        object? valor = celula.Value2;
                        cabecalhos.Add(valor?.ToString() ?? string.Empty);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject((object)celula);
                    }
                }
                return cabecalhos;
            }
            finally
            {
                Marshal.ReleaseComObject((object)usedRange);
            }
        }
        finally
        {
            Marshal.ReleaseComObject((object)worksheet);
        }
    }

    public List<Movimento> LerMovimentos(PlanilhaMapeamento mapeamento)
    {
        GarantirConectado();
        var worksheet = ObterWorksheet(mapeamento.NomeAba);
        try
        {
            var usedRange = worksheet.UsedRange;
            try
            {
                int primeiraLinha = mapeamento.LinhaCabecalho + 1;
                int ultimaLinha = (int)usedRange.Row + (int)usedRange.Rows.Count - 1;
                int primeiraColuna = (int)usedRange.Column;
                int ultimaColuna = primeiraColuna + (int)usedRange.Columns.Count - 1;

                var movimentos = new List<Movimento>();
                if (ultimaLinha < primeiraLinha)
                {
                    return movimentos;
                }

                var bloco = worksheet.Range[
                    worksheet.Cells[primeiraLinha, primeiraColuna],
                    worksheet.Cells[ultimaLinha, ultimaColuna]];
                try
                {
                    var valores = ExtrairMatriz(bloco, primeiraLinha, ultimaLinha, primeiraColuna, ultimaColuna);

                    for (int linha = primeiraLinha; linha <= ultimaLinha; linha++)
                    {
                        int linhaRelativa = linha - primeiraLinha + 1;
                        var sku = ObterTexto(valores, linhaRelativa, mapeamento, CamposMovimento.Sku, primeiraColuna);
                        if (string.IsNullOrWhiteSpace(sku))
                        {
                            continue;
                        }

                        var movimento = new Movimento
                        {
                            LinhaPlanilha = linha,
                            Sku = sku,
                            Produto = ObterTexto(valores, linhaRelativa, mapeamento, CamposMovimento.Produto, primeiraColuna),
                            Categoria = ObterTexto(valores, linhaRelativa, mapeamento, CamposMovimento.Categoria, primeiraColuna),
                            Cliente = ObterTexto(valores, linhaRelativa, mapeamento, CamposMovimento.Cliente, primeiraColuna),
                            Motivo = ObterTexto(valores, linhaRelativa, mapeamento, CamposMovimento.Motivo, primeiraColuna),
                            Data = ObterData(valores, linhaRelativa, mapeamento, primeiraColuna),
                            Tipo = ObterTipo(valores, linhaRelativa, mapeamento, primeiraColuna),
                            Quantidade = (int)ObterNumero(valores, linhaRelativa, mapeamento, CamposMovimento.Quantidade, primeiraColuna),
                            ValorUnitario = (decimal)ObterNumero(valores, linhaRelativa, mapeamento, CamposMovimento.ValorUnitario, primeiraColuna)
                        };
                        movimentos.Add(movimento);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject((object)bloco);
                }

                return movimentos;
            }
            finally
            {
                Marshal.ReleaseComObject((object)usedRange);
            }
        }
        finally
        {
            Marshal.ReleaseComObject((object)worksheet);
        }
    }

    public void AdicionarMovimento(PlanilhaMapeamento mapeamento, Movimento movimento)
    {
        GarantirConectado();
        var worksheet = ObterWorksheet(mapeamento.NomeAba);
        try
        {
            int colunaAncora = mapeamento.Colunas[CamposMovimento.Sku];
            int ultimaLinhaUsada;
            var ultimaCelula = worksheet.Cells[worksheet.Rows.Count, colunaAncora];
            try
            {
                var celulaFinal = ultimaCelula.End(XlUp);
                try
                {
                    ultimaLinhaUsada = (int)celulaFinal.Row;
                }
                finally
                {
                    Marshal.ReleaseComObject((object)celulaFinal);
                }
            }
            finally
            {
                Marshal.ReleaseComObject((object)ultimaCelula);
            }

            int novaLinha = Math.Max(ultimaLinhaUsada + 1, mapeamento.LinhaCabecalho + 1);

            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Data, movimento.Data);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Sku, movimento.Sku);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Produto, movimento.Produto);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Categoria, movimento.Categoria);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Tipo,
                movimento.Tipo == TipoMovimento.Entrada ? "Entrada" : "Saida");
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Quantidade, movimento.Quantidade);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.ValorUnitario, movimento.ValorUnitario);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Cliente, movimento.Cliente);
            EscreverCelula(worksheet, novaLinha, mapeamento, CamposMovimento.Motivo, movimento.Motivo);

            _workbook!.Save();
        }
        finally
        {
            Marshal.ReleaseComObject((object)worksheet);
        }
    }

    private static void EscreverCelula(dynamic worksheet, int linha, PlanilhaMapeamento mapeamento, string campo, object? valor)
    {
        if (!mapeamento.Colunas.TryGetValue(campo, out var coluna))
        {
            return;
        }

        var celula = worksheet.Cells[linha, coluna];
        try
        {
            celula.Value2 = valor ?? string.Empty;
        }
        finally
        {
            Marshal.ReleaseComObject((object)celula);
        }
    }

    // bloco.Value2 retorna um array 1-based (linha, coluna) relativo ao próprio bloco lido,
    // independente da posição real na planilha. Para uma única célula retorna um escalar.
    private static object?[,] ExtrairMatriz(dynamic bloco, int primeiraLinha, int ultimaLinha, int primeiraColuna, int ultimaColuna)
    {
        if (primeiraLinha == ultimaLinha && primeiraColuna == ultimaColuna)
        {
            var unico = new object?[2, 2];
            unico[1, 1] = bloco.Value2;
            return unico;
        }

        return (object[,])bloco.Value2;
    }

    // linhaRelativa é 1-based dentro do bloco lido (ExtrairMatriz); coluna vem do mapeamento
    // em índice absoluto da planilha e precisa ser deslocada para o índice relativo do bloco.
    private static string ObterTexto(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        var valor = ObterValorBruto(valores, linhaRelativa, mapeamento, campo, primeiraColuna);
        return valor?.ToString()?.Trim() ?? string.Empty;
    }

    private static double ObterNumero(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        var valor = ObterValorBruto(valores, linhaRelativa, mapeamento, campo, primeiraColuna);
        return valor switch
        {
            null => 0,
            double d => d,
            int i => i,
            _ => double.TryParse(valor.ToString(), out var parsed) ? parsed : 0
        };
    }

    private static DateTime ObterData(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, int primeiraColuna)
    {
        var valor = ObterValorBruto(valores, linhaRelativa, mapeamento, CamposMovimento.Data, primeiraColuna);
        return valor switch
        {
            null => default,
            double serial => DateTime.FromOADate(serial),
            DateTime dt => dt,
            _ => DateTime.TryParse(valor.ToString(), out var parsed) ? parsed : default
        };
    }

    private static TipoMovimento ObterTipo(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, int primeiraColuna)
    {
        var texto = ObterTexto(valores, linhaRelativa, mapeamento, CamposMovimento.Tipo, primeiraColuna);
        return texto.StartsWith("E", StringComparison.OrdinalIgnoreCase) ? TipoMovimento.Entrada : TipoMovimento.Saida;
    }

    private static object? ObterValorBruto(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        if (!mapeamento.Colunas.TryGetValue(campo, out var colunaAbsoluta))
        {
            return null;
        }

        int colunaRelativa = colunaAbsoluta - primeiraColuna + 1;
        return valores[linhaRelativa, colunaRelativa];
    }

    private dynamic ObterWorksheet(string nomeAba)
    {
        return _workbook!.Worksheets[nomeAba];
    }

    private void GarantirConectado()
    {
        if (_workbook == null)
        {
            throw new InvalidOperationException("Nenhuma planilha conectada.");
        }
    }

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    // Não existe equivalente a Marshal.GetActiveObject em .NET moderno (só em .NET
    // Framework), então buscamos a instância ativa do Excel via P/Invoke direto.
    private static dynamic? TentarObterExcelAtivo()
    {
        try
        {
            if (CLSIDFromProgID("Excel.Application", out var clsid) != 0)
            {
                return null;
            }

            if (GetActiveObject(ref clsid, IntPtr.Zero, out var instancia) != 0 || instancia == null)
            {
                return null;
            }

            return instancia;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static dynamic? ProcurarWorkbookAberto(dynamic app, string caminhoArquivo)
    {
        foreach (var workbook in app.Workbooks)
        {
            if (string.Equals((string)workbook.FullName, caminhoArquivo, StringComparison.OrdinalIgnoreCase))
            {
                return workbook;
            }
            Marshal.ReleaseComObject((object)workbook);
        }
        return null;
    }

    public void Dispose()
    {
        Desconectar();
        GC.SuppressFinalize(this);
    }
}
