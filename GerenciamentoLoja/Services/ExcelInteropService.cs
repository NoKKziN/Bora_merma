using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
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

        _workbook = ComRetry(() => (object)_app!.Workbooks.Open(caminhoArquivo));
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
        return ComRetry(() =>
        {
            var nomes = new List<string>();
            foreach (var aba in _workbook!.Worksheets)
            {
                nomes.Add((string)aba.Name);
                Marshal.ReleaseComObject((object)aba);
            }
            return nomes;
        });
    }

    public List<string> ObterCabecalhos(string nomeAba, int linhaCabecalho)
    {
        GarantirConectado();
        return ComRetry(() => ObterCabecalhosInterno(nomeAba, linhaCabecalho));
    }

    private List<string> ObterCabecalhosInterno(string nomeAba, int linhaCabecalho)
    {
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

    public List<ItemEstoque> LerItens(PlanilhaMapeamento mapeamento)
    {
        GarantirConectado();
        return ComRetry(() => LerItensInterno(mapeamento));
    }

    private List<ItemEstoque> LerItensInterno(PlanilhaMapeamento mapeamento)
    {
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

                var itens = new List<ItemEstoque>();
                if (ultimaLinha < primeiraLinha)
                {
                    return itens;
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
                        var sku = ObterTexto(valores, linhaRelativa, mapeamento, CamposItem.Sku, primeiraColuna);
                        if (string.IsNullOrWhiteSpace(sku))
                        {
                            continue;
                        }

                        var item = new ItemEstoque
                        {
                            LinhaPlanilha = linha,
                            Sku = sku,
                            Produto = ObterTexto(valores, linhaRelativa, mapeamento, CamposItem.Produto, primeiraColuna),
                            Categoria = ObterTexto(valores, linhaRelativa, mapeamento, CamposItem.Categoria, primeiraColuna),
                            Status = ObterTexto(valores, linhaRelativa, mapeamento, CamposItem.Status, primeiraColuna),
                            Cliente = ObterTextoOpcional(valores, linhaRelativa, mapeamento, CamposItem.Cliente, primeiraColuna),
                            ValorCusto = (decimal)ObterNumero(valores, linhaRelativa, mapeamento, CamposItem.ValorCusto, primeiraColuna),
                            PrecoAVista = (decimal)ObterNumero(valores, linhaRelativa, mapeamento, CamposItem.PrecoAVista, primeiraColuna),
                            PrecoCartao = (decimal)ObterNumero(valores, linhaRelativa, mapeamento, CamposItem.PrecoCartao, primeiraColuna),
                            DataEntrada = ObterData(valores, linhaRelativa, mapeamento, CamposItem.DataEntrada, primeiraColuna),
                            DataVenda = ObterDataOpcional(valores, linhaRelativa, mapeamento, CamposItem.DataVenda, primeiraColuna),
                            ValorRecebido = ObterNumeroOpcional(valores, linhaRelativa, mapeamento, CamposItem.ValorRecebido, primeiraColuna)
                        };
                        itens.Add(item);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject((object)bloco);
                }

                return itens;
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

    public void AdicionarItem(PlanilhaMapeamento mapeamento, ItemEstoque item)
    {
        GarantirConectado();
        ComRetry(() =>
        {
            var worksheet = ObterWorksheet(mapeamento.NomeAba);
            try
            {
                int novaLinha = LocalizarProximaLinhaLivre(worksheet, mapeamento);

                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.Sku, item.Sku);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.Produto, item.Produto);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.Categoria, item.Categoria);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.Status, item.Status);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.DataEntrada, item.DataEntrada);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.ValorCusto, item.ValorCusto);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.PrecoAVista, item.PrecoAVista);
                EscreverCelula(worksheet, novaLinha, mapeamento, CamposItem.PrecoCartao, item.PrecoCartao);

                _workbook!.Save();
            }
            finally
            {
                Marshal.ReleaseComObject((object)worksheet);
            }
        });
    }

    public void AtualizarItem(PlanilhaMapeamento mapeamento, int linhaPlanilha, IReadOnlyDictionary<string, object?> valores)
    {
        GarantirConectado();
        ComRetry(() =>
        {
            var worksheet = ObterWorksheet(mapeamento.NomeAba);
            try
            {
                foreach (var (campo, valor) in valores)
                {
                    EscreverCelula(worksheet, linhaPlanilha, mapeamento, campo, valor);
                }

                _workbook!.Save();
            }
            finally
            {
                Marshal.ReleaseComObject((object)worksheet);
            }
        });
    }

    private static int LocalizarProximaLinhaLivre(dynamic worksheet, PlanilhaMapeamento mapeamento)
    {
        int colunaAncora = mapeamento.Colunas[CamposItem.Sku];
        var ultimaCelula = worksheet.Cells[worksheet.Rows.Count, colunaAncora];
        try
        {
            var celulaFinal = ultimaCelula.End(XlUp);
            try
            {
                int ultimaLinhaUsada = (int)celulaFinal.Row;
                return Math.Max(ultimaLinhaUsada + 1, mapeamento.LinhaCabecalho + 1);
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

    private static string? ObterTextoOpcional(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        var texto = ObterTexto(valores, linhaRelativa, mapeamento, campo, primeiraColuna);
        return string.IsNullOrWhiteSpace(texto) ? null : texto;
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

    private static decimal? ObterNumeroOpcional(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        var valor = ObterValorBruto(valores, linhaRelativa, mapeamento, campo, primeiraColuna);
        return valor switch
        {
            null => null,
            double d => (decimal)d,
            int i => i,
            _ => double.TryParse(valor.ToString(), out var parsed) ? (decimal)parsed : null
        };
    }

    private static DateTime ObterData(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        var valor = ObterValorBruto(valores, linhaRelativa, mapeamento, campo, primeiraColuna);
        return valor switch
        {
            null => default,
            double serial => DateTime.FromOADate(serial),
            DateTime dt => dt,
            _ => DateTime.TryParse(valor.ToString(), out var parsed) ? parsed : default
        };
    }

    private static DateTime? ObterDataOpcional(object?[,] valores, int linhaRelativa, PlanilhaMapeamento mapeamento, string campo, int primeiraColuna)
    {
        var valor = ObterValorBruto(valores, linhaRelativa, mapeamento, campo, primeiraColuna);
        return valor switch
        {
            null => null,
            double serial => DateTime.FromOADate(serial),
            DateTime dt => dt,
            _ => DateTime.TryParse(valor.ToString(), out var parsed) ? parsed : null
        };
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

    // O Excel rejeita chamadas COM reentrantes enquanto está ocupado (recalculando,
    // exibindo um diálogo, ou sendo usado pelo cliente na tela) com RPC_E_CALL_REJECTED /
    // RPC_E_SERVERCALL_RETRYLATER. Isso é esperado em automação real, não um bug — a
    // prática padrão é tentar de novo algumas vezes com um pequeno intervalo.
    private const uint RpcECallRejected = 0x80010001;
    private const uint RpcEServerCallRetryLater = 0x8001010A;

    private static T ComRetry<T>(Func<T> acao, int tentativas = 6, int esperaMs = 500)
    {
        for (int tentativa = 1; tentativa <= tentativas; tentativa++)
        {
            try
            {
                return acao();
            }
            catch (COMException ex) when (tentativa < tentativas && EhErroTransitorio(ex))
            {
                Thread.Sleep(esperaMs);
            }
        }
        return acao();
    }

    private static void ComRetry(Action acao, int tentativas = 6, int esperaMs = 500)
    {
        ComRetry<object?>(() =>
        {
            acao();
            return null;
        }, tentativas, esperaMs);
    }

    private static bool EhErroTransitorio(COMException ex)
    {
        var codigo = unchecked((uint)ex.ErrorCode);
        return codigo == RpcECallRejected || codigo == RpcEServerCallRetryLater;
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
