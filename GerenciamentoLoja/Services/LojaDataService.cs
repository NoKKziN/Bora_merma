using System.Globalization;
using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

public class LojaDataService : ILojaDataService
{
    private static readonly CultureInfo PtBr = CultureInfo.GetCultureInfo("pt-BR");

    private readonly IExcelWorkbookService _excel;

    public PlanilhaMapeamento? Mapeamento { get; private set; }
    public bool EstaConectado => _excel.EstaConectado && Mapeamento != null;

    public event EventHandler? DadosAlterados;

    public LojaDataService(IExcelWorkbookService excel)
    {
        _excel = excel;
    }

    public void Conectar(PlanilhaMapeamento mapeamento)
    {
        Mapeamento = mapeamento;
        if (!_excel.EstaConectado || !string.Equals(_excel.ArquivoAtual, mapeamento.CaminhoArquivo, StringComparison.OrdinalIgnoreCase))
        {
            _excel.AbrirPlanilha(mapeamento.CaminhoArquivo);
        }
        DadosAlterados?.Invoke(this, EventArgs.Empty);
    }

    public void Desconectar()
    {
        _excel.Desconectar();
        Mapeamento = null;
        DadosAlterados?.Invoke(this, EventArgs.Empty);
    }

    public List<ItemEstoque> ObterItensDisponiveis()
    {
        return LerTudo()
            .Where(i => i.Disponivel)
            .OrderBy(i => i.Categoria)
            .ThenBy(i => i.Produto)
            .ToList();
    }

    public List<ItemEstoque> ObterVendas()
    {
        return LerTudo()
            .Where(i => i.Vendido)
            .OrderByDescending(i => i.DataVenda)
            .ToList();
    }

    public void RegistrarEntrada(string sku, string produto, string categoria, decimal valorCusto, decimal precoAVista, decimal precoCartao)
    {
        if (Mapeamento == null)
        {
            throw new InvalidOperationException("Nenhuma planilha conectada.");
        }

        _excel.AdicionarItem(Mapeamento, new ItemEstoque
        {
            Sku = sku,
            Produto = produto,
            Categoria = categoria,
            Status = string.Empty,
            DataEntrada = DateTime.Now,
            ValorCusto = valorCusto,
            PrecoAVista = precoAVista,
            PrecoCartao = precoCartao
        });
        DadosAlterados?.Invoke(this, EventArgs.Empty);
    }

    public void RegistrarVenda(ItemEstoque item, decimal valorRecebido, string? cliente)
    {
        if (Mapeamento == null)
        {
            throw new InvalidOperationException("Nenhuma planilha conectada.");
        }

        var valores = new Dictionary<string, object?>
        {
            [CamposItem.Status] = "VENDIDO",
            [CamposItem.DataVenda] = DateTime.Now,
            [CamposItem.ValorRecebido] = valorRecebido,
            [CamposItem.Cliente] = cliente
        };
        _excel.AtualizarItem(Mapeamento, item.LinhaPlanilha, valores);
        DadosAlterados?.Invoke(this, EventArgs.Empty);
    }

    public void DarBaixa(ItemEstoque item)
    {
        if (Mapeamento == null)
        {
            throw new InvalidOperationException("Nenhuma planilha conectada.");
        }

        var valores = new Dictionary<string, object?>
        {
            [CamposItem.Status] = "BAIXA"
        };
        _excel.AtualizarItem(Mapeamento, item.LinhaPlanilha, valores);
        DadosAlterados?.Invoke(this, EventArgs.Empty);
    }

    public List<BalancoPeriodo> ObterBalanco(AgrupamentoBalanco agrupamento)
    {
        var itens = LerTudo();

        var entradasPorPeriodo = itens
            .Where(i => i.Vendido && i.DataVenda.HasValue)
            .GroupBy(i => ChaveAgrupamento(i.DataVenda!.Value, agrupamento))
            .ToDictionary(g => g.Key, g => g.Sum(i => i.ValorRecebido ?? 0));

        var saidasPorPeriodo = itens
            .GroupBy(i => ChaveAgrupamento(i.DataEntrada, agrupamento))
            .ToDictionary(g => g.Key, g => g.Sum(i => i.ValorCusto));

        var chaves = entradasPorPeriodo.Keys.Union(saidasPorPeriodo.Keys).Distinct();

        return chaves
            .Select(chave => new BalancoPeriodo
            {
                Periodo = chave.Rotulo,
                DataReferencia = chave.Referencia,
                Entradas = entradasPorPeriodo.TryGetValue(chave, out var entrada) ? entrada : 0,
                Saidas = saidasPorPeriodo.TryGetValue(chave, out var saida) ? saida : 0
            })
            .OrderBy(b => b.DataReferencia)
            .ToList();
    }

    private List<ItemEstoque> LerTudo()
    {
        return Mapeamento == null ? new List<ItemEstoque>() : _excel.LerItens(Mapeamento);
    }

    private static (string Rotulo, DateTime Referencia) ChaveAgrupamento(DateTime data, AgrupamentoBalanco agrupamento)
    {
        return agrupamento switch
        {
            AgrupamentoBalanco.Dia => (data.ToString("dd/MM/yyyy"), data.Date),
            AgrupamentoBalanco.Semana => SemanaDe(data),
            AgrupamentoBalanco.Mes => (Capitalizar(data.ToString("MMMM/yyyy", PtBr)), new DateTime(data.Year, data.Month, 1)),
            AgrupamentoBalanco.Ano => (data.Year.ToString(), new DateTime(data.Year, 1, 1)),
            _ => (data.ToString("dd/MM/yyyy"), data.Date)
        };
    }

    private static (string Rotulo, DateTime Referencia) SemanaDe(DateTime data)
    {
        var inicioSemana = data.Date.AddDays(-(int)data.DayOfWeek);
        return ($"Semana de {inicioSemana:dd/MM/yyyy}", inicioSemana);
    }

    private static string Capitalizar(string texto)
    {
        return string.IsNullOrEmpty(texto) ? texto : char.ToUpper(texto[0], PtBr) + texto[1..];
    }
}
