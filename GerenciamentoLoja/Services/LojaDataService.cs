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

    public ResumoDashboard ObterResumoDashboard()
    {
        var itens = LerTudo();

        return new ResumoDashboard
        {
            TotalItens = itens.Count,
            TotalDisponiveis = itens.Count(i => i.Disponivel),
            TotalVendidos = itens.Count(i => i.Vendido),
            DisponiveisPorCategoria = itens
                .Where(i => i.Disponivel)
                .GroupBy(i => string.IsNullOrWhiteSpace(i.Categoria) ? "Sem categoria" : i.Categoria.Trim())
                .Select(g => new ContagemRotulo { Rotulo = g.Key, Quantidade = g.Count() })
                .OrderByDescending(c => c.Quantidade)
                .ThenBy(c => c.Rotulo)
                .ToList(),
            // "Produto que mais sai" conta peças vendidas com o mesmo nome de produto;
            // o SKU não serve porque é único por peça, então cada um venderia no máximo 1.
            ProdutoMaisVendido = itens
                .Where(i => i.Vendido)
                .GroupBy(i => string.IsNullOrWhiteSpace(i.Produto) ? "Sem nome" : i.Produto.Trim())
                .Select(g => new ContagemRotulo { Rotulo = g.Key, Quantidade = g.Count() })
                .OrderByDescending(c => c.Quantidade)
                .ThenBy(c => c.Rotulo)
                .FirstOrDefault(),
            MovimentacaoMensal = MontarMovimentacaoMensal(itens)
        };
    }

    private static List<MovimentoMensal> MontarMovimentacaoMensal(List<ItemEstoque> itens, int meses = 12)
    {
        // Data em branco (ou coluna não mapeada) chega como DateTime default e viraria
        // um mês fantasma no ano 1, esticando o gráfico inteiro — fora da contagem.
        var entradas = itens
            .Where(i => i.DataEntrada != default)
            .GroupBy(i => PrimeiroDiaDoMes(i.DataEntrada))
            .ToDictionary(g => g.Key, g => g.Count());

        var saidas = itens
            .Where(i => i.Vendido && i.DataVenda.HasValue && i.DataVenda.Value != default)
            .GroupBy(i => PrimeiroDiaDoMes(i.DataVenda!.Value))
            .ToDictionary(g => g.Key, g => g.Count());

        if (entradas.Count == 0 && saidas.Count == 0)
        {
            return new List<MovimentoMensal>();
        }

        // A janela termina no último mês com movimento, e não no mês atual: se o cliente
        // parou de lançar há um tempo, o gráfico ainda mostra os dados dele em vez de
        // uma sequência de meses vazios.
        var fim = entradas.Keys.Concat(saidas.Keys).Max();
        var inicio = fim.AddMonths(-(meses - 1));

        var resultado = new List<MovimentoMensal>();
        for (var mes = inicio; mes <= fim; mes = mes.AddMonths(1))
        {
            resultado.Add(new MovimentoMensal
            {
                DataReferencia = mes,
                // pt-BR abrevia os meses com ponto ("nov."), que no eixo do gráfico
                // vira "Nov./25" — o ponto sai para o rótulo ficar limpo.
                Periodo = Capitalizar(mes.ToString("MMM/yy", PtBr).Replace(".", string.Empty)),
                Entradas = entradas.TryGetValue(mes, out var entrada) ? entrada : 0,
                Saidas = saidas.TryGetValue(mes, out var saida) ? saida : 0
            });
        }
        return resultado;
    }

    private static DateTime PrimeiroDiaDoMes(DateTime data) => new(data.Year, data.Month, 1);

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
