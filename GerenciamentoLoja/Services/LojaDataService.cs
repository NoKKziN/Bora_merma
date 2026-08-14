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

    public List<ProdutoEstoque> ObterProdutosDisponiveis()
    {
        var movimentos = LerTudo();

        return movimentos
            .GroupBy(m => m.Sku)
            .Select(grupo =>
            {
                var maisRecente = grupo.OrderByDescending(m => m.Data).First();
                var categoria = grupo
                    .OrderByDescending(m => m.Data)
                    .Select(m => m.Categoria)
                    .FirstOrDefault(c => !string.IsNullOrWhiteSpace(c)) ?? string.Empty;

                return new ProdutoEstoque
                {
                    Sku = grupo.Key,
                    Nome = maisRecente.Produto,
                    Categoria = categoria,
                    QuantidadeDisponivel = grupo.Sum(m => m.Tipo == TipoMovimento.Entrada ? m.Quantidade : -m.Quantidade),
                    ValorUnitarioMedio = maisRecente.ValorUnitario
                };
            })
            .OrderBy(p => p.Categoria)
            .ThenBy(p => p.Nome)
            .ToList();
    }

    public List<Movimento> ObterVendas()
    {
        return LerTudo()
            .Where(m => m.EhVenda)
            .OrderByDescending(m => m.Data)
            .ToList();
    }

    public void RegistrarEntradaEstoque(string sku, string produto, string categoria, int quantidade, decimal valorUnitario)
    {
        Registrar(new Movimento
        {
            Data = DateTime.Now,
            Sku = sku,
            Produto = produto,
            Categoria = categoria,
            Tipo = TipoMovimento.Entrada,
            Quantidade = quantidade,
            ValorUnitario = valorUnitario,
            Motivo = "Reposicao"
        });
    }

    public void RegistrarSaidaEstoque(string sku, string produto, string categoria, int quantidade, decimal valorUnitario, string motivo)
    {
        Registrar(new Movimento
        {
            Data = DateTime.Now,
            Sku = sku,
            Produto = produto,
            Categoria = categoria,
            Tipo = TipoMovimento.Saida,
            Quantidade = quantidade,
            ValorUnitario = valorUnitario,
            Motivo = motivo
        });
    }

    public void RegistrarVenda(string sku, string produto, string categoria, int quantidade, decimal valorUnitario, string? cliente)
    {
        Registrar(new Movimento
        {
            Data = DateTime.Now,
            Sku = sku,
            Produto = produto,
            Categoria = categoria,
            Tipo = TipoMovimento.Saida,
            Quantidade = quantidade,
            ValorUnitario = valorUnitario,
            Motivo = "Venda",
            Cliente = cliente
        });
    }

    public List<BalancoPeriodo> ObterBalanco(AgrupamentoBalanco agrupamento)
    {
        var movimentos = LerTudo();

        return movimentos
            .GroupBy(m => ChaveAgrupamento(m.Data, agrupamento))
            .Select(grupo => new BalancoPeriodo
            {
                Periodo = grupo.Key.Rotulo,
                DataReferencia = grupo.Key.Referencia,
                // Saída de mercadoria = venda = dinheiro entrando no caixa.
                Entradas = grupo.Where(m => m.Tipo == TipoMovimento.Saida).Sum(m => m.ValorTotal),
                // Entrada de mercadoria = compra/reposição = dinheiro saindo do caixa.
                Saidas = grupo.Where(m => m.Tipo == TipoMovimento.Entrada).Sum(m => m.ValorTotal)
            })
            .OrderBy(b => b.DataReferencia)
            .ToList();
    }

    private List<Movimento> LerTudo()
    {
        return Mapeamento == null ? new List<Movimento>() : _excel.LerMovimentos(Mapeamento);
    }

    private void Registrar(Movimento movimento)
    {
        if (Mapeamento == null)
        {
            throw new InvalidOperationException("Nenhuma planilha conectada.");
        }

        _excel.AdicionarMovimento(Mapeamento, movimento);
        DadosAlterados?.Invoke(this, EventArgs.Empty);
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
