using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

// Camada de negócio sobre a planilha: traduz movimentos brutos em estoque,
// vendas e balanço financeiro, e concentra as regras de escrita.
public interface ILojaDataService
{
    bool EstaConectado { get; }
    PlanilhaMapeamento? Mapeamento { get; }

    event EventHandler? DadosAlterados;

    void Conectar(PlanilhaMapeamento mapeamento);
    void Desconectar();

    List<ProdutoEstoque> ObterProdutosDisponiveis();
    List<Movimento> ObterVendas();

    void RegistrarEntradaEstoque(string sku, string produto, string categoria, int quantidade, decimal valorUnitario);
    void RegistrarSaidaEstoque(string sku, string produto, string categoria, int quantidade, decimal valorUnitario, string motivo);
    void RegistrarVenda(string sku, string produto, string categoria, int quantidade, decimal valorUnitario, string? cliente);

    List<BalancoPeriodo> ObterBalanco(AgrupamentoBalanco agrupamento);
}
