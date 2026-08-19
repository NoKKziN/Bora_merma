using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

// Camada de negócio sobre a planilha: traduz as peças cadastradas em estoque
// disponível, vendas e balanço financeiro, e concentra as regras de escrita.
public interface ILojaDataService
{
    bool EstaConectado { get; }
    PlanilhaMapeamento? Mapeamento { get; }

    event EventHandler? DadosAlterados;

    void Conectar(PlanilhaMapeamento mapeamento);
    void Desconectar();

    List<ItemEstoque> ObterItensDisponiveis();
    List<ItemEstoque> ObterVendas();

    void RegistrarEntrada(string sku, string produto, string categoria, decimal valorCusto, decimal precoAVista, decimal precoCartao);
    void RegistrarVenda(ItemEstoque item, decimal valorRecebido, string? cliente);
    void DarBaixa(ItemEstoque item);

    List<BalancoPeriodo> ObterBalanco(AgrupamentoBalanco agrupamento);

    ResumoDashboard ObterResumoDashboard();
}
