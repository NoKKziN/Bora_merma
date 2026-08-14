using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

public interface IExcelWorkbookService : IDisposable
{
    bool EstaConectado { get; }
    string? ArquivoAtual { get; }

    event EventHandler? ConexaoAlterada;

    void AbrirPlanilha(string caminhoArquivo);
    void Desconectar();

    List<string> ObterNomesAbas();
    List<ColunaCabecalho> ObterCabecalhos(string nomeAba, int linhaCabecalho);
    List<List<string>> ObterAmostraLinhas(string nomeAba, int quantidadeLinhas);

    List<ItemEstoque> LerItens(PlanilhaMapeamento mapeamento);
    void AdicionarItem(PlanilhaMapeamento mapeamento, ItemEstoque item);
    void AtualizarItem(PlanilhaMapeamento mapeamento, int linhaPlanilha, IReadOnlyDictionary<string, object?> valores);
}
