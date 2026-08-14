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
    List<string> ObterCabecalhos(string nomeAba, int linhaCabecalho);

    List<Movimento> LerMovimentos(PlanilhaMapeamento mapeamento);
    void AdicionarMovimento(PlanilhaMapeamento mapeamento, Movimento movimento);
}
