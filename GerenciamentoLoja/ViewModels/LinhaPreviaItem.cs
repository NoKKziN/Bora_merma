namespace GerenciamentoLoja.ViewModels;

// Uma linha "crua" da planilha mostrada na tela de conexão, para o usuário
// escolher visualmente qual é a linha de cabeçalho em vez de digitar um número.
public class LinhaPreviaItem
{
    public int Numero { get; init; }
    public string Resumo { get; init; } = string.Empty;
}
