namespace GerenciamentoLoja.ViewModels;

// Implementado pelos ViewModels de tela que precisam recarregar dados da
// planilha toda vez que o usuário navega até eles pelo menu.
public interface INavigableViewModel
{
    void AoNavegar();
}
