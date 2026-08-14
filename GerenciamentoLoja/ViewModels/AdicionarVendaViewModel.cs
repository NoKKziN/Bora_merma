using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class AdicionarVendaViewModel : ObservableObject
{
    private readonly ILojaDataService _dados;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RegistrarVendaCommand))]
    private string sku = string.Empty;

    [ObservableProperty]
    private string produto = string.Empty;

    [ObservableProperty]
    private string categoria = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RegistrarVendaCommand))]
    [NotifyPropertyChangedFor(nameof(Total))]
    private int quantidade = 1;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RegistrarVendaCommand))]
    [NotifyPropertyChangedFor(nameof(Total))]
    private decimal valorUnitario;

    [ObservableProperty]
    private string? cliente;

    [ObservableProperty]
    private string? mensagem;

    public decimal Total => Quantidade * ValorUnitario;

    public AdicionarVendaViewModel(ILojaDataService dados)
    {
        _dados = dados;
    }

    private bool PodeRegistrar() =>
        _dados.EstaConectado && !string.IsNullOrWhiteSpace(Sku) && Quantidade > 0 && ValorUnitario >= 0;

    [RelayCommand(CanExecute = nameof(PodeRegistrar))]
    private void RegistrarVenda()
    {
        try
        {
            _dados.RegistrarVenda(Sku, Produto, Categoria, Quantidade, ValorUnitario, Cliente);
            Mensagem = "Venda registrada com sucesso!";
            Sku = string.Empty;
            Produto = string.Empty;
            Categoria = string.Empty;
            Quantidade = 1;
            ValorUnitario = 0;
            Cliente = null;
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao registrar venda: {ex.Message}";
        }
    }
}
