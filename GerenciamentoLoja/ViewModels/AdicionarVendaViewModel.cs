using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class AdicionarVendaViewModel : ObservableObject, INavigableViewModel
{
    private readonly ILojaDataService _dados;

    public ObservableCollection<ItemEstoque> ItensDisponiveis { get; } = new();

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RegistrarVendaCommand))]
    private ItemEstoque? itemSelecionado;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RegistrarVendaCommand))]
    private decimal valorRecebido;

    [ObservableProperty]
    private string? cliente;

    [ObservableProperty]
    private string? mensagem;

    public AdicionarVendaViewModel(ILojaDataService dados)
    {
        _dados = dados;
        _dados.DadosAlterados += (_, _) => CarregarItensDisponiveis();
    }

    public void AoNavegar() => CarregarItensDisponiveis();

    partial void OnItemSelecionadoChanged(ItemEstoque? value)
    {
        if (value != null && ValorRecebido == 0)
        {
            ValorRecebido = value.PrecoAVista;
        }
    }

    private void CarregarItensDisponiveis()
    {
        ItensDisponiveis.Clear();
        if (!_dados.EstaConectado)
        {
            Mensagem = "Conecte a planilha para registrar uma venda.";
            return;
        }

        Mensagem = null;
        foreach (var item in _dados.ObterItensDisponiveis())
        {
            ItensDisponiveis.Add(item);
        }
    }

    private bool PodeRegistrar() =>
        _dados.EstaConectado && ItemSelecionado != null && ValorRecebido >= 0;

    [RelayCommand(CanExecute = nameof(PodeRegistrar))]
    private void RegistrarVenda()
    {
        if (ItemSelecionado == null)
        {
            return;
        }

        try
        {
            _dados.RegistrarVenda(ItemSelecionado, ValorRecebido, Cliente);
            Mensagem = $"Venda registrada para o SKU {ItemSelecionado.Sku}.";
            ItemSelecionado = null;
            ValorRecebido = 0;
            Cliente = null;
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao registrar venda: {ex.Message}";
        }
    }
}
