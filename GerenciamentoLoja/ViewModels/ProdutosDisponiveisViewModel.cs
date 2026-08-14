using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class ProdutosDisponiveisViewModel : ObservableObject, INavigableViewModel
{
    private readonly ILojaDataService _dados;

    public ObservableCollection<ItemEstoque> Itens { get; } = new();

    [ObservableProperty]
    private string? filtroCategoria;

    [ObservableProperty]
    private string? mensagem;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(DarBaixaCommand))]
    private ItemEstoque? itemSelecionado;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RegistrarEntradaCommand))]
    private string skuEntrada = string.Empty;

    [ObservableProperty]
    private string produtoEntrada = string.Empty;

    [ObservableProperty]
    private string categoriaEntrada = string.Empty;

    [ObservableProperty]
    private decimal valorCustoEntrada;

    [ObservableProperty]
    private decimal precoAVistaEntrada;

    [ObservableProperty]
    private decimal precoCartaoEntrada;

    public ProdutosDisponiveisViewModel(ILojaDataService dados)
    {
        _dados = dados;
        _dados.DadosAlterados += (_, _) => CarregarItens();
    }

    public void AoNavegar() => CarregarItens();

    [RelayCommand]
    private void Filtrar() => CarregarItens();

    private void CarregarItens()
    {
        Itens.Clear();

        if (!_dados.EstaConectado)
        {
            Mensagem = "Conecte a planilha para ver o estoque.";
            return;
        }

        Mensagem = null;
        foreach (var item in _dados.ObterItensDisponiveis())
        {
            if (!string.IsNullOrWhiteSpace(FiltroCategoria) &&
                !item.Categoria.Contains(FiltroCategoria, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            Itens.Add(item);
        }
    }

    private bool PodeRegistrarEntrada() =>
        _dados.EstaConectado && !string.IsNullOrWhiteSpace(SkuEntrada);

    [RelayCommand(CanExecute = nameof(PodeRegistrarEntrada))]
    private void RegistrarEntrada()
    {
        try
        {
            _dados.RegistrarEntrada(SkuEntrada, ProdutoEntrada, CategoriaEntrada, ValorCustoEntrada, PrecoAVistaEntrada, PrecoCartaoEntrada);
            Mensagem = $"Entrada registrada para o SKU {SkuEntrada}.";
            SkuEntrada = string.Empty;
            ProdutoEntrada = string.Empty;
            CategoriaEntrada = string.Empty;
            ValorCustoEntrada = 0;
            PrecoAVistaEntrada = 0;
            PrecoCartaoEntrada = 0;
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao registrar entrada: {ex.Message}";
        }
    }

    private bool PodeDarBaixa() => _dados.EstaConectado && ItemSelecionado != null;

    [RelayCommand(CanExecute = nameof(PodeDarBaixa))]
    private void DarBaixa()
    {
        if (ItemSelecionado == null)
        {
            return;
        }

        try
        {
            _dados.DarBaixa(ItemSelecionado);
            Mensagem = $"Baixa registrada para o SKU {ItemSelecionado.Sku}.";
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao dar baixa: {ex.Message}";
        }
    }
}
