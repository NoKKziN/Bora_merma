using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class ProdutosDisponiveisViewModel : ObservableObject, INavigableViewModel
{
    private readonly ILojaDataService _dados;

    public ObservableCollection<ProdutoEstoque> Produtos { get; } = new();

    [ObservableProperty]
    private string? filtroCategoria;

    [ObservableProperty]
    private string? mensagem;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(DarEntradaCommand))]
    [NotifyCanExecuteChangedFor(nameof(DarSaidaCommand))]
    private string skuMovimento = string.Empty;

    [ObservableProperty]
    private string produtoMovimento = string.Empty;

    [ObservableProperty]
    private string categoriaMovimento = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(DarEntradaCommand))]
    [NotifyCanExecuteChangedFor(nameof(DarSaidaCommand))]
    private int quantidadeMovimento = 1;

    [ObservableProperty]
    private decimal valorUnitarioMovimento;

    public ProdutosDisponiveisViewModel(ILojaDataService dados)
    {
        _dados = dados;
        _dados.DadosAlterados += (_, _) => CarregarProdutos();
    }

    public void AoNavegar() => CarregarProdutos();

    [RelayCommand]
    private void Filtrar() => CarregarProdutos();

    private void CarregarProdutos()
    {
        Produtos.Clear();

        if (!_dados.EstaConectado)
        {
            Mensagem = "Conecte a planilha para ver o estoque.";
            return;
        }

        Mensagem = null;
        foreach (var produto in _dados.ObterProdutosDisponiveis())
        {
            if (!string.IsNullOrWhiteSpace(FiltroCategoria) &&
                !produto.Categoria.Contains(FiltroCategoria, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            Produtos.Add(produto);
        }
    }

    private bool PodeRegistrarMovimento() =>
        _dados.EstaConectado && !string.IsNullOrWhiteSpace(SkuMovimento) && QuantidadeMovimento > 0;

    [RelayCommand(CanExecute = nameof(PodeRegistrarMovimento))]
    private void DarEntrada()
    {
        try
        {
            _dados.RegistrarEntradaEstoque(SkuMovimento, ProdutoMovimento, CategoriaMovimento, QuantidadeMovimento, ValorUnitarioMovimento);
            Mensagem = $"Entrada registrada para o SKU {SkuMovimento}.";
            LimparFormularioMovimento();
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao registrar entrada: {ex.Message}";
        }
    }

    [RelayCommand(CanExecute = nameof(PodeRegistrarMovimento))]
    private void DarSaida()
    {
        try
        {
            _dados.RegistrarSaidaEstoque(SkuMovimento, ProdutoMovimento, CategoriaMovimento, QuantidadeMovimento, ValorUnitarioMovimento, "Ajuste");
            Mensagem = $"Saída registrada para o SKU {SkuMovimento}.";
            LimparFormularioMovimento();
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao registrar saída: {ex.Message}";
        }
    }

    private void LimparFormularioMovimento()
    {
        SkuMovimento = string.Empty;
        ProdutoMovimento = string.Empty;
        CategoriaMovimento = string.Empty;
        QuantidadeMovimento = 1;
        ValorUnitarioMovimento = 0;
    }
}
