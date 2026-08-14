using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class VendasRealizadasViewModel : ObservableObject, INavigableViewModel
{
    private readonly ILojaDataService _dados;

    public ObservableCollection<Movimento> Vendas { get; } = new();

    [ObservableProperty]
    private string? mensagem;

    [ObservableProperty]
    private decimal totalVendas;

    public VendasRealizadasViewModel(ILojaDataService dados)
    {
        _dados = dados;
        _dados.DadosAlterados += (_, _) => Carregar();
    }

    public void AoNavegar() => Carregar();

    private void Carregar()
    {
        Vendas.Clear();

        if (!_dados.EstaConectado)
        {
            Mensagem = "Conecte a planilha para ver as vendas.";
            TotalVendas = 0;
            return;
        }

        Mensagem = null;
        foreach (var venda in _dados.ObterVendas())
        {
            Vendas.Add(venda);
        }
        TotalVendas = Vendas.Sum(v => v.ValorTotal);
    }
}
