using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class BalancoGeralViewModel : ObservableObject, INavigableViewModel
{
    private readonly ILojaDataService _dados;

    public ObservableCollection<BalancoPeriodo> Periodos { get; } = new();
    public List<AgrupamentoBalanco> Agrupamentos { get; } = Enum.GetValues<AgrupamentoBalanco>().ToList();

    [ObservableProperty]
    private AgrupamentoBalanco agrupamentoSelecionado = AgrupamentoBalanco.Mes;

    [ObservableProperty]
    private string? mensagem;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SaldoTotal))]
    private decimal totalEntradas;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SaldoTotal))]
    private decimal totalSaidas;

    public decimal SaldoTotal => TotalEntradas - TotalSaidas;

    public BalancoGeralViewModel(ILojaDataService dados)
    {
        _dados = dados;
        _dados.DadosAlterados += (_, _) => Carregar();
    }

    public void AoNavegar() => Carregar();

    partial void OnAgrupamentoSelecionadoChanged(AgrupamentoBalanco value) => Carregar();

    private void Carregar()
    {
        Periodos.Clear();

        if (!_dados.EstaConectado)
        {
            Mensagem = "Conecte a planilha para ver o balanço financeiro.";
            TotalEntradas = 0;
            TotalSaidas = 0;
            return;
        }

        Mensagem = null;
        foreach (var periodo in _dados.ObterBalanco(AgrupamentoSelecionado))
        {
            Periodos.Add(periodo);
        }
        TotalEntradas = Periodos.Sum(p => p.Entradas);
        TotalSaidas = Periodos.Sum(p => p.Saidas);
    }
}
