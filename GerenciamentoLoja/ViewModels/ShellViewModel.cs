using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GerenciamentoLoja.Services;

namespace GerenciamentoLoja.ViewModels;

public partial class ShellViewModel : ObservableObject
{
    private readonly ILojaDataService _dados;

    public ConexaoPlanilhaViewModel Conexao { get; }
    public DashboardPlaceholderViewModel Dashboard { get; }
    public ProdutosDisponiveisViewModel Produtos { get; }
    public VendasRealizadasViewModel VendasRealizadas { get; }
    public AdicionarVendaViewModel AdicionarVenda { get; }
    public BalancoGeralViewModel Balanco { get; }

    [ObservableProperty]
    private object currentViewModel;

    [ObservableProperty]
    private string tituloPagina = "Dashboard";

    [ObservableProperty]
    private bool estaConectado;

    public ShellViewModel(
        ILojaDataService dados,
        ConexaoPlanilhaViewModel conexao,
        DashboardPlaceholderViewModel dashboard,
        ProdutosDisponiveisViewModel produtos,
        VendasRealizadasViewModel vendasRealizadas,
        AdicionarVendaViewModel adicionarVenda,
        BalancoGeralViewModel balanco)
    {
        _dados = dados;
        Conexao = conexao;
        Dashboard = dashboard;
        Produtos = produtos;
        VendasRealizadas = vendasRealizadas;
        AdicionarVenda = adicionarVenda;
        Balanco = balanco;

        Conexao.Conectado += (_, _) => NavegarPara(Dashboard, "Dashboard");
        _dados.DadosAlterados += (_, _) => EstaConectado = _dados.EstaConectado;

        EstaConectado = _dados.EstaConectado;
        currentViewModel = EstaConectado ? Dashboard : Conexao;
    }

    [RelayCommand]
    private void NavegarParaDashboard() => NavegarPara(Dashboard, "Dashboard");

    [RelayCommand]
    private void NavegarParaProdutos() => NavegarPara(Produtos, "Produtos disponíveis");

    [RelayCommand]
    private void NavegarParaVendasRealizadas() => NavegarPara(VendasRealizadas, "Vendas realizadas");

    [RelayCommand]
    private void NavegarParaAdicionarVenda() => NavegarPara(AdicionarVenda, "Adicionar venda");

    [RelayCommand]
    private void NavegarParaBalanco() => NavegarPara(Balanco, "Balanço geral");

    [RelayCommand]
    private void NavegarParaConexao() => NavegarPara(Conexao, "Conexão com a planilha");

    private void NavegarPara(object viewModel, string titulo)
    {
        if (viewModel is INavigableViewModel navegavel)
        {
            navegavel.AoNavegar();
        }
        CurrentViewModel = viewModel;
        TituloPagina = titulo;
    }
}
