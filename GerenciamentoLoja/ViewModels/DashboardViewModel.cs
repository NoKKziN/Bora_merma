using CommunityToolkit.Mvvm.ComponentModel;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;

namespace GerenciamentoLoja.ViewModels;

public partial class DashboardViewModel : ObservableObject, INavigableViewModel
{
    // Paleta alinhada ao tema Material do app (primária DeepPurple). A saída usa
    // laranja em vez do secundário Lime, que fica ilegível sobre fundo claro.
    private static readonly SKColor CorEntrada = new(0x5E, 0x35, 0xB1);
    private static readonly SKColor CorSaida = new(0xFB, 0x8C, 0x00);
    private static readonly SKColor CorTexto = new(0x61, 0x61, 0x61);
    private static readonly SKColor CorGrade = new(0xE0, 0xE0, 0xE0);

    // Acima disso os rótulos do eixo X viram um borrão. O excedente é somado em
    // "Outras" para o gráfico continuar representando o estoque inteiro.
    private const int MaximoCategoriasNoGrafico = 8;

    private readonly ILojaDataService _dados;

    [ObservableProperty]
    private string? mensagem;

    [ObservableProperty]
    private bool temDados;

    [ObservableProperty]
    private int totalDisponiveis;

    [ObservableProperty]
    private int totalVendidos;

    [ObservableProperty]
    private int totalItens;

    [ObservableProperty]
    private string produtoMaisVendido = "—";

    [ObservableProperty]
    private string produtoMaisVendidoDetalhe = "Nenhuma venda registrada";

    [ObservableProperty]
    private ISeries[] seriesCategorias = [];

    [ObservableProperty]
    private Axis[] eixoXCategorias = [];

    [ObservableProperty]
    private Axis[] eixoYCategorias = [];

    [ObservableProperty]
    private ISeries[] seriesMovimentacao = [];

    [ObservableProperty]
    private Axis[] eixoXMovimentacao = [];

    [ObservableProperty]
    private Axis[] eixoYMovimentacao = [];

    public DashboardViewModel(ILojaDataService dados)
    {
        _dados = dados;
        _dados.DadosAlterados += (_, _) => Carregar();
        Carregar();
    }

    public void AoNavegar() => Carregar();

    private void Carregar()
    {
        if (!_dados.EstaConectado)
        {
            Limpar("Conecte a planilha para ver o dashboard.");
            return;
        }

        ResumoDashboard resumo;
        try
        {
            resumo = _dados.ObterResumoDashboard();
        }
        catch (Exception ex)
        {
            Limpar($"Não foi possível ler a planilha: {ex.Message}");
            return;
        }

        Mensagem = null;
        TemDados = true;

        TotalItens = resumo.TotalItens;
        TotalDisponiveis = resumo.TotalDisponiveis;
        TotalVendidos = resumo.TotalVendidos;

        if (resumo.ProdutoMaisVendido is { } campeao)
        {
            ProdutoMaisVendido = campeao.Rotulo;
            ProdutoMaisVendidoDetalhe = campeao.Quantidade == 1
                ? "1 peça vendida"
                : $"{campeao.Quantidade} peças vendidas";
        }
        else
        {
            ProdutoMaisVendido = "—";
            ProdutoMaisVendidoDetalhe = "Nenhuma venda registrada";
        }

        MontarGraficoCategorias(resumo.DisponiveisPorCategoria);
        MontarGraficoMovimentacao(resumo.MovimentacaoMensal);
    }

    private void MontarGraficoCategorias(List<ContagemRotulo> categorias)
    {
        var exibidas = categorias.Take(MaximoCategoriasNoGrafico).ToList();
        var restante = categorias.Skip(MaximoCategoriasNoGrafico).Sum(c => c.Quantidade);
        if (restante > 0)
        {
            exibidas.Add(new ContagemRotulo { Rotulo = "Outras", Quantidade = restante });
        }

        SeriesCategorias =
        [
            new ColumnSeries<int>
            {
                Name = "Peças em estoque",
                Values = exibidas.Select(c => c.Quantidade).ToArray(),
                Fill = new SolidColorPaint(CorEntrada),
                MaxBarWidth = 56,
                DataLabelsPaint = new SolidColorPaint(CorTexto),
                DataLabelsPosition = LiveChartsCore.Measure.DataLabelsPosition.Top,
                DataLabelsSize = 12
            }
        ];
        EixoXCategorias = [EixoDeRotulos(exibidas.Select(c => c.Rotulo).ToArray())];
        EixoYCategorias = [EixoDeQuantidade()];
    }

    private void MontarGraficoMovimentacao(List<MovimentoMensal> meses)
    {
        SeriesMovimentacao =
        [
            new ColumnSeries<int>
            {
                Name = "Entraram",
                Values = meses.Select(m => m.Entradas).ToArray(),
                Fill = new SolidColorPaint(CorEntrada),
                MaxBarWidth = 24
            },
            new ColumnSeries<int>
            {
                Name = "Saíram",
                Values = meses.Select(m => m.Saidas).ToArray(),
                Fill = new SolidColorPaint(CorSaida),
                MaxBarWidth = 24
            }
        ];
        EixoXMovimentacao = [EixoDeRotulos(meses.Select(m => m.Periodo).ToArray())];
        EixoYMovimentacao = [EixoDeQuantidade()];
    }

    // ForceStepToMin com MinStep 1 impede o LiveCharts de pular rótulos quando o
    // eixo fica apertado — em um eixo de categorias, sumir com um rótulo desalinha
    // a leitura de todas as barras.
    private static Axis EixoDeRotulos(string[] rotulos) => new()
    {
        Labels = rotulos,
        MinStep = 1,
        ForceStepToMin = true,
        TextSize = 12,
        LabelsPaint = new SolidColorPaint(CorTexto),
        SeparatorsPaint = null
    };

    // São contagens de peças: o eixo precisa começar no zero e andar de 1 em 1,
    // senão aparecem marcas como "1,5 peças".
    private static Axis EixoDeQuantidade() => new()
    {
        MinLimit = 0,
        MinStep = 1,
        TextSize = 12,
        Labeler = valor => valor.ToString("N0"),
        LabelsPaint = new SolidColorPaint(CorTexto),
        SeparatorsPaint = new SolidColorPaint(CorGrade) { StrokeThickness = 1 }
    };

    private void Limpar(string mensagem)
    {
        Mensagem = mensagem;
        TemDados = false;
        TotalItens = 0;
        TotalDisponiveis = 0;
        TotalVendidos = 0;
        ProdutoMaisVendido = "—";
        ProdutoMaisVendidoDetalhe = "Nenhuma venda registrada";
        SeriesCategorias = [];
        SeriesMovimentacao = [];
        EixoXCategorias = [];
        EixoYCategorias = [];
        EixoXMovimentacao = [];
        EixoYMovimentacao = [];
    }
}
