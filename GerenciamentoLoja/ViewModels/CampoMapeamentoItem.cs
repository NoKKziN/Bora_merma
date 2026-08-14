using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciamentoLoja.ViewModels;

public partial class CampoMapeamentoItem : ObservableObject
{
    public string CampoLogico { get; }
    public string Rotulo { get; }
    public bool Obrigatorio { get; }
    public ObservableCollection<string> Opcoes { get; } = new();

    [ObservableProperty]
    private string? colunaSelecionada;

    public CampoMapeamentoItem(string campoLogico, string rotulo, bool obrigatorio)
    {
        CampoLogico = campoLogico;
        Rotulo = rotulo;
        Obrigatorio = obrigatorio;
    }
}
