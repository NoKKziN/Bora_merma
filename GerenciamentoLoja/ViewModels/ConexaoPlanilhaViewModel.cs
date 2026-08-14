using System.Collections.ObjectModel;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GerenciamentoLoja.Models;
using GerenciamentoLoja.Services;
using Microsoft.Win32;

namespace GerenciamentoLoja.ViewModels;

public partial class ConexaoPlanilhaViewModel : ObservableObject
{
    private const int LinhasDeAmostra = 12;

    private readonly ILojaDataService _dados;
    private readonly IExcelWorkbookService _excel;
    private readonly IConfigService _config;

    public event EventHandler? Conectado;

    [ObservableProperty]
    private string? caminhoArquivo;

    [ObservableProperty]
    private string? abaSelecionada;

    [ObservableProperty]
    private LinhaPreviaItem? linhaCabecalhoSelecionada;

    [ObservableProperty]
    private string? mensagem;

    [ObservableProperty]
    private bool carregando;

    public ObservableCollection<string> AbasDisponiveis { get; } = new();
    public ObservableCollection<LinhaPreviaItem> LinhasPrevia { get; } = new();
    public ObservableCollection<CampoMapeamentoItem> Campos { get; } = new();

    private int LinhaCabecalho => LinhaCabecalhoSelecionada?.Numero ?? 0;

    public ConexaoPlanilhaViewModel(ILojaDataService dados, IExcelWorkbookService excel, IConfigService config)
    {
        _dados = dados;
        _excel = excel;
        _config = config;

        Campos.Add(new CampoMapeamentoItem(CamposItem.Sku, "SKU / código da peça", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.Produto, "Nome / descrição do produto", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.Categoria, "Categoria (tipo do produto)", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.Status, "Status (ex: VENDIDO)", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.DataEntrada, "Data de entrada no estoque", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.DataVenda, "Data da venda", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.ValorRecebido, "Valor recebido na venda", true));
        Campos.Add(new CampoMapeamentoItem(CamposItem.ValorCusto, "Valor de custo (opcional)", false));
        Campos.Add(new CampoMapeamentoItem(CamposItem.PrecoAVista, "Preço à vista (opcional)", false));
        Campos.Add(new CampoMapeamentoItem(CamposItem.PrecoCartao, "Preço no cartão (opcional)", false));
        Campos.Add(new CampoMapeamentoItem(CamposItem.Cliente, "Cliente (opcional)", false));

        CarregarMapeamentoSalvo();
    }

    private void CarregarMapeamentoSalvo()
    {
        var mapeamento = _config.Carregar().Mapeamento;
        if (mapeamento == null || !File.Exists(mapeamento.CaminhoArquivo))
        {
            return;
        }

        CaminhoArquivo = mapeamento.CaminhoArquivo;
        Mensagem = "Encontramos uma conexão salva. Clique em \"Conectar\" para reabrir a planilha.";
    }

    [RelayCommand]
    private void SelecionarArquivo()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Planilhas do Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            Title = "Selecione a planilha da loja"
        };

        if (dialog.ShowDialog() == true)
        {
            CaminhoArquivo = dialog.FileName;
            CarregarAbas();
        }
    }

    private void CarregarAbas()
    {
        if (string.IsNullOrWhiteSpace(CaminhoArquivo))
        {
            return;
        }

        try
        {
            Carregando = true;
            Mensagem = null;
            _excel.AbrirPlanilha(CaminhoArquivo);

            AbasDisponiveis.Clear();
            foreach (var nome in _excel.ObterNomesAbas())
            {
                AbasDisponiveis.Add(nome);
            }

            AbaSelecionada = AbasDisponiveis.FirstOrDefault();
        }
        catch (Exception ex)
        {
            Mensagem = $"Não foi possível abrir a planilha: {ex.Message}";
        }
        finally
        {
            Carregando = false;
        }
    }

    partial void OnAbaSelecionadaChanged(string? value)
    {
        CarregarAmostra();
    }

    partial void OnLinhaCabecalhoSelecionadaChanged(LinhaPreviaItem? value)
    {
        CarregarColunas();
    }

    private void CarregarAmostra()
    {
        LinhasPrevia.Clear();
        LinhaCabecalhoSelecionada = null;

        if (string.IsNullOrWhiteSpace(AbaSelecionada) || !_excel.EstaConectado)
        {
            return;
        }

        try
        {
            var linhas = _excel.ObterAmostraLinhas(AbaSelecionada, LinhasDeAmostra);

            int melhorIndice = 0;
            int melhorContagem = -1;
            for (int i = 0; i < linhas.Count; i++)
            {
                var preenchidas = linhas[i].Count(c => !string.IsNullOrWhiteSpace(c));
                if (preenchidas > melhorContagem)
                {
                    melhorContagem = preenchidas;
                    melhorIndice = i;
                }

                var resumo = string.Join("  |  ", linhas[i].Where(c => !string.IsNullOrWhiteSpace(c)).Take(6));
                LinhasPrevia.Add(new LinhaPreviaItem
                {
                    Numero = i + 1,
                    Resumo = string.IsNullOrWhiteSpace(resumo) ? "(linha vazia)" : resumo
                });
            }

            LinhaCabecalhoSelecionada = LinhasPrevia.Count > 0 ? LinhasPrevia[melhorIndice] : null;
        }
        catch (Exception ex)
        {
            Mensagem = $"Não foi possível ler as linhas da aba selecionada: {ex.Message}";
        }
    }

    private void CarregarColunas()
    {
        if (string.IsNullOrWhiteSpace(AbaSelecionada) || !_excel.EstaConectado || LinhaCabecalho < 1)
        {
            return;
        }

        try
        {
            var cabecalhos = _excel.ObterCabecalhos(AbaSelecionada, LinhaCabecalho)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .ToList();

            foreach (var campo in Campos)
            {
                var selecaoAtual = campo.ColunaSelecionada;
                campo.Opcoes.Clear();
                foreach (var cabecalho in cabecalhos)
                {
                    campo.Opcoes.Add(cabecalho);
                }
                campo.ColunaSelecionada = selecaoAtual != null && cabecalhos.Contains(selecaoAtual) ? selecaoAtual : null;
            }
        }
        catch (Exception ex)
        {
            Mensagem = $"Não foi possível ler as colunas da aba selecionada: {ex.Message}";
        }
    }

    private bool PodeConectar()
    {
        return !string.IsNullOrWhiteSpace(CaminhoArquivo)
            && !string.IsNullOrWhiteSpace(AbaSelecionada)
            && LinhaCabecalho > 0
            && Campos.Where(c => c.Obrigatorio).All(c => !string.IsNullOrWhiteSpace(c.ColunaSelecionada));
    }

    [RelayCommand(CanExecute = nameof(PodeConectar))]
    private void Conectar()
    {
        try
        {
            Carregando = true;
            Mensagem = null;

            if (!_excel.EstaConectado || !string.Equals(_excel.ArquivoAtual, CaminhoArquivo, StringComparison.OrdinalIgnoreCase))
            {
                _excel.AbrirPlanilha(CaminhoArquivo!);
                if (AbasDisponiveis.Count == 0)
                {
                    foreach (var nome in _excel.ObterNomesAbas())
                    {
                        AbasDisponiveis.Add(nome);
                    }
                    AbaSelecionada ??= AbasDisponiveis.FirstOrDefault();
                }
            }

            var mapeamento = new PlanilhaMapeamento
            {
                CaminhoArquivo = CaminhoArquivo!,
                NomeAba = AbaSelecionada!,
                LinhaCabecalho = LinhaCabecalho
            };

            var cabecalhos = _excel.ObterCabecalhos(AbaSelecionada!, LinhaCabecalho);
            foreach (var campo in Campos)
            {
                if (string.IsNullOrWhiteSpace(campo.ColunaSelecionada))
                {
                    continue;
                }
                var indice = cabecalhos.FindIndex(c => c == campo.ColunaSelecionada);
                if (indice >= 0)
                {
                    mapeamento.Colunas[campo.CampoLogico] = indice + 1;
                }
            }

            if (!mapeamento.EstaCompleto())
            {
                Mensagem = "Preencha todos os campos obrigatórios antes de conectar.";
                return;
            }

            _dados.Conectar(mapeamento);
            _config.Salvar(new AppConfig { Mapeamento = mapeamento });

            Mensagem = "Conectado com sucesso!";
            Conectado?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            Mensagem = $"Erro ao conectar: {ex.Message}";
        }
        finally
        {
            Carregando = false;
        }
    }
}
