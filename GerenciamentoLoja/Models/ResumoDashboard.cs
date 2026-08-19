namespace GerenciamentoLoja.Models;

// Números já agregados para o dashboard. Existem como um pacote único porque
// todos saem da mesma leitura da planilha: cada ida ao Excel via COM é cara e
// pode falhar se o cliente estiver mexendo na janela, então lemos uma vez só.
public class ResumoDashboard
{
    public int TotalItens { get; init; }
    public int TotalDisponiveis { get; init; }
    public int TotalVendidos { get; init; }

    public List<ContagemRotulo> DisponiveisPorCategoria { get; init; } = new();
    public ContagemRotulo? ProdutoMaisVendido { get; init; }
    public List<MovimentoMensal> MovimentacaoMensal { get; init; } = new();
}

public class ContagemRotulo
{
    public string Rotulo { get; init; } = string.Empty;
    public int Quantidade { get; init; }
}

// Quantidade de peças (não valores em reais — isso é o Balanço geral) que
// entraram no estoque e que saíram por venda em cada mês.
public class MovimentoMensal
{
    public string Periodo { get; init; } = string.Empty;
    public DateTime DataReferencia { get; init; }
    public int Entradas { get; init; }
    public int Saidas { get; init; }
}
