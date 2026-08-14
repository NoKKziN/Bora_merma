namespace GerenciamentoLoja.Models;

// Cada linha da planilha do cliente representa uma peça individual (não uma
// quantidade de um produto): SKU único por unidade, com data de entrada no
// estoque e, quando vendida, data/valor da venda e status "Vendido".
public class ItemEstoque
{
    public int LinhaPlanilha { get; set; }
    public string Sku { get; set; } = string.Empty;
    public string Produto { get; set; } = string.Empty;
    public string Categoria { get; set; } = string.Empty;
    public decimal ValorCusto { get; set; }
    public decimal PrecoAVista { get; set; }
    public decimal PrecoCartao { get; set; }
    public string Status { get; set; } = string.Empty;
    public DateTime DataEntrada { get; set; }
    public DateTime? DataVenda { get; set; }
    public decimal? ValorRecebido { get; set; }
    public string? Cliente { get; set; }

    public bool Vendido => Status.Trim().Equals("VENDIDO", StringComparison.OrdinalIgnoreCase);
    public bool Disponivel => !Vendido;
}
