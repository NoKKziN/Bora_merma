namespace GerenciamentoLoja.Models;

public class ProdutoEstoque
{
    public string Sku { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
    public string Categoria { get; set; } = string.Empty;
    public int QuantidadeDisponivel { get; set; }
    public decimal ValorUnitarioMedio { get; set; }
}
