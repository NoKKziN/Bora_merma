namespace GerenciamentoLoja.Models;

public class Movimento
{
    public int LinhaPlanilha { get; set; }
    public DateTime Data { get; set; }
    public string Sku { get; set; } = string.Empty;
    public string Produto { get; set; } = string.Empty;
    public string Categoria { get; set; } = string.Empty;
    public TipoMovimento Tipo { get; set; }
    public int Quantidade { get; set; }
    public decimal ValorUnitario { get; set; }
    public string? Cliente { get; set; }
    public string? Motivo { get; set; }

    public decimal ValorTotal => Quantidade * ValorUnitario;

    public bool EhVenda => Tipo == TipoMovimento.Saida &&
        (string.IsNullOrWhiteSpace(Motivo) || Motivo.Equals("Venda", StringComparison.OrdinalIgnoreCase));
}
