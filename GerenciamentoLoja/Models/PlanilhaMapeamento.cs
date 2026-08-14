namespace GerenciamentoLoja.Models;

// Campos lógicos que o sistema precisa localizar na planilha do cliente.
public static class CamposMovimento
{
    public const string Data = "Data";
    public const string Sku = "SKU";
    public const string Produto = "Produto";
    public const string Categoria = "Categoria";
    public const string Tipo = "Tipo";
    public const string Quantidade = "Quantidade";
    public const string ValorUnitario = "ValorUnitario";
    public const string Cliente = "Cliente";
    public const string Motivo = "Motivo";

    public static readonly string[] Obrigatorios =
    {
        Data, Sku, Produto, Tipo, Quantidade, ValorUnitario
    };

    public static readonly string[] Opcionais =
    {
        Categoria, Cliente, Motivo
    };

    public static readonly string[] Todos = Obrigatorios.Concat(Opcionais).ToArray();
}

public class PlanilhaMapeamento
{
    public string CaminhoArquivo { get; set; } = string.Empty;
    public string NomeAba { get; set; } = string.Empty;
    public int LinhaCabecalho { get; set; } = 1;

    // Campo lógico -> índice de coluna na planilha (1-based)
    public Dictionary<string, int> Colunas { get; set; } = new();

    public bool EstaCompleto()
    {
        return !string.IsNullOrWhiteSpace(CaminhoArquivo)
            && !string.IsNullOrWhiteSpace(NomeAba)
            && CamposMovimento.Obrigatorios.All(campo => Colunas.ContainsKey(campo));
    }
}
