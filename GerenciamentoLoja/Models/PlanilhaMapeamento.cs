namespace GerenciamentoLoja.Models;

// Campos lógicos que o sistema precisa localizar na planilha do cliente.
// O layout segue a aba "CONTROLE DE VENDAS" real da loja: uma linha por peça,
// com data de entrada no estoque e (quando vendida) data/valor da venda.
public static class CamposItem
{
    public const string Sku = "SKU";
    public const string Produto = "Produto";
    public const string Categoria = "Categoria";
    public const string Status = "Status";
    public const string DataEntrada = "DataEntrada";
    public const string DataVenda = "DataVenda";
    public const string ValorRecebido = "ValorRecebido";
    public const string ValorCusto = "ValorCusto";
    public const string PrecoAVista = "PrecoAVista";
    public const string PrecoCartao = "PrecoCartao";
    public const string Cliente = "Cliente";

    public static readonly string[] Obrigatorios =
    {
        Sku, Produto, Categoria, Status, DataEntrada, DataVenda, ValorRecebido
    };

    public static readonly string[] Opcionais =
    {
        ValorCusto, PrecoAVista, PrecoCartao, Cliente
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
            && CamposItem.Obrigatorios.All(campo => Colunas.ContainsKey(campo));
    }
}
