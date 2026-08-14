namespace GerenciamentoLoja.Models;

// Associa o texto de um cabeçalho ao número real da coluna na planilha (1-based,
// A=1). Necessário porque a área usada de uma aba nem sempre começa na coluna A
// (ex: "CONTROLE DE VENDAS" começa na coluna F) — sem isso, reconstruir o número
// da coluna a partir da posição numa lista leva a índices errados.
public record ColunaCabecalho(int Coluna, string Texto);
