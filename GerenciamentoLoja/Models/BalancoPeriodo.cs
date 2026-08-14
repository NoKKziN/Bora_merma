namespace GerenciamentoLoja.Models;

public class BalancoPeriodo
{
    public string Periodo { get; set; } = string.Empty;
    public DateTime DataReferencia { get; set; }
    public decimal Entradas { get; set; }
    public decimal Saidas { get; set; }
    public decimal Saldo => Entradas - Saidas;
}

public enum AgrupamentoBalanco
{
    Dia,
    Semana,
    Mes,
    Ano
}
