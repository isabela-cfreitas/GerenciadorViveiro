using System;

namespace GerenciadorViveiro.Models;
public class Venda {
    public DateTime Data { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string Produto { get; set; } = string.Empty;
    public int Quantidade { get; set; }
    public decimal PrecoU { get; set; }
    public decimal PrecoT => Quantidade * PrecoU;
    public string FormaPagamento { get; set; } = string.Empty;
}