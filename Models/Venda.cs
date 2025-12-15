using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.Models;

public partial class Venda : ObservableObject {
    [ObservableProperty]
    private DateTime data = DateTime.Today;

    [ObservableProperty]
    private string cliente = string.Empty;

    [ObservableProperty]
    private string planta = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ValorTotal))]
    private int quantidade;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ValorTotal))]
    private decimal valor;

    [ObservableProperty]
    private string formaPagamento = "Dinheiro";

    public decimal ValorTotal => Quantidade * Valor;
}