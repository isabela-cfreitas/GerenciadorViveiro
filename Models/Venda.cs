using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.Models;

public partial class Venda : ObservableObject {
    [ObservableProperty]
    private DateTime data = DateTime.Today;

    [ObservableProperty]
    private string cliente = string.Empty;

    [ObservableProperty]
    private string produto = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PrecoT))]
    private int quantidade;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PrecoT))]
    private decimal precoU;

    [ObservableProperty]
    private string formaPagamento = "Dinheiro";

    public decimal PrecoT => Quantidade * PrecoU;
}