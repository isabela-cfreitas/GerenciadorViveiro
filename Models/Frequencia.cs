using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.Models;

public partial class Frequencia(string produto = "", int quantidade= 0, decimal precoU = 0) : ObservableObject {
    [ObservableProperty]
    private string produto = produto;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PrecoT))]
    private int quantidade = quantidade;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PrecoT))]
    private decimal precoU = precoU;

    public decimal PrecoT => Quantidade * PrecoU;
}