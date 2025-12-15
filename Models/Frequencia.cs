using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.Models;

public partial class Frequencia(string planta = "", int quantidade= 0, decimal valor = 0) : ObservableObject {
    [ObservableProperty]
    private string planta = planta;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ValorTotal))]
    private int quantidade = quantidade;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ValorTotal))]
    private decimal valor = valor;

    public decimal ValorTotal => Quantidade * Valor;
}