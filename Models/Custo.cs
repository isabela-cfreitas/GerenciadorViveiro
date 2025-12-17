using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.Models;

public partial class Custo : ObservableObject
{
    [ObservableProperty]
    private string atividade = string.Empty;

    [ObservableProperty]
    private string elemento = string.Empty;

    [ObservableProperty]
    private int quantidade;

    [ObservableProperty]
    private decimal valorTotal;
}
