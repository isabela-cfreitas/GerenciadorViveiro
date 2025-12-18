using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using GerenciadorViveiro.Models;

namespace GerenciadorViveiro.ViewModels;

public partial class FrequenciasViewModel : ObservableObject
{
    private readonly VendasViewModel _vendasViewModel;

    [ObservableProperty]
    private ObservableCollection<Frequencia> frequencias = new();

    [ObservableProperty]
    private int? anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private int mesSelecionado = DateTime.Today.Month;

    public FrequenciasViewModel(VendasViewModel vendasViewModel)
    {
        _vendasViewModel = vendasViewModel;

        // Atualiza quando vendas mudam
        _vendasViewModel.VendasAlteradas += (s, e) => CalcularFrequencias();
        
        // Atualiza anos disponíveis quando necessário
        _vendasViewModel.AnosAtualizados += (s, e) =>
        {
            // Verifica se o ano selecionado ainda é válido
            if (AnoSelecionado == null || !_vendasViewModel.AnosDisponiveis.Contains(AnoSelecionado.Value))
            {
                if (_vendasViewModel.AnosDisponiveis.Any())
                    AnoSelecionado = _vendasViewModel.AnosDisponiveis.Last();
            }
        };

        // Calcula frequências iniciais
        CalcularFrequencias();
    }

    partial void OnAnoSelecionadoChanged(int? value)
    {
        CalcularFrequencias();
    }

    partial void OnMesSelecionadoChanged(int value)
    {
        if (value < 1 || value > 12)
            return;

        CalcularFrequencias();
    }

    public void CalcularFrequencias()
    {
        if (AnoSelecionado == null) return;

        Dictionary<string, Frequencia> dict = new();

        var vendasPeriodo = _vendasViewModel.ObterVendasPorPeriodo(
            AnoSelecionado.Value, 
            MesSelecionado
        );

        foreach (var venda in vendasPeriodo)
        {
            if (venda.Quantidade == 0) continue;

            if (!dict.TryGetValue(venda.Planta, out Frequencia? freq))
            {
                dict.Add(
                    venda.Planta,
                    new Frequencia(venda.Planta, venda.Quantidade, venda.Valor)
                );
            }
            else
            {
                freq.Valor = freq.ValorTotal + venda.ValorTotal;
                freq.Quantidade += venda.Quantidade;
                freq.Valor /= freq.Quantidade;
            }
        }

        Frequencias = new(dict.Values);
    }

    // Propriedade auxiliar para binding no XAML
    public ObservableCollection<int> AnosDisponiveis => _vendasViewModel.AnosDisponiveis;
    
    public ObservableCollection<int> MesesDisponiveis { get; } = new(Enumerable.Range(1, 12));
}