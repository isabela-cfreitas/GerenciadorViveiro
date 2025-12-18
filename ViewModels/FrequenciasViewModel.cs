using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using GerenciadorViveiro.Models;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using Avalonia.Controls;

namespace GerenciadorViveiro.ViewModels;

public partial class FrequenciasViewModel : ObservableObject
{
    private readonly VendasViewModel _vendasViewModel;
    private readonly ConfiguracoesViewModel _configuracoesVM;

    [ObservableProperty]
    private ObservableCollection<Frequencia> frequencias = new();

    [ObservableProperty]
    private int? anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private int mesSelecionado = DateTime.Today.Month;

    [ObservableProperty]
    private string? mensagemErro;

    public FrequenciasViewModel(VendasViewModel vendasViewModel, ConfiguracoesViewModel configuracoesVM)
    {
        _vendasViewModel = vendasViewModel;
        _configuracoesVM = configuracoesVM;

        //atualiza pq vendas mudaram, observador foi notificado!!
        _vendasViewModel.VendasAlteradas += (s, e) => CalcularFrequencias();
        
        _vendasViewModel.AnosAtualizados += (s, e) =>
        {
            //evitando erro de ano inválido
            if (AnoSelecionado == null || !_vendasViewModel.AnosDisponiveis.Contains(AnoSelecionado.Value))
            {
                if (_vendasViewModel.AnosDisponiveis.Any())
                    AnoSelecionado = _vendasViewModel.AnosDisponiveis.Last();
            }
        };

        //calcula quais sao as frequencias a partir do arquivo de vendas
        CalcularFrequencias();
    }

    partial void OnAnoSelecionadoChanged(int? value)
    {
        if (value == null)
        {
            MensagemErro = "Ano não pode ser vazio.";
            return;
        }

        if (!ValidarAno(value.Value))
        {
            MensagemErro = "Ano inválido. Digite um ano entre 1900 e 2100.";
            return;
        }

        MensagemErro = null;
        CalcularFrequencias();
    }

    partial void OnMesSelecionadoChanged(int value)
    {
        if (!ValidarMes(value))
        {
            MensagemErro = "Mês inválido. Digite um mês entre 1 e 12.";
            return;
        }

        MensagemErro = null;
        CalcularFrequencias();
    }

    private bool ValidarAno(int ano)
    {
        return ano >= 1900 && ano <= 2100;
    }

    private bool ValidarMes(int mes)
    {
        return mes >= 1 && mes <= 12;
    }

    public void CalcularFrequencias()
    {
        if (AnoSelecionado == null || !ValidarAno(AnoSelecionado.Value) || !ValidarMes(MesSelecionado))
        {
            Frequencias.Clear();
            return;
        }

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

    [RelayCommand]
    private async Task ExportarFrequencias()
    {
        if (!_configuracoesVM.PastaConfigurada)
        {
            await MostrarMensagem("Erro", "Configure a pasta base nas Configurações primeiro.");
            return;
        }

        if (Frequencias.Count == 0)
        {
            await MostrarMensagem("Aviso", "Não há frequências para exportar.");
            return;
        }

        try
        {
            _configuracoesVM.CriarPastas();

            var caminhoArquivo = Path.Combine(
                _configuracoesVM.PastaFrequencias, 
                $"frequencias_{AnoSelecionado}_{MesSelecionado:00}.xlsx"
            );

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Frequências");

            //cabeçalhos da tabela
            ws.Cell(1, 1).Value = "Planta";
            ws.Cell(1, 2).Value = "Quantidade";
            ws.Cell(1, 3).Value = "Valor";
            ws.Cell(1, 4).Value = "Valor Total";

            ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
            ws.Range(1, 1, 1, 4).Style.Fill.BackgroundColor = XLColor.LightGray;

            int linha = 2;
            foreach (var freq in Frequencias)
            {
                ws.Cell(linha, 1).Value = freq.Planta;
                ws.Cell(linha, 2).Value = freq.Quantidade;
                ws.Cell(linha, 3).Value = freq.Valor;
                ws.Cell(linha, 4).Value = freq.ValorTotal;
                linha++;
            }

            ws.Columns().AdjustToContents();
            workbook.SaveAs(caminhoArquivo);

            await MostrarMensagem("Sucesso", $"Frequências exportadas com sucesso!\n{caminhoArquivo}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao exportar frequências: {ex.Message}");
            await MostrarMensagem("Erro", $"Erro ao exportar: {ex.Message}");
        }
    }

    private async Task MostrarMensagem(string titulo, string mensagem){
        var box = MessageBoxManager.GetMessageBoxStandard(
            new MsBox.Avalonia.Dto.MessageBoxStandardParams
            {
                ButtonDefinitions = ButtonEnum.Ok,
                ContentTitle = titulo,
                ContentMessage = mensagem,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                SizeToContent = SizeToContent.WidthAndHeight,
                MaxWidth = 500
            });
        await box.ShowAsync();
    }

    public ObservableCollection<int> AnosDisponiveis => _vendasViewModel.AnosDisponiveis;
    
    public ObservableCollection<int> MesesDisponiveis { get; } = new(Enumerable.Range(1, 12));
}