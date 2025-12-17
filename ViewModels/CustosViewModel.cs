using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using GerenciadorViveiro.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using GerenciadorViveiro.ViewModels.Interfaces;

namespace GerenciadorViveiro.ViewModels;

public partial class CustosViewModel : ObservableObject, IEditableGridViewModel
{
    private string PastaCustos => "C:/Users/isaca/Dropbox/Custos";

    private List<Custo> _clipboard = new();

    [ObservableProperty]
    private ObservableCollection<Custo> custosSelecionados = new();

    [ObservableProperty]
    private ObservableCollection<Custo> custos = new();

    [ObservableProperty]
    private int anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private int mesSelecionado = DateTime.Today.Month;

    public CustosViewModel()
    {
        Directory.CreateDirectory(PastaCustos);
        CarregarCustos();
    }

    private string CaminhoArquivo =>
        Path.Combine(PastaCustos, $"custos_{AnoSelecionado}_{MesSelecionado:00}.xlsx");

    partial void OnAnoSelecionadoChanged(int value) => CarregarCustos();
    partial void OnMesSelecionadoChanged(int value){
        if (value < 1 || value > 12)
            return;

        CarregarCustos();
    }

    public void CarregarCustos(){
        Custos.Clear();

        if (MesSelecionado < 1 || MesSelecionado > 12)
            return;

        if (!File.Exists(CaminhoArquivo))
        {
            CriarArquivo();
            return;
        }

        using var workbook = new XLWorkbook(CaminhoArquivo);
        var ws = workbook.Worksheet(1);

        foreach (var row in ws.RowsUsed().Skip(1))
        {
            Custos.Add(new Custo
            {
                Atividade = row.Cell(1).GetString(),
                Elemento = row.Cell(2).GetString(),
                Quantidade = row.Cell(3).GetValue<int>(),
                ValorTotal = row.Cell(4).GetValue<decimal>()
            });
        }
    }

    [RelayCommand]
    private void SalvarCustos()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Custos");

        ws.Cell(1, 1).Value = "Atividade";
        ws.Cell(1, 2).Value = "Elemento";
        ws.Cell(1, 3).Value = "Quantidade";
        ws.Cell(1, 4).Value = "Valor Total";

        ws.Range(1, 1, 1, 4).Style.Font.Bold = true;

        int linha = 2;
        foreach (var c in Custos)
        {
            ws.Cell(linha, 1).Value = c.Atividade;
            ws.Cell(linha, 2).Value = c.Elemento;
            ws.Cell(linha, 3).Value = c.Quantidade;
            ws.Cell(linha, 4).Value = c.ValorTotal;
            linha++;
        }

        ws.Columns().AdjustToContents();
        workbook.SaveAs(CaminhoArquivo);
    }

    [RelayCommand]
    public void NovaLinha()
    {
        Custos.Add(new Custo());
    }

    [RelayCommand]
    private void ExcluirLinhas()
    {
        if (CustosSelecionados == null || CustosSelecionados.Count == 0)
            return;

        var paraRemover = CustosSelecionados.ToList();
        foreach (var custo in paraRemover)
        {
            Custos.Remove(custo);
        }
    }

    private void CriarArquivo()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Custos");

        ws.Cell(1, 1).Value = "Atividade";
        ws.Cell(1, 2).Value = "Elemento";
        ws.Cell(1, 3).Value = "Quantidade";
        ws.Cell(1, 4).Value = "Valor Total";

        ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
        wb.SaveAs(CaminhoArquivo);
    }

    private Custo ClonarCusto(Custo c)
    {
        return new Custo
        {
            Atividade = c.Atividade,
            Elemento = c.Elemento,
            Quantidade = c.Quantidade,
            ValorTotal = c.ValorTotal
        };
    }

    public void CopiarSelecionados(){
        _clipboard.Clear();

        foreach (var custo in CustosSelecionados)
        {
            _clipboard.Add(ClonarCusto(custo));
        }
    }

   public void Colar(int index){
        if (_clipboard.Count == 0)
            return;

        int insertIndex = index + 1;

        foreach (var custo in _clipboard)
        {
            Custos.Insert(insertIndex++, ClonarCusto(custo));
        }
    }

    public void RecortarSelecionados(){
        CopiarSelecionados();

        var paraRemover = CustosSelecionados.ToList();
        foreach (var custo in paraRemover)
        {
            Custos.Remove(custo);
        }
    }

    public void ApagarSelecionados(){
        var paraRemover = CustosSelecionados.ToList();
        foreach (var custo in paraRemover)
        {
            Custos.Remove(custo);
        }
    }

    public void Salvar() => SalvarCustos();
}