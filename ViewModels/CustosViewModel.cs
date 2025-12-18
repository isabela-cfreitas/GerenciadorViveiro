using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using GerenciadorViveiro.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using GerenciadorViveiro.ViewModels.Interfaces;

namespace GerenciadorViveiro.ViewModels;

public partial class CustosViewModel : ObservableObject, IEditableGridViewModel
{
    private readonly ConfiguracoesViewModel _configuracoesVM;
    private List<Custo> _clipboard = new();

    [ObservableProperty]
    private ObservableCollection<Custo> custosSelecionados = new();

    [ObservableProperty]
    private ObservableCollection<Custo> custos = new();

    [ObservableProperty]
    private int anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private int mesSelecionado = DateTime.Today.Month;

    [ObservableProperty]
    private string? mensagemErro;

    //observer
    public event EventHandler? CustosAlterados;

    public CustosViewModel(ConfiguracoesViewModel configuracoesVM){
        _configuracoesVM = configuracoesVM;
        
        //recarrega quando os caminhos mudarem
        _configuracoesVM.PastasAlteradas += (s, e) =>
        {
            CarregarCustos();
        };

        _configuracoesVM.CriarPastas();
        CarregarCustos();
    }

    private string CaminhoArquivo =>
        Path.Combine(_configuracoesVM.PastaCustos, $"custos_{AnoSelecionado}_{MesSelecionado:00}.xlsx");

    partial void OnAnoSelecionadoChanged(int value){
        if (!ValidarAno(value)){
            MensagemErro = "Ano inválido. Digite um ano entre 1900 e 2100.";
            return;
        }
        MensagemErro = null;
        CarregarCustos();
    }

    partial void OnMesSelecionadoChanged(int value){
        if (!ValidarMes(value)){
            MensagemErro = "Mês inválido. Digite um mês entre 1 e 12.";
            return;
        }
        MensagemErro = null;
        CarregarCustos();
    }

    private bool ValidarAno(int ano){
        return ano >= 1900 && ano <= 2100;
    }

    private bool ValidarMes(int mes){
        return mes >= 1 && mes <= 12;
    }

    public void CarregarCustos(){
        Custos.Clear();

        if (!ValidarAno(AnoSelecionado) || !ValidarMes(MesSelecionado))
            return;

        if (!File.Exists(CaminhoArquivo)){
            CriarArquivo();
            return;
        }

        try{
            using var workbook = new XLWorkbook(CaminhoArquivo);
            var ws = workbook.Worksheet(1);

            foreach (var row in ws.RowsUsed().Skip(1)){
                Custos.Add(new Custo
                {
                    Atividade = row.Cell(1).GetString(),
                    Elemento = row.Cell(2).GetString(),
                    Quantidade = row.Cell(3).GetValue<int>(),
                    ValorTotal = row.Cell(4).GetValue<decimal>()
                });
            }
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao carregar custos: {ex.Message}");
        }
    }

    [RelayCommand]
    private void SalvarCustos(){
        if (!ValidarAno(AnoSelecionado) || !ValidarMes(MesSelecionado)){
            MensagemErro = "Não é possível salvar com ano ou mês inválido.";
            return;
        }

        try{
            _configuracoesVM.CriarPastas();

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Custos");

            ws.Cell(1, 1).Value = "Atividade";
            ws.Cell(1, 2).Value = "Elemento";
            ws.Cell(1, 3).Value = "Quantidade";
            ws.Cell(1, 4).Value = "Valor Total";

            ws.Range(1, 1, 1, 4).Style.Font.Bold = true;

            int linha = 2;
            foreach (var c in Custos){
                ws.Cell(linha, 1).Value = c.Atividade;
                ws.Cell(linha, 2).Value = c.Elemento;
                ws.Cell(linha, 3).Value = c.Quantidade;
                ws.Cell(linha, 4).Value = c.ValorTotal;
                linha++;
            }

            ws.Columns().AdjustToContents();
            workbook.SaveAs(CaminhoArquivo);
            CustosAlterados?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao salvar custos: {ex.Message}");
            MensagemErro = $"Erro ao salvar: {ex.Message}";
        }
    }

    [RelayCommand]
    public void NovaLinha(){
        Custos.Add(new Custo());
        CustosAlterados?.Invoke(this, EventArgs.Empty);
    }

    [RelayCommand]
    private void ExcluirLinhas(){
        if (CustosSelecionados == null || CustosSelecionados.Count == 0)
            return;

        var paraRemover = CustosSelecionados.ToList();
        foreach (var custo in paraRemover){
            Custos.Remove(custo);
        }
        CustosAlterados?.Invoke(this, EventArgs.Empty);
    }

    private void CriarArquivo(){
        //nao sobrescreve!!! improtante
        if (File.Exists(CaminhoArquivo)){
            Console.WriteLine($"Arquivo já existe, não será sobrescrito: {CaminhoArquivo}");
            return;
        }

        try{
            _configuracoesVM.CriarPastas();

            Console.WriteLine($"Criando novo arquivo: {CaminhoArquivo}");

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Custos");

            ws.Cell(1, 1).Value = "Atividade";
            ws.Cell(1, 2).Value = "Elemento";
            ws.Cell(1, 3).Value = "Quantidade";
            ws.Cell(1, 4).Value = "Valor Total";

            ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
            wb.SaveAs(CaminhoArquivo);
            Console.WriteLine("Arquivo criado com sucesso!");
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao criar arquivo: {ex.Message}");
        }
    }

    private Custo ClonarCusto(Custo c){
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

        foreach (var custo in CustosSelecionados){
            _clipboard.Add(ClonarCusto(custo));
        }
    }

    public void Colar(int index){
        if (_clipboard.Count == 0)
            return;

        int insertIndex = index + 1;

        foreach (var custo in _clipboard){
            Custos.Insert(insertIndex++, ClonarCusto(custo));
        }
    }

    public void RecortarSelecionados(){
        CopiarSelecionados();

        var paraRemover = CustosSelecionados.ToList();
        foreach (var custo in paraRemover){
            Custos.Remove(custo);
        }
    }

    public void ApagarSelecionados(){
        var paraRemover = CustosSelecionados.ToList();
        foreach (var custo in paraRemover){
            Custos.Remove(custo);
        }
    }

    public Task SalvarAsync(){
        SalvarCustos();
        return Task.CompletedTask;
    }
}