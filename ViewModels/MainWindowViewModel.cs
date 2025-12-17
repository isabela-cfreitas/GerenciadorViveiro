using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using GerenciadorViveiro.Models;
using System.Diagnostics;
using Avalonia.Logging;
using Avalonia.Controls;
using System.Collections.Generic;
using Avalonia.Interactivity;
using Avalonia.Input;

using GerenciadorViveiro.ViewModels.Interfaces;

namespace GerenciadorViveiro.ViewModels;


public partial class MainWindowViewModel : ObservableObject, IEditableGridViewModel {
    private readonly string caminhoArquivo = "C:/Users/isaca/Dropbox/vendas.xlsx";

    private List<Venda> _clipboardVendas = new();

    [ObservableProperty]
    private CustosViewModel custosVM = new();

    [ObservableProperty]
    private ObservableCollection<Venda> vendasSelecionadas = new();

    [ObservableProperty]
    private string filtroCliente = string.Empty;

    [ObservableProperty]
    private ObservableCollection<Venda> vendas = new();

    [ObservableProperty]
    private ObservableCollection<Venda> vendasFiltradas = new();

    [ObservableProperty]
    private ObservableCollection<Frequencia> frequencias = new();

    [ObservableProperty]
    private int? anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private int mesSelecionado = DateTime.Today.Month;

    [ObservableProperty]
    private ObservableCollection<int> anosDisponiveis = new();

    [ObservableProperty]
    private ObservableCollection<int> mesesDisponiveis = new(
        Enumerable.Range(1, 12)
    );

    public MainWindowViewModel() {
        CarregarVendas();
        AtualizarAnosDisponiveis();

        // Salva automaticamente quando a coleção muda
        Vendas.CollectionChanged += (s, e) => {
            SalvarVendas();
            AtualizarAnosDisponiveis();
            AplicarFiltro();
        };

        // Salva quando edita propriedades de um item
        foreach (var venda in Vendas) {
            venda.PropertyChanged += (s, e) => SalvarVendas();
        }

        AplicarFiltro();
    }

    private void AtualizarAnosDisponiveis() {
        AnosDisponiveis.Clear();

        var anos = Vendas
            .Select(v => v.Data.Year)
            .Distinct()
            .OrderBy(a => a);

        foreach (var ano in anos) {
            AnosDisponiveis.Add(ano);
        }

        if (
            AnoSelecionado == null ||
            !AnosDisponiveis.Contains(AnoSelecionado.Value)
        ) {
            if (AnosDisponiveis.Any())
                AnoSelecionado = AnosDisponiveis.Last();
        }
    }

    partial void OnAnoSelecionadoChanged(int? value) {
        CalcularFrequencias();
    }

    partial void OnMesSelecionadoChanged(int value) {
        if (value < 1 || value > 12)
            return;

        CalcularFrequencias();
    }

    partial void OnFiltroClienteChanged(string value) {
    AplicarFiltro();
    }

    private void AplicarFiltro() {
        VendasFiltradas.Clear();

        var filtradas = string.IsNullOrWhiteSpace(FiltroCliente)
            ? Vendas
            : Vendas.Where(v =>
                !string.IsNullOrWhiteSpace(v.Cliente) &&
                v.Cliente.Contains(FiltroCliente, StringComparison.OrdinalIgnoreCase)
            );

        foreach (var venda in filtradas) {
            VendasFiltradas.Add(venda);
        }
    }

    private Venda ClonarVenda(Venda v) {
        var nova = new Venda {
            Data = v.Data,
            Planta = v.Planta,
            Quantidade = v.Quantidade,
            Valor = v.Valor,
            Cliente = v.Cliente,
            FormaPagamento = v.FormaPagamento
        };

        nova.PropertyChanged += (s, e) => {
            SalvarVendas();
            if (e.PropertyName == nameof(Venda.Data))
                AtualizarAnosDisponiveis();
        };

        return nova;
    }

    public void CopiarLinhasSelecionadas() {
        _clipboardVendas.Clear();

        foreach (var venda in VendasSelecionadas) {
            _clipboardVendas.Add(ClonarVenda(venda));
        }
    }

    public void ColarLinhas(int indiceBase) {
        if (_clipboardVendas.Count == 0)
            return;

        int insertIndex = indiceBase + 1;

        foreach (var venda in _clipboardVendas) {
            Vendas.Insert(insertIndex++, ClonarVenda(venda));
        }

        AplicarFiltro();
    }

    public void RecortarLinhasSelecionadas() {
        CopiarLinhasSelecionadas();

        var paraRemover = VendasSelecionadas.ToList();
        foreach (var venda in paraRemover) {
            Vendas.Remove(venda);
        }

        AplicarFiltro();
    }

    public void ApagarLinhasSelecionadas () {
        var paraRemover = VendasSelecionadas.ToList();
        foreach (var venda in paraRemover) {
            Vendas.Remove(venda);
        }

        AplicarFiltro();
    }

    [RelayCommand]
    public void AdicionarLinha() {
        var novaVenda = new Venda {
            Data = DateTime.Today,
            Cliente = string.Empty,
            Planta = string.Empty,
            Quantidade = 0,
            Valor = 0,
            FormaPagamento = "Dinheiro"
        };
        
        // Conecta o evento PropertyChanged para salvar automaticamente
        novaVenda.PropertyChanged += (s, e) => {
            SalvarVendas();

            if (e.PropertyName == nameof(Venda.Data)) {
                AtualizarAnosDisponiveis();
            }
        };
        
        Vendas.Add(novaVenda);
        AplicarFiltro();
    }

    [RelayCommand]
    private async Task ExcluirLinhas() {
        if (VendasSelecionadas == null || VendasSelecionadas.Count == 0) {
            await MostrarMensagem("Aviso", "Nenhuma venda selecionada.");
            return;
        }

        var resultado = await MostrarConfirmacao("Confirmação", 
            $"Deseja realmente excluir {VendasSelecionadas.Count} venda(s)?");

        if (resultado) {
            var vendasParaRemover = VendasSelecionadas.ToList();
            foreach (var venda in vendasParaRemover) {
                Vendas.Remove(venda);
            }
            SalvarVendas();
        }
        AplicarFiltro();

    }

    [RelayCommand]
    public void Salvar() {
        SalvarVendas();
    }

    private void CarregarVendas() {
        Vendas.Clear();

        if (!File.Exists(caminhoArquivo)) {
            CriarArquivoExcel();
            return;
        }

        try {
            using var workbook = new XLWorkbook(caminhoArquivo);
            var worksheet = workbook.Worksheet(1);

            var rangeUsed = worksheet.RangeUsed();
            if (rangeUsed != null && !rangeUsed.IsEmpty()) {
                var rows = rangeUsed.RowsUsed().Skip(1); // Pula o cabeçalho

                foreach (var row in rows) {
                    var venda = new Venda {
                        Data = row.Cell(1).GetDateTime(),
                        Planta = row.Cell(2).GetString(),
                        Quantidade = row.Cell(3).GetValue<int>(),
                        Valor = row.Cell(4).GetValue<decimal>(),
                        Cliente = row.Cell(6).GetString(),
                        FormaPagamento = row.Cell(7).GetString()
                    };

                    // Conecta evento para salvar automaticamente ao editar
                    venda.PropertyChanged += (s, e) => {
                        SalvarVendas();

                        if (e.PropertyName == nameof(Venda.Data)) {
                            AtualizarAnosDisponiveis();
                        }
                    };

                    Vendas.Add(venda);
                }
            }
        }
        catch (Exception ex) {
            Console.WriteLine($"Erro ao carregar vendas: {ex.Message}");
        }

        AplicarFiltro();
    }

    private void SalvarVendas() {
        try {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Vendas");
            // Cabeçalhos
            worksheet.Cell(1, 1).Value = "Data";
            worksheet.Cell(1, 2).Value = "Planta";
            worksheet.Cell(1, 3).Value = "Quantidade";
            worksheet.Cell(1, 4).Value = "Valor";
            worksheet.Cell(1, 5).Value = "Valor Total";
            worksheet.Cell(1, 6).Value = "Cliente";
            worksheet.Cell(1, 7).Value = "Forma de Pagamento";

            // Estilizar cabeçalho
            var headerRange = worksheet.Range(1, 1, 1, 7);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            // Dados
            int linha = 2;
            foreach (var venda in Vendas) {
                worksheet.Cell(linha, 1).Value = venda.Data;
                worksheet.Cell(linha, 2).Value = venda.Planta;
                worksheet.Cell(linha, 3).Value = venda.Quantidade;
                worksheet.Cell(linha, 4).Value = venda.Valor;
                worksheet.Cell(linha, 5).Value = venda.ValorTotal;
                worksheet.Cell(linha, 6).Value = venda.Cliente;
                worksheet.Cell(linha, 7).Value = venda.FormaPagamento;
                linha++;
            }

            // Ajustar largura das colunas
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(caminhoArquivo);
        }
        catch (Exception ex) {
            Console.WriteLine($"Erro ao salvar vendas: {ex.Message}");
        }
    }

    private void SalvarFrequencias() {
        try {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Frequências");

            // Cabeçalhos
            worksheet.Cell(1, 1).Value = "Planta";
            worksheet.Cell(1, 2).Value = "Quantidade";
            worksheet.Cell(1, 3).Value = "Valor Unitário";
            worksheet.Cell(1, 4).Value = "Valor total";

            // Estilizar cabeçalho
            var headerRange = worksheet.Range(1, 1, 1, 4);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            // Dados
            int linha = 2;
            foreach (var frequencia in Frequencias) {
                worksheet.Cell(linha, 1).Value = frequencia.Planta;
                worksheet.Cell(linha, 2).Value = frequencia.Quantidade;
                worksheet.Cell(linha, 3).Value = frequencia.Valor;
                worksheet.Cell(linha, 4).Value = frequencia.ValorTotal;
                linha++;
            }

            // Ajustar largura das colunas
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(caminhoArquivo);
        }
        catch (Exception ex) {
            Console.WriteLine($"Erro ao salvar vendas: {ex.Message}");
        }
    }

    public void CalcularFrequencias() {
        if (AnoSelecionado == null) return;

        Dictionary<string, Frequencia> dict = new();

        var vendasPeriodo = Vendas.Where(v =>
            v.Data.Year == AnoSelecionado &&
            v.Data.Month == MesSelecionado
        );

        foreach (var venda in vendasPeriodo) {
            if (venda.Quantidade == 0) continue;

            if (!dict.TryGetValue(venda.Planta, out Frequencia? freq)) {
                dict.Add(
                    venda.Planta,
                    new Frequencia(venda.Planta, venda.Quantidade, venda.Valor)
                );
            }
            else {
                freq.Valor = freq.ValorTotal + venda.ValorTotal;
                freq.Quantidade += venda.Quantidade;
                freq.Valor /= freq.Quantidade;
            }
        }

        Frequencias = new(dict.Values);
    }

    private void CriarArquivoExcel() {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Vendas");

        worksheet.Cell(1, 1).Value = "Data";
        worksheet.Cell(1, 2).Value = "Cliente";
        worksheet.Cell(1, 3).Value = "Planta";
        worksheet.Cell(1, 4).Value = "Quantidade";
        worksheet.Cell(1, 5).Value = "Valor";
        worksheet.Cell(1, 6).Value = "Forma de Pagamento";

        // Estilizar cabeçalho
        var headerRange = worksheet.Range(1, 1, 1, 6);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

        workbook.SaveAs(caminhoArquivo);
    }

    private async Task MostrarMensagem(string titulo, string mensagem) {
    var box = MessageBoxManager.GetMessageBoxStandard(
        new MsBox.Avalonia.Dto.MessageBoxStandardParams {
            ButtonDefinitions = ButtonEnum.Ok,
            ContentTitle = titulo,
            ContentMessage = mensagem,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            MaxWidth = 500
        });
    await box.ShowAsync();
}

    private async Task<bool> MostrarConfirmacao(string titulo, string mensagem) {
    var box = MessageBoxManager.GetMessageBoxStandard(
        new MsBox.Avalonia.Dto.MessageBoxStandardParams {
            ButtonDefinitions = ButtonEnum.YesNo,
            ContentTitle = titulo,
            ContentMessage = mensagem,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.WidthAndHeight,
            MaxWidth = 500
        });
    var result = await box.ShowAsync();
    return result == ButtonResult.Yes;
}

public void ApagarSelecionados() => ApagarLinhasSelecionadas();
public void CopiarSelecionados() => CopiarLinhasSelecionadas();
public void RecortarSelecionados() => RecortarLinhasSelecionadas();
public void Colar(int index) => ColarLinhas(index);
public void NovaLinha() => AdicionarLinha();

}