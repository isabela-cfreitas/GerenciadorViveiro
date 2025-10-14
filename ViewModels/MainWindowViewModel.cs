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

namespace GerenciadorViveiro.ViewModels;

public partial class MainWindowViewModel : ObservableObject {
    private readonly string caminhoArquivo = "vendas.xlsx";

    [ObservableProperty]
    private ObservableCollection<Venda> vendasSelecionadas = new();

    [ObservableProperty]
    private ObservableCollection<Venda> vendas = new();

    

    public MainWindowViewModel() {
        CarregarVendas();

        // Salva automaticamente quando a coleção muda
        Vendas.CollectionChanged += (s, e) => SalvarVendas();

        // Salva quando edita propriedades de um item
        foreach (var venda in Vendas) {
            venda.PropertyChanged += (s, e) => SalvarVendas();
    }
    }


    [RelayCommand]
    private void AdicionarLinha() {
        var novaVenda = new Venda {
            Data = DateTime.Today,
            Cliente = string.Empty,
            Produto = string.Empty,
            Quantidade = 0,
            PrecoU = 0,
            FormaPagamento = "Dinheiro"
        };
        
        // Conecta o evento PropertyChanged para salvar automaticamente
        novaVenda.PropertyChanged += (s, e) => SalvarVendas();
        
        Vendas.Add(novaVenda);
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
    }

    [RelayCommand]
    private void Salvar() {
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
                        Cliente = row.Cell(2).GetString(),
                        Produto = row.Cell(3).GetString(),
                        Quantidade = row.Cell(4).GetValue<int>(),
                        PrecoU = row.Cell(5).GetValue<decimal>(),
                        FormaPagamento = row.Cell(6).GetString()
                    };

                    // Conecta evento para salvar automaticamente ao editar
                    venda.PropertyChanged += (s, e) => SalvarVendas();

                    Vendas.Add(venda);
                }
            }
        }
        catch (Exception ex) {
            Console.WriteLine($"Erro ao carregar vendas: {ex.Message}");
        }
    }

    private void SalvarVendas() {
        try {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Vendas");

            // Cabeçalhos
            worksheet.Cell(1, 1).Value = "Data";
            worksheet.Cell(1, 2).Value = "Cliente";
            worksheet.Cell(1, 3).Value = "Produto";
            worksheet.Cell(1, 4).Value = "Quantidade";
            worksheet.Cell(1, 5).Value = "Preço";
            worksheet.Cell(1, 6).Value = "Forma de Pagamento";

            // Estilizar cabeçalho
            var headerRange = worksheet.Range(1, 1, 1, 6);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            // Dados
            int linha = 2;
            foreach (var venda in Vendas) {
                worksheet.Cell(linha, 1).Value = venda.Data;
                worksheet.Cell(linha, 2).Value = venda.Cliente;
                worksheet.Cell(linha, 3).Value = venda.Produto;
                worksheet.Cell(linha, 4).Value = venda.Quantidade;
                worksheet.Cell(linha, 5).Value = venda.PrecoU;
                worksheet.Cell(linha, 6).Value = venda.FormaPagamento;
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

    private void CriarArquivoExcel() {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Vendas");

        worksheet.Cell(1, 1).Value = "Data";
        worksheet.Cell(1, 2).Value = "Cliente";
        worksheet.Cell(1, 3).Value = "Produto";
        worksheet.Cell(1, 4).Value = "Quantidade";
        worksheet.Cell(1, 5).Value = "Preço";
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
}