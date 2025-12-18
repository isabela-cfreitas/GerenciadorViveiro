using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using GerenciadorViveiro.ViewModels.Interfaces;

namespace GerenciadorViveiro.ViewModels;

public partial class BalancoViewModel : ObservableObject, IEditableGridViewModel
{
    private readonly VendasViewModel _vendasVM;
    private readonly CustosViewModel _custosVM;
    private string PastaBalancos => "C:/Users/isaca/Dropbox/Balancos";

    private List<ItemBalanco> _clipboard = new();

    [ObservableProperty]
    private ObservableCollection<ItemBalanco> balancosSelecionados = new();

    [ObservableProperty]
    private ObservableCollection<ItemBalanco> balancos = new();

    [ObservableProperty]
    private int anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private int mesSelecionado = DateTime.Today.Month;

    public BalancoViewModel(VendasViewModel vendasVM, CustosViewModel custosVM)
    {
        _vendasVM = vendasVM;
        _custosVM = custosVM;

        Directory.CreateDirectory(PastaBalancos);
        CarregarBalancos();

        // Recalcula quando vendas mudam
        _vendasVM.VendasAlteradas += (s, e) => AtualizarTotais();
    }

    private string CaminhoArquivo =>
        Path.Combine(PastaBalancos, $"balanco_{AnoSelecionado}_{MesSelecionado:00}.xlsx");

    partial void OnAnoSelecionadoChanged(int value) => CarregarBalancos();
    
    partial void OnMesSelecionadoChanged(int value)
    {
        if (value < 1 || value > 12)
            return;
        CarregarBalancos();
    }

    private void AtualizarTotais()
    {
        // Atualiza rendas e custos baseado nos dados reais
        var rendaBruta = _vendasVM.ObterTotalVendasPorPeriodo(AnoSelecionado, MesSelecionado);
        var custoTotal = _custosVM.Custos.Sum(c => c.ValorTotal);
        var margemLucro = rendaBruta - custoTotal;

        // Atualiza a primeira linha se existir
        if (Balancos.Count > 0)
        {
            var linha = Balancos[0];
            linha.RendaBruta = rendaBruta;
            linha.CustoTotal = custoTotal;
            linha.MargemLucro = margemLucro;
        }
    }

    public void CarregarBalancos()
    {
        Balancos.Clear();

        if (MesSelecionado < 1 || MesSelecionado > 12)
            return;

        if (!File.Exists(CaminhoArquivo))
        {
            CriarArquivo();
            return;
        }

        try
        {
            using var workbook = new XLWorkbook(CaminhoArquivo);
            var ws = workbook.Worksheet(1);

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                // Função auxiliar para ler decimal com segurança
                decimal GetDecimalSafe(int coluna)
                {
                    var cell = row.Cell(coluna);
                    if (cell.IsEmpty())
                        return 0m;
                    
                    if (decimal.TryParse(cell.GetString(), out decimal valor))
                        return valor;
                    
                    try
                    {
                        return cell.GetValue<decimal>();
                    }
                    catch
                    {
                        return 0m;
                    }
                }

                Balancos.Add(new ItemBalanco
                {
                    Descricao = row.Cell(1).GetString(),
                    RendaBruta = GetDecimalSafe(2),
                    CustoTotal = GetDecimalSafe(3),
                    MargemLucro = GetDecimalSafe(4),
                    PercentualLuis = GetDecimalSafe(5),
                    PercentualPedro = GetDecimalSafe(6),
                    PercentualViveiro = GetDecimalSafe(7)
                });
            }

            AtualizarTotais();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao carregar balanços: {ex.Message}");
            CriarArquivo();
        }
    }

    [RelayCommand]
    private void SalvarBalancos()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Balanço");

        ws.Cell(1, 1).Value = "Descrição";
        ws.Cell(1, 2).Value = "Renda Bruta";
        ws.Cell(1, 3).Value = "Custo Total";
        ws.Cell(1, 4).Value = "Margem Lucro";
        ws.Cell(1, 5).Value = "% Luís";
        ws.Cell(1, 6).Value = "% Pedro";
        ws.Cell(1, 7).Value = "% Viveiro";
        ws.Cell(1, 8).Value = "R$ Luís";
        ws.Cell(1, 9).Value = "R$ Pedro";
        ws.Cell(1, 10).Value = "R$ Viveiro";

        ws.Range(1, 1, 1, 10).Style.Font.Bold = true;

        int linha = 2;
        foreach (var b in Balancos)
        {
            ws.Cell(linha, 1).Value = b.Descricao;
            ws.Cell(linha, 2).Value = b.RendaBruta;
            ws.Cell(linha, 3).Value = b.CustoTotal;
            ws.Cell(linha, 4).Value = b.MargemLucro;
            ws.Cell(linha, 5).Value = b.PercentualLuis;
            ws.Cell(linha, 6).Value = b.PercentualPedro;
            ws.Cell(linha, 7).Value = b.PercentualViveiro;
            ws.Cell(linha, 8).Value = b.ValorLuis;
            ws.Cell(linha, 9).Value = b.ValorPedro;
            ws.Cell(linha, 10).Value = b.ValorViveiro;
            linha++;
        }

        ws.Columns().AdjustToContents();
        workbook.SaveAs(CaminhoArquivo);
    }

    [RelayCommand]
    public void NovaLinha()
    {
        Balancos.Add(new ItemBalanco 
        { 
            Descricao = "Divisão Lucros",
            PercentualLuis = 40m,
            PercentualPedro = 40m,
            PercentualViveiro = 20m
        });
        AtualizarTotais();
    }

    [RelayCommand]
    private void ExcluirLinhas()
    {
        if (BalancosSelecionados == null || BalancosSelecionados.Count == 0)
            return;

        var paraRemover = BalancosSelecionados.ToList();
        foreach (var balanco in paraRemover)
        {
            Balancos.Remove(balanco);
        }
    }

    private void CriarArquivo()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Balanço");

        ws.Cell(1, 1).Value = "Descrição";
        ws.Cell(1, 2).Value = "Renda Bruta";
        ws.Cell(1, 3).Value = "Custo Total";
        ws.Cell(1, 4).Value = "Margem Lucro";
        ws.Cell(1, 5).Value = "% Luís";
        ws.Cell(1, 6).Value = "% Pedro";
        ws.Cell(1, 7).Value = "% Viveiro";
        ws.Cell(1, 8).Value = "R$ Luís";
        ws.Cell(1, 9).Value = "R$ Pedro";
        ws.Cell(1, 10).Value = "R$ Viveiro";

        ws.Range(1, 1, 1, 10).Style.Font.Bold = true;

        // Cria linha inicial
        ws.Cell(2, 1).Value = "Divisão Lucros";
        ws.Cell(2, 5).Value = 40;
        ws.Cell(2, 6).Value = 40;
        ws.Cell(2, 7).Value = 20;

        wb.SaveAs(CaminhoArquivo);

        // Carrega a linha criada
        Balancos.Add(new ItemBalanco
        {
            Descricao = "Divisão Lucros",
            PercentualLuis = 40m,
            PercentualPedro = 40m,
            PercentualViveiro = 20m
        });
        AtualizarTotais();
    }

    private ItemBalanco ClonarBalanco(ItemBalanco b)
    {
        return new ItemBalanco
        {
            Descricao = b.Descricao,
            RendaBruta = b.RendaBruta,
            CustoTotal = b.CustoTotal,
            MargemLucro = b.MargemLucro,
            PercentualLuis = b.PercentualLuis,
            PercentualPedro = b.PercentualPedro,
            PercentualViveiro = b.PercentualViveiro
        };
    }

    public void CopiarSelecionados()
    {
        _clipboard.Clear();
        foreach (var balanco in BalancosSelecionados)
        {
            _clipboard.Add(ClonarBalanco(balanco));
        }
    }

    public void Colar(int index)
    {
        if (_clipboard.Count == 0)
            return;

        int insertIndex = index + 1;
        foreach (var balanco in _clipboard)
        {
            Balancos.Insert(insertIndex++, ClonarBalanco(balanco));
        }
    }

    public void RecortarSelecionados()
    {
        CopiarSelecionados();
        var paraRemover = BalancosSelecionados.ToList();
        foreach (var balanco in paraRemover)
        {
            Balancos.Remove(balanco);
        }
    }

    public void ApagarSelecionados()
    {
        var paraRemover = BalancosSelecionados.ToList();
        foreach (var balanco in paraRemover)
        {
            Balancos.Remove(balanco);
        }
    }

    public void Salvar() => SalvarBalancos();

    // Classe interna - não cria arquivo separado
    public partial class ItemBalanco : ObservableObject
    {
        [ObservableProperty]
        private string descricao = string.Empty;

        [ObservableProperty]
        private decimal rendaBruta;

        [ObservableProperty]
        private decimal custoTotal;

        [ObservableProperty]
        private decimal margemLucro;

        [ObservableProperty]
        private decimal percentualLuis = 40m;

        [ObservableProperty]
        private decimal percentualPedro = 40m;

        [ObservableProperty]
        private decimal percentualViveiro = 20m;

        public decimal ValorLuis => MargemLucro * (PercentualLuis / 100);
        public decimal ValorPedro => MargemLucro * (PercentualPedro / 100);
        public decimal ValorViveiro => MargemLucro * (PercentualViveiro / 100);

        partial void OnMargemLucroChanged(decimal value)
        {
            OnPropertyChanged(nameof(ValorLuis));
            OnPropertyChanged(nameof(ValorPedro));
            OnPropertyChanged(nameof(ValorViveiro));
        }

        partial void OnPercentualLuisChanged(decimal value)
        {
            OnPropertyChanged(nameof(ValorLuis));
        }

        partial void OnPercentualPedroChanged(decimal value)
        {
            OnPropertyChanged(nameof(ValorPedro));
        }

        partial void OnPercentualViveiroChanged(decimal value)
        {
            OnPropertyChanged(nameof(ValorViveiro));
        }
    }
}