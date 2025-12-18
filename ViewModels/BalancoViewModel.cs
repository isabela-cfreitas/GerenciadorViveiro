using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
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

    public BalancoViewModel(VendasViewModel vendasVM, CustosViewModel custosVM)
    {
        _vendasVM = vendasVM;
        _custosVM = custosVM;

        Directory.CreateDirectory(PastaBalancos);
        CarregarBalancos();

        //atualizar quando mudar vendas ou custos
        _vendasVM.VendasAlteradas += (s, e) => AtualizarTodosOsMeses();
        _custosVM.CustosAlterados += (s, e) => AtualizarTodosOsMeses();
    }

    private string CaminhoArquivo =>
        Path.Combine(PastaBalancos, $"balanco_{AnoSelecionado}.xlsx");

    partial void OnAnoSelecionadoChanged(int value) => CarregarBalancos();

    private void AtualizarTodosOsMeses()
    {
        foreach (var item in Balancos)
        {
            AtualizarMes(item);
        }
    }

    private void AtualizarMes(ItemBalanco item)
    {
        item.Ano = AnoSelecionado;
        item.RendaBruta = _vendasVM.ObterTotalVendasPorPeriodo(AnoSelecionado, item.Mes);
        
        //carregar arquivo de custo esepecífico do mês
        var custosMes = ObterCustosMes(item.Mes);
        item.CustoTotal = custosMes;
        item.MargemLucro = item.RendaBruta - item.CustoTotal;
    }

    private decimal ObterCustosMes(int mes)
    {
        var arquivoCustos = Path.Combine("C:/Users/isaca/Dropbox/Custos", $"custos_{AnoSelecionado}_{mes:00}.xlsx");
        
        if (!File.Exists(arquivoCustos))
            return 0m;

        try
        {
            using var workbook = new XLWorkbook(arquivoCustos);
            var ws = workbook.Worksheet(1);
            
            decimal total = 0m;
            foreach (var row in ws.RowsUsed().Skip(1))
            {
                var valorCell = row.Cell(4);
                if (!valorCell.IsEmpty())
                {
                    if (decimal.TryParse(valorCell.GetString(), out decimal valor))
                        total += valor;
                    else
                        total += valorCell.GetValue<decimal>();
                }
            }
            return total;
        }
        catch
        {
            return 0m;
        }
    }

    public void CarregarBalancos()
    {
        Balancos.Clear();

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

                var item = new ItemBalanco
                {
                    Mes = row.Cell(1).GetValue<int>(),
                    Ano = AnoSelecionado,
                    PercentualLuis = GetDecimalSafe(6),
                    PercentualPedro = GetDecimalSafe(7),
                    PercentualViveiro = GetDecimalSafe(8)
                };

                AtualizarMes(item);
                Balancos.Add(item);
            }
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

        // Cabeçalhos
        ws.Cell(1, 1).Value = "Mês";
        ws.Cell(1, 2).Value = "Nome";
        ws.Cell(1, 3).Value = "Renda Bruta";
        ws.Cell(1, 4).Value = "Custo Total";
        ws.Cell(1, 5).Value = "Margem Lucro";
        ws.Cell(1, 6).Value = "% Luís";
        ws.Cell(1, 7).Value = "% Pedro";
        ws.Cell(1, 8).Value = "% Viveiro";
        ws.Cell(1, 9).Value = "R$ Luís";
        ws.Cell(1, 10).Value = "R$ Pedro";
        ws.Cell(1, 11).Value = "R$ Viveiro";

        ws.Range(1, 1, 1, 11).Style.Font.Bold = true;
        ws.Range(1, 1, 1, 11).Style.Fill.BackgroundColor = XLColor.LightGray;

        int linha = 2;
        foreach (var b in Balancos)
        {
            ws.Cell(linha, 1).Value = b.Mes;
            ws.Cell(linha, 2).Value = b.NomeMes;
            ws.Cell(linha, 3).Value = b.RendaBruta;
            ws.Cell(linha, 4).Value = b.CustoTotal;
            ws.Cell(linha, 5).Value = b.MargemLucro;
            ws.Cell(linha, 6).Value = b.PercentualLuis;
            ws.Cell(linha, 7).Value = b.PercentualPedro;
            ws.Cell(linha, 8).Value = b.PercentualViveiro;
            ws.Cell(linha, 9).Value = b.ValorLuis;
            ws.Cell(linha, 10).Value = b.ValorPedro;
            ws.Cell(linha, 11).Value = b.ValorViveiro;
            linha++;
        }

        // Linha de totais
        ws.Cell(linha, 2).Value = "TOTAL ANO";
        ws.Cell(linha, 3).Value = Balancos.Sum(b => b.RendaBruta);
        ws.Cell(linha, 4).Value = Balancos.Sum(b => b.CustoTotal);
        ws.Cell(linha, 5).Value = Balancos.Sum(b => b.MargemLucro);
        ws.Cell(linha, 9).Value = Balancos.Sum(b => b.ValorLuis);
        ws.Cell(linha, 10).Value = Balancos.Sum(b => b.ValorPedro);
        ws.Cell(linha, 11).Value = Balancos.Sum(b => b.ValorViveiro);
        
        ws.Range(linha, 1, linha, 11).Style.Font.Bold = true;
        ws.Range(linha, 1, linha, 11).Style.Fill.BackgroundColor = XLColor.LightBlue;

        ws.Columns().AdjustToContents();
        workbook.SaveAs(CaminhoArquivo);
    }

    [RelayCommand]
    public void NovaLinha()
    {
        // Não faz sentido adicionar linha - são sempre 12 meses fixos
        // Mas mantém o método para implementar a interface
    }

    [RelayCommand]
    private void ExcluirLinhas()
    {
        // Não permite excluir linhas - são sempre 12 meses fixos
        // Mas mantém o método para implementar a interface
    }

    private void CriarArquivo()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Balanço");

        // Cabeçalhos
        ws.Cell(1, 1).Value = "Mês";
        ws.Cell(1, 2).Value = "Nome";
        ws.Cell(1, 3).Value = "Renda Bruta";
        ws.Cell(1, 4).Value = "Custo Total";
        ws.Cell(1, 5).Value = "Margem Lucro";
        ws.Cell(1, 6).Value = "% Luís";
        ws.Cell(1, 7).Value = "% Pedro";
        ws.Cell(1, 8).Value = "% Viveiro";
        ws.Cell(1, 9).Value = "R$ Luís";
        ws.Cell(1, 10).Value = "R$ Pedro";
        ws.Cell(1, 11).Value = "R$ Viveiro";

        ws.Range(1, 1, 1, 11).Style.Font.Bold = true;
        ws.Range(1, 1, 1, 11).Style.Fill.BackgroundColor = XLColor.LightGray;

        // Cria 12 linhas (uma por mês) com valores padrão
        for (int mes = 1; mes <= 12; mes++)
        {
            ws.Cell(mes + 1, 1).Value = mes;
            ws.Cell(mes + 1, 2).Value = new DateTime(2000, mes, 1).ToString("MMMM", new CultureInfo("pt-BR"));
            ws.Cell(mes + 1, 6).Value = 40;
            ws.Cell(mes + 1, 7).Value = 40;
            ws.Cell(mes + 1, 8).Value = 20;
        }

        wb.SaveAs(CaminhoArquivo);

        // Carrega os 12 meses
        for (int mes = 1; mes <= 12; mes++)
        {
            var item = new ItemBalanco
            {
                Mes = mes,
                Ano = AnoSelecionado,
                PercentualLuis = 40m,
                PercentualPedro = 40m,
                PercentualViveiro = 20m
            };
            AtualizarMes(item);
            Balancos.Add(item);
        }
    }

    private ItemBalanco ClonarBalanco(ItemBalanco b)
    {
        return new ItemBalanco
        {
            Mes = b.Mes,
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
        if (_clipboard.Count == 0 || index < 0 || index >= Balancos.Count)
            return;

        // Ao colar, apenas copia os percentuais para o mês selecionado
        var destino = Balancos[index];
        var origem = _clipboard[0];

        destino.PercentualLuis = origem.PercentualLuis;
        destino.PercentualPedro = origem.PercentualPedro;
        destino.PercentualViveiro = origem.PercentualViveiro;
    }

    public void RecortarSelecionados()
    {
        CopiarSelecionados();
        // Não remove linhas, apenas limpa percentuais
        foreach (var balanco in BalancosSelecionados)
        {
            balanco.PercentualLuis = 0;
            balanco.PercentualPedro = 0;
            balanco.PercentualViveiro = 0;
        }
    }

    public void ApagarSelecionados()
    {
        // Reseta percentuais para padrão
        foreach (var balanco in BalancosSelecionados)
        {
            balanco.PercentualLuis = 40m;
            balanco.PercentualPedro = 40m;
            balanco.PercentualViveiro = 20m;
        }
    }

    public void Salvar() => SalvarBalancos();

    // Classe interna
    public partial class ItemBalanco : ObservableObject
    {
        [ObservableProperty]
        private int mes;

        [ObservableProperty]
        private int ano;

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

        public string NomeMes => new DateTime(Ano, Mes, 1).ToString("MMMM/yyyy", new CultureInfo("pt-BR"));

        public decimal ValorLuis => MargemLucro * (PercentualLuis / 100);
        public decimal ValorPedro => MargemLucro * (PercentualPedro / 100);
        public decimal ValorViveiro => MargemLucro * (PercentualViveiro / 100);

        partial void OnMesChanged(int value)
        {
            OnPropertyChanged(nameof(NomeMes));
        }

        partial void OnAnoChanged(int value) // ADICIONE ESTE MÉTODO
        {
            OnPropertyChanged(nameof(NomeMes));
        }

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