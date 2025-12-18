using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using GerenciadorViveiro.ViewModels.Interfaces;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using Avalonia.Controls;

namespace GerenciadorViveiro.ViewModels;

public partial class BalancoViewModel : ObservableObject, IEditableGridViewModel
{
    private readonly VendasViewModel _vendasVM;
    private readonly CustosViewModel _custosVM;
    private readonly ConfiguracoesViewModel _configuracoesVM;

    private List<ItemBalanco> _clipboard = new();

    [ObservableProperty]
    private ObservableCollection<ItemBalanco> balancosSelecionados = new();

    [ObservableProperty]
    private ObservableCollection<ItemBalanco> balancos = new();

    [ObservableProperty]
    private int anoSelecionado = DateTime.Today.Year;

    [ObservableProperty]
    private string? mensagemErro;

    public BalancoViewModel(VendasViewModel vendasVM, CustosViewModel custosVM, ConfiguracoesViewModel configuracoesVM){
        _vendasVM = vendasVM;
        _custosVM = custosVM;
        _configuracoesVM = configuracoesVM;

        if (_configuracoesVM.PastaConfigurada){
            _configuracoesVM.CriarPastas();
            InicializarBalancos();
        }
        
        //se a pasta base ou a pasta de balanco mudarem
        _configuracoesVM.PastasAlteradas += (s, e) =>{
            if (_configuracoesVM.PastaConfigurada)
                InicializarBalancos();
        };

        //atualizar quando mudar vendas ou custos
        _vendasVM.VendasAlteradas += (s, e) => AtualizarTodosOsMeses();
        _custosVM.CustosAlterados += (s, e) => AtualizarTodosOsMeses();
    }

    partial void OnAnoSelecionadoChanged(int value){
        if (!ValidarAno(value)){
            MensagemErro = "Ano inválido. Digite um ano entre 1900 e 2100.";
            return;
        }
        MensagemErro = null;
        InicializarBalancos();
    }

    private bool ValidarAno(int ano){
        return ano >= 1900 && ano <= 2100;
    }

    private void InicializarBalancos(){
        Balancos.Clear();

        if (!ValidarAno(AnoSelecionado))
            return;

        //sempre cria os 12 meses do ano
        for (int mes = 1; mes <= 12; mes++){
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

    private void AtualizarTodosOsMeses(){
        foreach (var item in Balancos){
            AtualizarMes(item);
        }
    }

    private void AtualizarMes(ItemBalanco item){
        item.Ano = AnoSelecionado;
        item.RendaBruta = _vendasVM.ObterTotalVendasPorPeriodo(AnoSelecionado, item.Mes);
        
        //carregar arquivo de custo específico do mês
        var custosMes = ObterCustosMes(item.Mes);
        item.CustoTotal = custosMes;
        item.MargemLucro = item.RendaBruta - item.CustoTotal;
    }

    private decimal ObterCustosMes(int mes){
        var arquivoCustos = Path.Combine(_configuracoesVM.PastaCustos, $"custos_{AnoSelecionado}_{mes:00}.xlsx");
        
        if (!File.Exists(arquivoCustos))
            return 0m;

        try{
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
        catch{
            return 0m;
        }
    }

    [RelayCommand]
    private async Task ExportarBalanco(){
        if (!_configuracoesVM.PastaConfigurada){
            await MostrarMensagem("Erro", "Configure a pasta base nas Configurações primeiro.");
            return;
        }

        try{
            _configuracoesVM.CriarPastas();

            var caminhoArquivo = Path.Combine(_configuracoesVM.PastaBalancos, $"balanco_{AnoSelecionado}.xlsx");

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
            foreach (var b in Balancos){
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
            workbook.SaveAs(caminhoArquivo);

            await MostrarMensagem("Sucesso", $"Balanço exportado com sucesso!\n{caminhoArquivo}");
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao exportar balanço: {ex.Message}");
            await MostrarMensagem("Erro", $"Erro ao exportar: {ex.Message}");
        }
    }

    [RelayCommand]
    public void NovaLinha(){
        //apesar de esse método não fazer sentido para balanco tem que declarar ele aqui por causa da interface
        //erro meu por nao ter pensado nisso quando fiz a interface :)
    }

    [RelayCommand]
    private void ExcluirLinhas(){
        //esse aqui tbm
    }

    private ItemBalanco ClonarBalanco(ItemBalanco b){
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

        //percentuais tem valor padrao, mas da p alterar
        var destino = Balancos[index];
        var origem = _clipboard[0];

        destino.PercentualLuis = origem.PercentualLuis;
        destino.PercentualPedro = origem.PercentualPedro;
        destino.PercentualViveiro = origem.PercentualViveiro;
    }

    public void RecortarSelecionados(){
        CopiarSelecionados();
        // Não remove linhas, apenas limpa percentuais
        foreach (var balanco in BalancosSelecionados){
            balanco.PercentualLuis = 0;
            balanco.PercentualPedro = 0;
            balanco.PercentualViveiro = 0;
        }
    }

    public void ApagarSelecionados(){
        // Reseta percentuais para padrão
        foreach (var balanco in BalancosSelecionados){
            balanco.PercentualLuis = 40m;
            balanco.PercentualPedro = 40m;
            balanco.PercentualViveiro = 20m;
        }
    }

    public async Task SalvarAsync(){
        await ExportarBalanco();
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

    // Classe interna
    public partial class ItemBalanco : ObservableObject{
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

        partial void OnMesChanged(int value){
            OnPropertyChanged(nameof(NomeMes));
        }

        partial void OnAnoChanged(int value){
            OnPropertyChanged(nameof(NomeMes));
        }

        partial void OnMargemLucroChanged(decimal value)
        {
            OnPropertyChanged(nameof(ValorLuis));
            OnPropertyChanged(nameof(ValorPedro));
            OnPropertyChanged(nameof(ValorViveiro));
        }

        partial void OnPercentualLuisChanged(decimal value){
            OnPropertyChanged(nameof(ValorLuis));
        }

        partial void OnPercentualPedroChanged(decimal value){
            OnPropertyChanged(nameof(ValorPedro));
        }

        partial void OnPercentualViveiroChanged(decimal value){
            OnPropertyChanged(nameof(ValorViveiro));
        }
    }
}