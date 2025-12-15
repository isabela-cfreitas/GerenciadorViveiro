using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using GerenciadorViveiro.ViewModels;
using System.Linq;

namespace GerenciadorViveiro.Views;

public partial class MainWindow : Window {
    private MainWindowViewModel _viewModel;

    public MainWindow() {
        InitializeComponent();
        _viewModel = new MainWindowViewModel();
        DataContext = _viewModel;

        // Captura a seleção múltipla
        VendasDataGrid.SelectionChanged += (s, e) => {
            if (_viewModel != null) {
                _viewModel.VendasSelecionadas.Clear();
                foreach (var item in VendasDataGrid.SelectedItems.Cast<Models.Venda>()) {
                    _viewModel.VendasSelecionadas.Add(item);
                }
            }
        };

        VendasDataGrid.AddHandler(KeyDownEvent, VendasDataGrid_KeyDownTunnel, RoutingStrategies.Tunnel);
    }

    private void OnTabSelectionChanged(object? sender, SelectionChangedEventArgs e){
        if (!IsInitialized)
            return;

        //segurança extra
        if (VendasDataGrid == null || _viewModel == null)
            return;

        VendasDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        VendasDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        //atualiza frequências ao trocar de aba
        _viewModel.CalcularFrequencias();
    }
    private void CalcularFrequencias(object? sender, TappedEventArgs e) {
        _viewModel.CalcularFrequencias();
    }

    private void VendasDataGrid_KeyDownTunnel(object? sender, KeyEventArgs e){
        if (e.Key != Key.Enter)
            return;

        // BLOQUEIA o comportamento padrão ANTES do DataGrid processar
        e.Handled = true;

        // Salva a célula atual
        VendasDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        VendasDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        int rowIndex = VendasDataGrid.SelectedIndex;
        var column = VendasDataGrid.CurrentColumn;

        if (rowIndex < 0 || column == null)
            return;

        int colIndex = column.DisplayIndex;
        int totalColumns = VendasDataGrid.Columns.Count;

        // Última coluna → próxima linha
        if (colIndex == totalColumns - 1)
        {
            if (rowIndex < VendasDataGrid.ItemsSource.Cast<object>().Count() - 1)
            {
                VendasDataGrid.SelectedIndex = rowIndex + 1;
                VendasDataGrid.CurrentColumn = VendasDataGrid.Columns[0];
            }
        }
        else
        {
            // Próxima coluna
            VendasDataGrid.CurrentColumn = VendasDataGrid.Columns[colIndex + 1];
        }

        // Força o foco na nova célula
        VendasDataGrid.Focus();
        
        // Pequeno delay para garantir que a célula seja selecionada antes de entrar em edição
        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
            VendasDataGrid.BeginEdit();
        }, Avalonia.Threading.DispatcherPriority.Background);
    }
}