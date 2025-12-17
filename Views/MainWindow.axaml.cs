using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using GerenciadorViveiro.ViewModels;
using GerenciadorViveiro.ViewModels.Interfaces;
using System.Linq;

namespace GerenciadorViveiro.Views;

public partial class MainWindow : Window {
    private MainWindowViewModel _viewModel;

    public MainWindow() {
        InitializeComponent();
        _viewModel = new MainWindowViewModel();
        DataContext = _viewModel;

        // Configura eventos para a tabela de Vendas
        VendasDataGrid.SelectionChanged += (s, e) => {
            if (_viewModel != null) {
                _viewModel.VendasSelecionadas.Clear();
                foreach (var item in VendasDataGrid.SelectedItems.Cast<Models.Venda>()) {
                    _viewModel.VendasSelecionadas.Add(item);
                }
            }
        };
        VendasDataGrid.AddHandler(KeyDownEvent, DataGrid_KeyDownTunnel, RoutingStrategies.Tunnel);

        // Configura eventos para a tabela de Custos
        CustosDataGrid.SelectionChanged += (s, e) => {
            if (_viewModel?.CustosVM != null) {
                _viewModel.CustosVM.CustosSelecionados.Clear();
                foreach (var item in CustosDataGrid.SelectedItems.Cast<Models.Custo>()) {
                    _viewModel.CustosVM.CustosSelecionados.Add(item);
                }
            }
        };
        CustosDataGrid.AddHandler(KeyDownEvent, DataGrid_KeyDownTunnel, RoutingStrategies.Tunnel);
    }

    private void OnTabSelectionChanged(object? sender, SelectionChangedEventArgs e){
        if (!IsInitialized)
            return;

        if (VendasDataGrid == null || _viewModel == null)
            return;

        VendasDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        VendasDataGrid.CommitEdit(DataGridEditingUnit.Row, true);
        
        CustosDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        CustosDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        _viewModel.CalcularFrequencias();
    }

    private void CalcularFrequencias(object? sender, TappedEventArgs e) {
        _viewModel.CalcularFrequencias();
    }

    private void DataGrid_KeyDownTunnel(object? sender, KeyEventArgs e) {
        if (sender is not DataGrid dataGrid)
            return;

        // Identifica qual ViewModel usar baseado no DataGrid
        IEditableGridViewModel? viewModel = dataGrid.Name switch {
            nameof(VendasDataGrid) => _viewModel,
            nameof(CustosDataGrid) => _viewModel.CustosVM,
            _ => null
        };

        if (viewModel == null)
            return;

        // DEL
        if (e.Key == Key.Delete) {
            viewModel.ApagarSelecionados();
            e.Handled = true;
            return;
        }

        // CTRL + C
        if (e.Key == Key.C && e.KeyModifiers.HasFlag(KeyModifiers.Control)) {
            viewModel.CopiarSelecionados();
            e.Handled = true;
            return;
        }

        // CTRL + N
        if (e.Key == Key.N && e.KeyModifiers.HasFlag(KeyModifiers.Control)) {
            viewModel.NovaLinha();
            e.Handled = true;
            return;
        }

        // CTRL + S
        if (e.Key == Key.S && e.KeyModifiers.HasFlag(KeyModifiers.Control)) {
            viewModel.Salvar();
            e.Handled = true;
            return;
        }

        // CTRL + V
        if (e.Key == Key.V && e.KeyModifiers.HasFlag(KeyModifiers.Control)) {
            int index = dataGrid.SelectedIndex;
            if (index >= 0)
                viewModel.Colar(index);

            e.Handled = true;
            return;
        }

        // CTRL + X
        if (e.Key == Key.X && e.KeyModifiers.HasFlag(KeyModifiers.Control)) {
            viewModel.RecortarSelecionados();
            e.Handled = true;
            return;
        }

        // ENTER - navega entre células
        if (e.Key != Key.Enter)
            return;

        e.Handled = true;

        dataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        dataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        int rowIndex = dataGrid.SelectedIndex;
        var column = dataGrid.CurrentColumn;

        if (rowIndex < 0 || column == null)
            return;

        int colIndex = column.DisplayIndex;
        int totalColumns = dataGrid.Columns.Count;

        if (colIndex == totalColumns - 1) {
            if (rowIndex < dataGrid.ItemsSource.Cast<object>().Count() - 1) {
                dataGrid.SelectedIndex = rowIndex + 1;
                dataGrid.CurrentColumn = dataGrid.Columns[0];
            }
        }
        else {
            dataGrid.CurrentColumn = dataGrid.Columns[colIndex + 1];
        }

        dataGrid.Focus();
        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => dataGrid.BeginEdit(),
            Avalonia.Threading.DispatcherPriority.Background
        );
    }
}