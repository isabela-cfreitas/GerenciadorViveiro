using Avalonia.Controls;
using Avalonia.Input;
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

}