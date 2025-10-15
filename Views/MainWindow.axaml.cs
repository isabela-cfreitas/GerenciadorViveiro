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
    private void CalcularFrequencias(object? sender, TappedEventArgs e) {
        _viewModel.CalcularFrequencias();
    }

}