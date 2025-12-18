using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using GerenciadorViveiro.ViewModels;
using GerenciadorViveiro.ViewModels.Interfaces;
using System.Linq;
using System.Threading.Tasks;

namespace GerenciadorViveiro.Views;

public partial class MainWindow : Window
{
    private MainWindowViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainWindowViewModel();
        DataContext = _viewModel;

        ConfigurarEventosVendas();
        ConfigurarEventosCustos();
        ConfigurarEventosBalanco();
        
        //nao deixa acessar outras abas se não configurado caminho da pasta base
        MainTabControl.SelectionChanged += ValidarConfiguracaoAntesNavegar;
    }

    private void ValidarConfiguracaoAntesNavegar(object? sender, SelectionChangedEventArgs e)
    {
        if (!IsInitialized || _viewModel == null)
            return;

        //se não está configurado e tentou sair da aba de configurações
        if (!_viewModel.ConfiguracoesVM.PastaConfigurada && 
            MainTabControl.SelectedIndex != 0)
        {
            //força voltar para aba de configurações
            MainTabControl.SelectedIndex = 0;
        }
    }

    private void ConfigurarEventosVendas()
    {
        VendasDataGrid.SelectionChanged += (s, e) =>
        {
            if (_viewModel?.VendasVM != null)
            {
                _viewModel.VendasVM.VendasSelecionadas.Clear();
                foreach (var item in VendasDataGrid.SelectedItems.Cast<Models.Venda>())
                {
                    _viewModel.VendasVM.VendasSelecionadas.Add(item);
                }
            }
        };
        VendasDataGrid.AddHandler(KeyDownEvent, DataGrid_KeyDownTunnel, RoutingStrategies.Tunnel);
    }

    private void ConfigurarEventosCustos()
    {
        CustosDataGrid.SelectionChanged += (s, e) =>
        {
            if (_viewModel?.CustosVM != null)
            {
                _viewModel.CustosVM.CustosSelecionados.Clear();
                foreach (var item in CustosDataGrid.SelectedItems.Cast<Models.Custo>())
                {
                    _viewModel.CustosVM.CustosSelecionados.Add(item);
                }
            }
        };
        CustosDataGrid.AddHandler(KeyDownEvent, DataGrid_KeyDownTunnel, RoutingStrategies.Tunnel);
    }

    private async void SelecionarPastaBase(object? sender, RoutedEventArgs e)
    {
        var folder = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Selecione a pasta base onde os dados serão salvos",
            AllowMultiple = false
        });

        if (folder.Count > 0)
        {
            _viewModel.ConfiguracoesVM.PastaBase = folder[0].Path.LocalPath;
        }
    }

    private void OnTabSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (!IsInitialized || _viewModel == null)
            return;

        if (VendasDataGrid != null)
        {
            VendasDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            VendasDataGrid.CommitEdit(DataGridEditingUnit.Row, true);
        }

        if (CustosDataGrid != null)
        {
            CustosDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            CustosDataGrid.CommitEdit(DataGridEditingUnit.Row, true);
        }

        _viewModel.FrequenciasVM.CalcularFrequencias();
    }

    private void CalcularFrequencias(object? sender, TappedEventArgs e)
    {
        _viewModel.FrequenciasVM.CalcularFrequencias();
    }

    private void ConfigurarEventosBalanco()
    {
        BalancoDataGrid.SelectionChanged += (s, e) =>
        {
            if (_viewModel?.BalancoVM != null)
            {
                _viewModel.BalancoVM.BalancosSelecionados.Clear();
                foreach (var item in BalancoDataGrid.SelectedItems.Cast<BalancoViewModel.ItemBalanco>())
                {
                    _viewModel.BalancoVM.BalancosSelecionados.Add(item);
                }
            }
        };
        BalancoDataGrid.AddHandler(KeyDownEvent, DataGrid_KeyDownTunnel, RoutingStrategies.Tunnel);
    }

    //private void DataGrid_KeyDownTunnel(object? sender, KeyEventArgs e)
    private async void DataGrid_KeyDownTunnel(object? sender, KeyEventArgs e)
    {
        if (sender is not DataGrid dataGrid)
            return;

        //identifica qual ViewModel usar baseado no DataGrid
        IEditableGridViewModel? viewModel = dataGrid.Name switch
        {
            nameof(VendasDataGrid) => _viewModel.VendasVM,
            nameof(CustosDataGrid) => _viewModel.CustosVM,
            nameof(BalancoDataGrid) => _viewModel.BalancoVM,
            _ => null
        };

        if (viewModel == null)
            return;

        // DEL - Apagar
        if (e.Key == Key.Delete)
        {
            viewModel.ApagarSelecionados();
            e.Handled = true;
            return;
        }

        // CTRL + C - Copiar
        if (e.Key == Key.C && e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            viewModel.CopiarSelecionados();
            e.Handled = true;
            return;
        }

        // CTRL + N - Nova linha
        if (e.Key == Key.N && e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            viewModel.NovaLinha();
            e.Handled = true;
            return;
        }

        // CTRL + S - Salvar
        if (e.Key == Key.S && e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            //viewModel.Salvar();
            await viewModel.SalvarAsync();
            e.Handled = true;
            return;
        }

        // CTRL + V - Colar
        if (e.Key == Key.V && e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            int index = dataGrid.SelectedIndex;
            if (index >= 0)
                viewModel.Colar(index);

            e.Handled = true;
            return;
        }

        // CTRL + X - Recortar
        if (e.Key == Key.X && e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            viewModel.RecortarSelecionados();
            e.Handled = true;
            return;
        }

        // ENTER - Navega entre células
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

        if (colIndex == totalColumns - 1)
        {
            if (rowIndex < dataGrid.ItemsSource.Cast<object>().Count() - 1)
            {
                dataGrid.SelectedIndex = rowIndex + 1;
                dataGrid.CurrentColumn = dataGrid.Columns[0];
            }
        }
        else
        {
            dataGrid.CurrentColumn = dataGrid.Columns[colIndex + 1];
        }

        dataGrid.Focus();
        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => dataGrid.BeginEdit(),
            Avalonia.Threading.DispatcherPriority.Background
        );
    }
}