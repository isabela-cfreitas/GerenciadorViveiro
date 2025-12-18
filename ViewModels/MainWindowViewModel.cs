using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.ViewModels;

/// <summary>
/// ViewModel principal que orquestra as outras ViewModels.
/// Não possui lógica de negócio, apenas coordena as telas.
/// </summary>
public partial class MainWindowViewModel : ObservableObject
{
    // ViewModels das diferentes telas/funcionalidades
    [ObservableProperty]
    private ConfiguracoesViewModel configuracoesVM;

    [ObservableProperty]
    private VendasViewModel vendasVM;

    [ObservableProperty]
    private FrequenciasViewModel frequenciasVM;

    [ObservableProperty]
    private CustosViewModel custosVM;

    [ObservableProperty]
    private BalancoViewModel balancoVM;

    public MainWindowViewModel()
    {
        // Inicializa configurações primeiro
        ConfiguracoesVM = new ConfiguracoesViewModel();

        // Inicializa na ordem correta (vendas primeiro, pois outros dependem dele)
        VendasVM = new VendasViewModel();
        FrequenciasVM = new FrequenciasViewModel(VendasVM);
        CustosVM = new CustosViewModel();
        BalancoVM = new BalancoViewModel(VendasVM, CustosVM);
    }
}