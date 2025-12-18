using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    //todas as tabelas sendo criadas aqui
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
        //tem q inicializar configurações primeiro para poder pegar o caminho da pasta base que vai armazenar tudo
        ConfiguracoesVM = new ConfiguracoesViewModel();

        //inicializa na ordem correta (vendas primeiro, é que os outros dependem dele)
        //todos recebem um objeto do menu de configurações por conta do caminho da pasta base
        VendasVM = new VendasViewModel(ConfiguracoesVM);
        FrequenciasVM = new FrequenciasViewModel(VendasVM, ConfiguracoesVM);
        CustosVM = new CustosViewModel(ConfiguracoesVM);
        BalancoVM = new BalancoViewModel(VendasVM, CustosVM, ConfiguracoesVM);
    }
}