using Avalonia.Controls;
using GerenciadorViveiro.ViewModels;

namespace GerenciadorViveiro.Views;
public partial class MainWindow : Window {
    public MainWindow() {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
    }
}