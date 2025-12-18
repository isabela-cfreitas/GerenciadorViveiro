using System.Threading.Tasks;

namespace GerenciadorViveiro.ViewModels.Interfaces;
public interface IEditableGridViewModel
{
    void ApagarSelecionados();
    void CopiarSelecionados();
    void RecortarSelecionados();
    void Colar(int index);
    void NovaLinha();
    Task SalvarAsync();
}
