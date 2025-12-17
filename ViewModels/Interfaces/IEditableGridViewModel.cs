namespace GerenciadorViveiro.ViewModels.Interfaces;
public interface IEditableGridViewModel
{
    void ApagarSelecionados();
    void CopiarSelecionados();
    void RecortarSelecionados();
    void Colar(int index);
    void NovaLinha();
    void Salvar();
}
