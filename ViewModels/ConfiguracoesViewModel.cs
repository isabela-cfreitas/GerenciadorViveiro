using System;
using System.IO;
using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace GerenciadorViveiro.ViewModels;

/// <summary>
/// ViewModel para configurações do aplicativo (caminhos de arquivos, etc)
/// </summary>
public partial class ConfiguracoesViewModel : ObservableObject
{
    private const string ArquivoConfig = "config.json";

    [ObservableProperty]
    private string pastaVendas = "C:/Users/isaca/Dropbox";

    [ObservableProperty]
    private string pastaCustos = "C:/Users/isaca/Dropbox/Custos";

    // Eventos para notificar quando os caminhos mudarem
    public event EventHandler? PastasAlteradas;

    public ConfiguracoesViewModel()
    {
        CarregarConfiguracoes();
        CriarPastas();
    }

    partial void OnPastaVendasChanged(string value)
    {
        SalvarConfiguracoes();
        PastasAlteradas?.Invoke(this, EventArgs.Empty);
    }

    partial void OnPastaCustosChanged(string value)
    {
        SalvarConfiguracoes();
        PastasAlteradas?.Invoke(this, EventArgs.Empty);
    }

    public string CaminhoArquivoVendas => Path.Combine(PastaVendas, "vendas.xlsx");

    public void CriarPastas()
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(PastaVendas))
                Directory.CreateDirectory(Path.GetDirectoryName(CaminhoArquivoVendas) ?? PastaVendas);

            if (!string.IsNullOrWhiteSpace(PastaCustos))
                Directory.CreateDirectory(PastaCustos);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao criar pastas: {ex.Message}");
        }
    }

    private void CarregarConfiguracoes()
    {
        try
        {
            if (!File.Exists(ArquivoConfig))
                return;

            var json = File.ReadAllText(ArquivoConfig);
            var config = JsonSerializer.Deserialize<ConfigData>(json);

            if (config != null)
            {
                PastaVendas = config.PastaVendas ?? PastaVendas;
                PastaCustos = config.PastaCustos ?? PastaCustos;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao carregar configurações: {ex.Message}");
        }
    }

    private void SalvarConfiguracoes()
    {
        try
        {
            var config = new ConfigData
            {
                PastaVendas = PastaVendas,
                PastaCustos = PastaCustos
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(config, options);
            File.WriteAllText(ArquivoConfig, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao salvar configurações: {ex.Message}");
        }
    }

    private class ConfigData
    {
        public string? PastaVendas { get; set; }
        public string? PastaCustos { get; set; }
    }
}