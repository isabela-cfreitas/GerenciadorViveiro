using System;
using System.IO;
using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;

namespace GerenciadorViveiro.ViewModels;

public partial class ConfiguracoesViewModel : ObservableObject
{
    private const string ArquivoConfig = "config.json";

    [ObservableProperty]
    private string pastaBase = string.Empty;

    //observer para notificar quando caminho muda
    public event EventHandler? PastasAlteradas;

    public ConfiguracoesViewModel(){
        CarregarConfiguracoes();
    }

    partial void OnPastaBaseChanged(string value){
        SalvarConfiguracoes();
        CriarPastas();
        OnPropertyChanged(nameof(PastaConfigurada));
        PastasAlteradas?.Invoke(this, EventArgs.Empty);
    }

    public bool PastaConfigurada => !string.IsNullOrWhiteSpace(PastaBase);

    public string CaminhoArquivoVendas => Path.Combine(PastaBase, "vendas.xlsx");
    public string PastaCustos => Path.Combine(PastaBase, "Custos");
    public string PastaBalancos => Path.Combine(PastaBase, "Balancos");
    public string PastaFrequencias => Path.Combine(PastaBase, "Frequencias");

    public void CriarPastas(){
        if (!PastaConfigurada)
            return;

        try{
            //isso evita sobrescrever arquivo se ja existir e poupa trabalho de ficar indo no dropbox buscar versao antiga
            //ou até de perder o arquivo se nao usar nuvem
            if (!string.IsNullOrWhiteSpace(PastaBase))
                Directory.CreateDirectory(PastaBase);

            Directory.CreateDirectory(PastaCustos);
            Directory.CreateDirectory(PastaBalancos);
            Directory.CreateDirectory(PastaFrequencias);
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao criar pastas: {ex.Message}");
        }
    }

    private void CarregarConfiguracoes(){
        try{
            if (!File.Exists(ArquivoConfig))
                return;

            var json = File.ReadAllText(ArquivoConfig);
            var config = JsonSerializer.Deserialize<ConfigData>(json);

            if (config != null && !string.IsNullOrWhiteSpace(config.PastaBase)){
                PastaBase = config.PastaBase;
            }
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao carregar configurações: {ex.Message}");
        }
    }

    private void SalvarConfiguracoes(){
        try{
            var config = new ConfigData
            {
                PastaBase = PastaBase
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(config, options);
            File.WriteAllText(ArquivoConfig, json);
        }
        catch (Exception ex){
            Console.WriteLine($"Erro ao salvar configurações: {ex.Message}");
        }
    }

    private class ConfigData{
        public string? PastaBase { get; set; }
    }
}