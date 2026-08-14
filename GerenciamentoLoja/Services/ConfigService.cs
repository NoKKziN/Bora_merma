using System.IO;
using System.Text.Json;
using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

public class ConfigService : IConfigService
{
    private readonly string _caminhoConfig;

    public ConfigService()
    {
        var pasta = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GerenciamentoLoja");
        Directory.CreateDirectory(pasta);
        _caminhoConfig = Path.Combine(pasta, "config.json");
    }

    public AppConfig Carregar()
    {
        if (!File.Exists(_caminhoConfig))
        {
            return new AppConfig();
        }

        try
        {
            var json = File.ReadAllText(_caminhoConfig);
            return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return new AppConfig();
        }
    }

    public void Salvar(AppConfig config)
    {
        var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_caminhoConfig, json);
    }
}
