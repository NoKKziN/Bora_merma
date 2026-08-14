using GerenciamentoLoja.Models;

namespace GerenciamentoLoja.Services;

public interface IConfigService
{
    AppConfig Carregar();
    void Salvar(AppConfig config);
}
