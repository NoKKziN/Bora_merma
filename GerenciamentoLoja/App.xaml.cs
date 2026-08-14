using System.Windows;
using GerenciamentoLoja.Services;
using GerenciamentoLoja.ViewModels;

namespace GerenciamentoLoja
{
    public partial class App : Application
    {
        private IExcelWorkbookService? _excelService;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var configService = new ConfigService();
            var excelService = new ExcelInteropService();
            _excelService = excelService;
            var dadosService = new LojaDataService(excelService);

            var shellViewModel = new ShellViewModel(
                dadosService,
                new ConexaoPlanilhaViewModel(dadosService, excelService, configService),
                new DashboardPlaceholderViewModel(),
                new ProdutosDisponiveisViewModel(dadosService),
                new VendasRealizadasViewModel(dadosService),
                new AdicionarVendaViewModel(dadosService),
                new BalancoGeralViewModel(dadosService));

            var mainWindow = new MainWindow(shellViewModel);
            mainWindow.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _excelService?.Dispose();
            base.OnExit(e);
        }
    }
}
