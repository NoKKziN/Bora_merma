using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Markup;
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

            ConfigurarCulturaPtBr();

            var configService = new ConfigService();
            var excelService = new ExcelInteropService();
            _excelService = excelService;
            var dadosService = new LojaDataService(excelService);

            var shellViewModel = new ShellViewModel(
                dadosService,
                new ConexaoPlanilhaViewModel(dadosService, excelService, configService),
                new DashboardViewModel(dadosService),
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

        // Por padrão, bindings do WPF com StringFormat (ex: {0:C2}) usam
        // FrameworkElement.Language, que é "en-US" fixo, independente da cultura do
        // Windows ou da thread — por isso os valores apareciam em dólar. Isso força
        // pt-BR (moeda R$, vírgula decimal) em todo o app, já que ele é feito para uma
        // loja brasileira independentemente da configuração regional da máquina.
        private static void ConfigurarCulturaPtBr()
        {
            var cultura = CultureInfo.GetCultureInfo("pt-BR");

            Thread.CurrentThread.CurrentCulture = cultura;
            Thread.CurrentThread.CurrentUICulture = cultura;
            CultureInfo.DefaultThreadCurrentCulture = cultura;
            CultureInfo.DefaultThreadCurrentUICulture = cultura;

            FrameworkElement.LanguageProperty.OverrideMetadata(
                typeof(FrameworkElement),
                new FrameworkPropertyMetadata(XmlLanguage.GetLanguage(cultura.IetfLanguageTag)));
        }
    }
}
