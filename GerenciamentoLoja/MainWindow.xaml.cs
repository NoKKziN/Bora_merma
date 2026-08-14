using System.Windows;
using GerenciamentoLoja.ViewModels;

namespace GerenciamentoLoja
{
    public partial class MainWindow : Window
    {
        public MainWindow(ShellViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
