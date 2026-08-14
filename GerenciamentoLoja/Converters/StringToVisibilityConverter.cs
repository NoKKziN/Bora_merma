using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace GerenciamentoLoja.Converters;

// Oculta o elemento quando a string associada estiver vazia (ex: mensagens de erro/aviso).
public class StringToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object parameter, CultureInfo culture)
    {
        return string.IsNullOrWhiteSpace(value as string) ? Visibility.Collapsed : Visibility.Visible;
    }

    public object ConvertBack(object? value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
