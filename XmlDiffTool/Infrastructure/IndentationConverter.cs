using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace XmlDiffTool.Infrastructure
{
    public class IndentationConverter : IValueConverter
    {
        private const double IndentSize = 16.0;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int level)
            {
                return new Thickness(level * IndentSize, 0, 0, 0);
            }

            return new Thickness(0);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
