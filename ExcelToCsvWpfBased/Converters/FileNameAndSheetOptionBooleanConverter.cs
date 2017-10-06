using System;
using System.Globalization;
using System.Windows.Data;

namespace ExcelToCsvWpfBased
{
    internal class FileNameAndSheetOptionBooleanConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var fileName = (string)values[0];
            var isAllSheetSelected = (bool)values[1];
            var isSingleSheetSelected = (bool)values[2];
            var selectedSheet = (string)values[3];

            if (string.IsNullOrEmpty(fileName))
            {
                return false;
            }

            if (isAllSheetSelected)
            {
                return true;
            }

            return (isSingleSheetSelected && !string.IsNullOrEmpty(selectedSheet));
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException(string.Format("{0} is one way converter", typeof(FileNameAndSheetOptionBooleanConverter).FullName));
        }
    }
}
