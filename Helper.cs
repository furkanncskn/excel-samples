using System;
using System.Globalization;
using System.Windows.Forms;

namespace Excel
{
    public static class Helper
    {
        public static string CreateFile(string Title)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new())
                {
                    if (saveFileDialog == null) { throw new NullReferenceException(); }

                    saveFileDialog.Filter = "Excel 2010|*.xlsx|Excel|*.xls";

                    saveFileDialog.Title = Title;

                    if (saveFileDialog.ShowDialog() != DialogResult.OK) { return String.Empty; }

                    return saveFileDialog.FileName;
                };
            }
            catch (Exception ex)
            {
                Message.ShowMessage(ex.Message, Message.MessageType.Error);

                return String.Empty;
            }
        }

        public static string OpenFile()
        {
            using (OpenFileDialog openFileDialog = new())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }

            return string.Empty;
        }

        public static CultureInfo GetSpesificCultereInfo(
            string name,
            string decimalSeperator,
            string groupSeperator,
            string symbol
            )
        {
            try
            {
                CultureInfo cultureInfo = CultureInfo.GetCultureInfo(name);

                NumberFormatInfo numFormatInfo = (NumberFormatInfo)cultureInfo.NumberFormat.Clone();

                numFormatInfo.NumberDecimalSeparator = decimalSeperator;

                numFormatInfo.CurrencyGroupSeparator = groupSeperator;

                numFormatInfo.CurrencySymbol = symbol;

                return cultureInfo;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static NumberFormatInfo GetCustomNumberFormatInfo()
        {
            try
            {
                return GetSpesificCultereInfo(
                    name: "tr-TR",
                    decimalSeperator: ",",
                    groupSeperator: ".",
                    symbol: "U+20BA"
                ).NumberFormat;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void releaseObject(object? obj)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj!) > 0) ;

                obj = null;
            }
            catch (System.Exception)
            {
                obj = null;

                throw;
            }
            finally
            {
                System.GC.Collect();
            }
        }
    }
}
