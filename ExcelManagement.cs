using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Excel
{
    public class ExcelManagement
    {
        private Microsoft.Office.Interop.Excel.Application? _app;

        private Workbook? _workbook = null;

        private Worksheet? _worksheet = null;

        private Range? _range = null;

        private string _path = string.Empty;

        public Workbook? Workbook { get => _workbook; set => _workbook = value; }

        public Worksheet? Worksheet { get => _worksheet; set => _worksheet = value; }

        public Range? Range { get => _range; set => _range = value; }

        public string Path { get => _path; set => _path = value; }

        public ExcelManagement(string path)
        {
            _path = path;
        }

        public void CreateNewFile(string sheetName = "InitialSheet")
        {
            try
            {
                DeleteIfExistsFile(_path);

                Create(_path, sheetName, 1);

                SaveAs();

                Close();
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        public void Create(string path, string sheetName = "InitialSheet", int sheet = 1)
        {
            try
            {
                _path = path;

                _app = new Microsoft.Office.Interop.Excel.Application();

                if (_app == null) return;

                _workbook = _app?.Workbooks.Add(System.Type.Missing);

                _worksheet = _app?.Worksheets[sheet];

                if (_worksheet == null) return;

                _worksheet.Name = sheetName;

                _range = _worksheet.UsedRange;
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        public void Open(string path, int sheet = 1, bool visible = false)
        {
            try
            {
                _path = path;

                _app = new Microsoft.Office.Interop.Excel.Application();

                if (_app == null) return;

                _app.Visible = visible;

                _workbook = _app?.Workbooks.Open(path);

                _worksheet = _app?.Worksheets[sheet];

                if (_worksheet == null) return;

                _range = _worksheet.UsedRange;
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        public void Close()
        {
            try
            {
                _workbook?.Close();

                _app?.Quit();

                Helper.releaseObject(_worksheet);

                Helper.releaseObject(_workbook);

                Helper.releaseObject(_app);

                Helper.releaseObject(_range);

                System.GC.Collect();

                System.GC.WaitForPendingFinalizers();

                System.GC.Collect();
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        public void SaveAs()
        {
            if (_workbook != null)
            {
                _workbook.SaveAs(_path);
            }
        }

        public void SaveCopyAs()
        {
            if (_workbook != null)
            {
                _workbook.SaveCopyAs(_path);
            }
        }

        public void Save()
        {
            if (_workbook != null)
            {
                _workbook.Save();
            }
        }

        public void WriteToCell(int row, int col, object value)
        {
            if (_worksheet != null)
            {
                _worksheet.Cells[row, col] = value;
            }
        }

        public static void DeleteIfExistsFile(string fileName)
        {
            try
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
            catch (System.Exception)
            {
                throw;
            }
        }
    }
}
