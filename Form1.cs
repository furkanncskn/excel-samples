using System;
using System.IO;
using System.Windows.Forms;

namespace Excel
{
    public partial class Form1 : Form
    {
        private static ExcelManagement? _excel;

        public Form1()
        {
            InitializeComponent();
        }

        private void BtnCreateExcel_Click(object sender, System.EventArgs e)
        {
            _excel!.Path = Helper.CreateFile("Create Excel File");

            _excel?.CreateNewFile();
        }

        private void BtnWriteToExcel_Click(object sender, EventArgs e)
        {
            _excel?.Open(_excel.Path, sheet: 1, visible: true);

            _excel?.WriteToCell(1, 1, Convert.ToInt32(txtCellValue.Text, Helper.GetCustomNumberFormatInfo()));

            _excel?.SaveAs();

            _excel?.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (_excel == null)
            {
                _excel = new ExcelManagement(Directory.GetCurrentDirectory());
            }
        }
    }
}
