using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComputerConfigurator
{
    public partial class Form2 : Form
    {
        public Form2(Settings settings)
        {
            InitializeComponent();

            // Получить объект приложения Excel.
            Excel.Application excel_app = new Excel.Application();
            Excel.Workbook workbook = excel_app.Workbooks.Open(
                Path.GetFullPath("../../Resources/CPU"),
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

            

            
        }

        private void MotherBoardExtraction(MotherBoard Mom, int j)
        {

        }

        private Computer computer = new Computer();
        //private List<Computer> computers;

        private void example()
        {
            // Получить объект приложения Excel.
            Excel.Application excel_app = new Excel.Application();

            // Сделать Excel видимым (необязательно).
            //excel_app.Visible = true;

            // Откройте рабочую книгу только для чтения.
            Excel.Workbook workbook = excel_app.Workbooks.Open(
                Path.GetFullPath("../../Resources/CPU"),
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Получить первый рабочий лист.
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

            // Получить заголовки и значения.
            Excel.Range range;
            range = (Excel.Range)sheet.Cells[2, 2];
            double A= range.Value2;
            label1.Text = Convert.ToString(A);
            //label1.Text = (string)range.Value2;

            // Закройте книгу без сохранения изменений.
            workbook.Close(false, Type.Missing, Type.Missing);

            // Закройте сервер Excel.
            excel_app.Quit();
        }
    }
}
