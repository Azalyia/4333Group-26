using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Group4333
{
    public partial class SadrievaAzalyia4333 : Window
    {
        public SadrievaAzalyia4333()
        {
            InitializeComponent();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel|*.xlsx" };
            if (ofd.ShowDialog() != true) return;

            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(ofd.FileName);
            Excel.Worksheet worksheet = workbook.Sheets[1];
            var range = worksheet.UsedRange;

            using (var db = new Context())
            {
                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    try
                    {
                        db.Employees.Add(new Employees
                        {
                            EmployeeCode = worksheet.Cells[i, 1].Text,
                            Position = worksheet.Cells[i, 2].Text,
                            FullName = worksheet.Cells[i, 3].Text,
                            Login = worksheet.Cells[i, 4].Text,
                            Password = worksheet.Cells[i, 5].Text,
                            LastEntry = worksheet.Cells[i, 6].Text,
                            EntryType = worksheet.Cells[i, 7].Text 
                        });
                    }
                    catch { continue; }
                }
                db.SaveChanges();
            }
            workbook.Close(false);
            app.Quit();
            MessageBox.Show("Импорт завершен");
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Employees> allData;
            using (var db = new Context()) { allData = db.Employees.ToList(); }

            var entryTypes = allData.Select(x => x.EntryType).Distinct().ToList();

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = entryTypes.Count;
            Excel.Workbook wb = app.Workbooks.Add();

            for (int i = 0; i < entryTypes.Count; i++)
            {
                string currentType = entryTypes[i];
                Excel.Worksheet sh = wb.Worksheets[i + 1];
                sh.Name = currentType.Length > 30 ? currentType.Substring(0, 30) : currentType; 

                sh.Cells[1, 1] = "Код клиента";
                sh.Cells[1, 2] = "Должность";
                sh.Cells[1, 3] = "Логин";

                var filteredData = allData.Where(x => x.EntryType == currentType).ToList();
                int row = 2;
                foreach (var emp in filteredData)
                {
                    sh.Cells[row, 1] = emp.EmployeeCode;
                    sh.Cells[row, 2] = emp.Position;
                    sh.Cells[row, 3] = emp.Login;
                    row++;
                }
                sh.Columns.AutoFit();
            }
            app.Visible = true;
            MessageBox.Show("Экспорт завершен");
        }
    }
}
