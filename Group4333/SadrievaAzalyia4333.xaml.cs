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
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System.IO;

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

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            using (var db = new Context())
            {
                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    try
                    {
                        db.Employees.Add(new Employees
                        {
                            EmployeeCode = ((Excel.Range)range.Cells[i, 1]).Text.ToString(),
                            Position = ((Excel.Range)range.Cells[i, 2]).Text.ToString(),
                            FullName = ((Excel.Range)range.Cells[i, 3]).Text.ToString(),
                            Login = ((Excel.Range)range.Cells[i, 4]).Text.ToString(),
                            Password = ((Excel.Range)range.Cells[i, 5]).Text.ToString(),
                            LastEntry = ((Excel.Range)range.Cells[i, 6]).Text.ToString(),
                            EntryType = ((Excel.Range)range.Cells[i, 7]).Text.ToString()
                        });
                    }
                    catch { /* Пропуск пустых строк */ }
                }
                db.SaveChanges();
            }
            workbook.Close(false);
            app.Quit();
            MessageBox.Show("Импорт завершен");
        }
        private void BtnImportJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "JSON|*.json" };
            if (ofd.ShowDialog() != true) return;

            try
            {
                string jsonContent = File.ReadAllText(ofd.FileName);
                List<Employees> importedEmployees = JsonConvert.DeserializeObject<List<Employees>>(jsonContent);

                if (importedEmployees != null)
                {
                    using (var db = new Context())
                    {
                        foreach (var emp in importedEmployees)
                        {
                            db.Employees.Add(emp);
                        }
                        db.SaveChanges();
                    }
                    MessageBox.Show("Данные из JSON успешно импортированы в БД");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при импорте JSON: " + ex.Message);
            }
        }

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Employees> allData;
            using (var db = new Context())
            {
                allData = db.Employees.ToList();
            }

            if (!allData.Any())
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            var groupedData = allData.GroupBy(x => x.EntryType).ToList();

            var wordApp = new Word.Application();
            Word.Document document = wordApp.Documents.Add();

            foreach (var group in groupedData)
            {
                Word.Paragraph groupHeader = document.Paragraphs.Add();
                groupHeader.Range.Text = $"Тип входа: {group.Key ?? "Не указан"}";
                groupHeader.Range.Font.Bold = 1;
                groupHeader.Range.Font.Size = 14;
                groupHeader.Range.InsertParagraphAfter();

                Word.Table table = document.Tables.Add(groupHeader.Range, group.Count() + 1, 3);
                table.Borders.Enable = 1;

                table.Cell(1, 1).Range.Text = "Код сотрудника";
                table.Cell(1, 2).Range.Text = "ФИО";
                table.Cell(1, 3).Range.Text = "Логин";
                table.Rows[1].Range.Font.Bold = 1; 

                int rowIndex = 2;
                foreach (var emp in group)
                {
                    table.Cell(rowIndex, 1).Range.Text = emp.EmployeeCode;
                    table.Cell(rowIndex, 2).Range.Text = emp.FullName;
                    table.Cell(rowIndex, 3).Range.Text = emp.Login;
                    rowIndex++;
                }

                if (group != groupedData.Last())
                {
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
            }

            wordApp.Visible = true;
            MessageBox.Show("Экспорт в Word по типам входа завершен");
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
                Excel.Worksheet sh = (Excel.Worksheet)wb.Worksheets[i + 1];
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
