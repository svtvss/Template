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
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Sosnovskaya_4337.xaml
    /// </summary>
    public partial class Sosnovskaya_4337 : Window
    {
        private const int _sheetsCount = 3;
        public Sosnovskaya_4337()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
            {
                return;
            }

            string[,] list; //for data in excel
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (ISRPO_LR2Entities db = new ISRPO_LR2Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    db.Services.Add(new Services() 
                    { 
                        IdServices = int.Parse(list[i, 0]),
                        NameServices = list[i, 1], 
                        TypeOfService = list[i, 2], 
                        CodeService = list[i, 3], 
                        Cost = int.Parse(list[i, 4]) 
                    });
                    MessageBox.Show($"{list[i, 1]}");
                }
                MessageBox.Show("Успешно!");
                db.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Services> service;

            using (ISRPO_LR2Entities db = new ISRPO_LR2Entities())
            {
                service = db.Services.ToList();
            }

            List<string[]> TypeServices = new List<string[]>() { //for sheets name
                new string[]{ "Прокат" },
                new string[]{ "Обучение" },
                new string[]{ "Подъем" },
            };

            var app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = _sheetsCount;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < _sheetsCount; i++)
            {
                int startRowIndex = 1;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория - {TypeServices[i][0]}";

                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.Merge();
                headerRange.Value = $"Категория - {TypeServices[i][0]}";
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Стоимость";

                startRowIndex++;

                foreach (Services services in service.OrderBy(a => a.Cost))
                {
                    if (services.TypeOfService == TypeServices[i][0])
                    {
                        worksheet.Cells[1][startRowIndex] = services.IdServices;
                        worksheet.Cells[2][startRowIndex] = services.NameServices;
                        worksheet.Cells[3][startRowIndex] = services.Cost;
                        startRowIndex++;
                    }
                }

                Microsoft.Office.Interop.Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        class ServicesJSON
        {
            public int IdServices { get; set; }
            public string NameServices { get; set; }
            public string TypeOfService { get; set; }
            public string CodeService { get; set; }
            public int Cost { get; set; }
        }

        private void BnImportJSON_Click(object sender, RoutedEventArgs e)
        {
            string json = File.ReadAllText(@"C:\Users\tanya\OneDrive\Рабочий стол\Импорт\1.json");
            var service = JsonSerializer.Deserialize<List<ServicesJSON>>(json);
            using (ISRPO_LR2Entities entities = new ISRPO_LR2Entities())
            {
                foreach (ServicesJSON serviceJSON in service)
                {
                    try
                    {
                        entities.Services.Add(new Services()
                        {
                            IdServices = serviceJSON.IdServices,
                            NameServices = serviceJSON.NameServices,
                            TypeOfService = serviceJSON.TypeOfService,
                            CodeService = serviceJSON.CodeService,
                            Cost = serviceJSON.Cost
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                MessageBox.Show("Успешно!");
                entities.SaveChanges();
            }
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Services> servece;

            using (ISRPO_LR2Entities db = new ISRPO_LR2Entities())
            {
                servece = db.Services.ToList();
            }

            var app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = app.Documents.Add();

            for (int i = 0; i < _sheetsCount; i++)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range range = paragraph.Range;

                List<string[]> ServecCategories = new List<string[]>() { //for sheets name
                    new string[]{ "Прокат" },
                    new string[]{ "Обучение" },
                    new string[]{ "Подъем" },
                };

                var data = i == 0 ? servece.Where(o => o.TypeOfService == "Прокат")
                        : i == 1 ? servece.Where(o => o.TypeOfService == "Обучение")
                        : i == 2 ? servece.Where(o => o.TypeOfService == "Подъем") : servece; //sort for task
                List<Services> currentSer = data.ToList();
                int countStaffsInCategory = currentSer.Count();

                Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                Microsoft.Office.Interop.Word.Table serTable = document.Tables.Add(tableRange, countStaffsInCategory + 1, 3);
                serTable.Borders.InsideLineStyle =
                serTable.Borders.OutsideLineStyle =
                Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                serTable.Range.Cells.VerticalAlignment =
                Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                range.Text = Convert.ToString($"Категория - {ServecCategories[i][0]}");
                range.InsertParagraphAfter();

                Microsoft.Office.Interop.Word.Range cellRange = serTable.Cell(1, 1).Range;
                cellRange.Text = "ID";
                cellRange = serTable.Cell(1, 2).Range;
                cellRange.Text = "Название услуги";
                cellRange = serTable.Cell(1, 3).Range;
                cellRange.Text = "Стоимость";
                serTable.Rows[1].Range.Bold = 1;
                serTable.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int j = 1;
                foreach (var currentServe in currentSer.OrderBy(a => a.Cost))
                {
                    cellRange = serTable.Cell(j + 1, 1).Range;
                    cellRange.Text = $"{currentServe.IdServices}";
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = serTable.Cell(j + 1, 2).Range;
                    cellRange.Text = currentServe.NameServices;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = serTable.Cell(j + 1, 3).Range;
                    cellRange.Text = currentServe.Cost.ToString();
                    cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    j++;
                }

                if (i > 0)
                {
                    range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                }
            }
            app.Visible = true;
        }
    }
}
