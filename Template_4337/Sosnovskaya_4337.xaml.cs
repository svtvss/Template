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
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

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

            using (Entities entities = new Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    entities.services.Add(new services() 
                    { 
                        ID = int.Parse(list[i, 0]), 
                        Name_service = list[i, 1], 
                        Type_service = list[i, 2], 
                        Code_service = list[i, 3], 
                        Price = int.Parse(list[i, 4]) 
                    });
                }
                MessageBox.Show("Успешно!");
                entities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<services> service;

            using (Entities entities = new Entities())
            {
                service = entities.services.ToList();
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

                foreach (services services in service.OrderBy(a => a.Price))
                {
                    if (services.Type_service == TypeServices[i][0])
                    {
                        worksheet.Cells[1][startRowIndex] = services.ID;
                        worksheet.Cells[2][startRowIndex] = services.Name_service;
                        worksheet.Cells[3][startRowIndex] = services.Price;
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
    }
}
