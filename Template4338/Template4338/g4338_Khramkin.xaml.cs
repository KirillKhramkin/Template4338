using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;

namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для g4338_Khramkin.xaml
    /// </summary>
    public partial class g4338_Khramkin : Window
    {
        public g4338_Khramkin()
        {
            InitializeComponent();
        }
        private void ImportExcel(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls; *xlsx",
                Title = "Выберите файлы excel для импорта в базу данных",
            };

            var result = openFileDialog.ShowDialog();

            if (!result.HasValue || !result.Value)
                return;

            var excelWork = new Excel.Application();
            var bookWork = excelWork.Workbooks.Open(openFileDialog.FileName);

            var bookWorkSheet = (Excel.Worksheet)bookWork.Sheets[1];
            var lastCell = bookWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            var columns = lastCell.Column;
            var rows = bookWorkSheet.Cells[bookWorkSheet.Rows.Count, 1].End(-4162).Row;

            var list = new string[rows, columns];

            for (var i = 0; i < columns; i++)
                for (var j = 0; j < rows; j++)
                    list[j, i] = bookWorkSheet.Cells[j + 1, i + 1].Text;

            var orders = new List<Order>();
            MessageBox.Show($"{rows}");

            for (var i = 1; i < rows; i++)
            {
                var tempOrder = new Order();

                tempOrder.CodeStaff = list[i, 0];
                tempOrder.Position = list[i, 1];
                tempOrder.FullName = list[i, 2];
                tempOrder.Log = list[i, 3];
                tempOrder.Password = list[i, 4];
                tempOrder.LastEnter = list[i, 5];
                tempOrder.TypeEnter = list[i, 6];

                orders.Add(tempOrder);
            }

            System.GC.Collect();

            try
            {
                using (var context = new isrpo3Context())
                {
                    context.Order.AddRange(orders);
                    context.SaveChanges();
                }

                MessageBox.Show($"Добавление в базу данных прошло успешно {orders.Count}");
            }
            catch
            {
                MessageBox.Show("Ошибка базы данных");
            }
        }

        private void ExportExcel(object sender, RoutedEventArgs e)
        {
            const int idCol = 1;
            const int codeOrderCol = 2;
            const int dateOfCreateCol = 3;


            using (var context = new isrpo3Context())
            {
                var status = context.Order.ToList();
                List<Order> orders = new List<Order>();
                foreach (var order in status)
                {
                    if (order.Position == "Администратор")
                    {
                        orders.Add(order);
                    }
                }
                foreach (var order in status)
                {
                    if (order.Position == "Продавец")
                    {
                        orders.Add(order);
                    }
                }
                foreach (var order in status)
                {
                    if(order.Position == "Старший смены")
                    {
                        orders.Add(order);
                    }
                }
                
                
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = status.Count;
                var workbook = app.Workbooks.Add(Type.Missing);

                for (var i = 0; i < status.Count; i++)
                {
                    var worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = $"List{i}";

                    var startIndexRow = 2;

                    worksheet.Cells[idCol][1] = "Код клиента";
                    worksheet.Cells[codeOrderCol][1] = "Фио";
                    worksheet.Cells[dateOfCreateCol][1] = "Логин";


                    foreach (var item in orders)
                    {
                        worksheet.Cells[idCol][startIndexRow] = item.CodeStaff;
                        worksheet.Cells[codeOrderCol][startIndexRow] = item.FullName;
                        worksheet.Cells[dateOfCreateCol][startIndexRow] = item.Log;
                        startIndexRow++;
                    }
                }

                app.Visible = true;
            }
        }

        private async void ImportJson(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Title = "Выберите файлы json для импорта в базу данных",
            };

            var result = openFileDialog.ShowDialog();

            if (!result.HasValue || !result.Value)
                return;

            var orders = new List<Order>();

            using (var fs = new FileStream(openFileDialog.FileName, FileMode.OpenOrCreate))
            {
                orders = await JsonSerializer.DeserializeAsync<List<Order>>(fs);
            }


            try
            {
                using (var context = new isrpo3Context())
                {
                    await context.Order.AddRangeAsync(orders);
                    await context.SaveChangesAsync();
                    MessageBox.Show("Импортировано в базу данных");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка в доблавение в базу данных {ex.Message}");
            }
        }

        private void ExportWord(object sender, RoutedEventArgs e)
        {
            const int idCol = 1;
            const int codeOrderCol = 2;
            const int dateOfCreateCol = 3;

             using (var context = new isrpo3Context())
            {

                var app = new Word.Application();
                var document = app.Documents.Add();
                var status = context.Order.ToList();
                List<Order> orders0 = new List<Order>();
                List<Order> orders1 = new List<Order>();
                List<Order> orders2 = new List<Order>();
                foreach (var order in status)
                {
                    if (order.Position == "Администратор")
                    {
                        orders0.Add(order);
                    }
                }
                foreach (var order in status)
                {
                    if (order.Position == "Продавец")
                    {
                        orders1.Add(order);
                    }
                }
                foreach (var order in status)
                {
                    if (order.Position == "Старший смены")
                    {
                        orders2.Add(order);
                    }
                }
                List<List<Order>> MegaOrders = new List<List<Order>>();
                MegaOrders.Add(orders0);
                MegaOrders.Add(orders1);
                MegaOrders.Add(orders2);
                foreach (var orders in MegaOrders)
                {
                    var startIndexRow = 2;

                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = orders[0].Position;
                    range.InsertParagraphAfter();

                    var talbe = document.Paragraphs.Add();
                    var tableRange = talbe.Range;
                    var table = document.Tables.Add(tableRange, orders.Count() + 1, 3);
                    table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    table.Cell(1, idCol).Range.Text = "Код сотрудника";
                    table.Cell(1, codeOrderCol).Range.Text = "ФИО";
                    table.Cell(1, dateOfCreateCol).Range.Text = "Логин";

                    foreach (var item in orders)
                    {
                        table.Cell(startIndexRow, idCol).Range.Text = item.CodeStaff;
                        table.Cell(startIndexRow, codeOrderCol).Range.Text = item.FullName;
                        table.Cell(startIndexRow, dateOfCreateCol).Range.Text = item.Log;

                        startIndexRow++;
                    }

                    table.AllowAutoFit = true;
                    tableRange.InsertParagraphAfter();

                    app.Visible = true;
                }
            }
        }
    }
}
