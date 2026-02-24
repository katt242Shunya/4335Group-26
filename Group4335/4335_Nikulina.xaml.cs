using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace Group4335
{
    public partial class _4335_Nikulina : Window
    {
        private DatabaseHelper dbHelper;

        public _4335_Nikulina()
        {
            InitializeComponent();
            dbHelper = new DatabaseHelper();

            ExcelPackage.License.SetNonCommercialPersonal("4335_Nikulina Lab Work");
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    Title = "Выберите файл 2.xlsx"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;

                    var orders = ReadOrdersFromExcel(filePath);
                    dbHelper.SaveOrdersToDatabase(orders);

                    MessageBox.Show($"Импорт завершен! Загружено {orders.Count} записей.",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Order> ReadOrdersFromExcel(string filePath)
        {
            var orders = new List<Order>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                // ИСПРАВЛЕНО: начинаем с 1 строки, так как заголовков больше нет
                for (int row = 1; row <= rowCount; row++)
                {
                    try
                    {
                        var order = new Order
                        {
                            Id = Convert.ToInt32(worksheet.Cells[row, 1].Value ?? 0),
                            OrderCode = worksheet.Cells[row, 2].Value?.ToString() ?? "",
                            CreationDate = Convert.ToDateTime(worksheet.Cells[row, 3].Value ?? DateTime.Now),
                            ClientCode = worksheet.Cells[row, 4].Value?.ToString() ?? "",
                            Services = worksheet.Cells[row, 5].Value?.ToString() ?? "",
                            Status = worksheet.Cells[row, 6].Value?.ToString() ?? ""
                        };
                        orders.Add(order);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Ошибка в строке {row}: {ex.Message}");
                    }
                }
            }

            return orders;
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var orders = dbHelper.GetOrdersFromDatabase();

                if (orders.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта. Сначала импортируйте данные.",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Группируем ТОЛЬКО по существующим статусам, исключая пустые
                var groupedByStatus = orders
                    .Where(o => !string.IsNullOrWhiteSpace(o.Status)) // Отсеиваем пустые статусы
                    .GroupBy(o => o.Status)
                    .ToDictionary(g => g.Key, g => g.ToList());

                // Проверяем, остались ли статусы после фильтрации
                if (groupedByStatus.Count == 0)
                {
                    MessageBox.Show("Нет записей с заполненным статусом для экспорта.",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ExportToExcel(groupedByStatus);

                MessageBox.Show($"Экспорт завершен! Создано {groupedByStatus.Count} листов: {string.Join(", ", groupedByStatus.Keys)}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToExcel(Dictionary<string, List<Order>> groupedOrders)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = System.IO.Path.Combine(desktopPath,
                $"Экспорт_статусы_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            using (var package = new ExcelPackage())
            {
                foreach (var status in groupedOrders.Keys)
                {
                    string sheetName = status.Length > 30 ? status.Substring(0, 30) : status;

                    foreach (char invalidChar in Path.GetInvalidFileNameChars())
                    {
                        sheetName = sheetName.Replace(invalidChar, '_');
                    }

                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // Заголовки (они нужны в выходном файле, поэтому оставляем)
                    worksheet.Cells[1, 1].Value = "Id";
                    worksheet.Cells[1, 2].Value = "Код заказа";
                    worksheet.Cells[1, 3].Value = "Дата создания";
                    worksheet.Cells[1, 4].Value = "Код клиента";
                    worksheet.Cells[1, 5].Value = "Услуги";

                    using (var range = worksheet.Cells[1, 1, 1, 5])
                    {
                        range.Style.Font.Bold = true;
                    }

                    int row = 2; // Данные начинаем со 2 строки, так как в 1 строке заголовки
                    foreach (var order in groupedOrders[status])
                    {
                        worksheet.Cells[row, 1].Value = order.Id;
                        worksheet.Cells[row, 2].Value = order.OrderCode;
                        worksheet.Cells[row, 3].Value = order.CreationDate.ToString("dd.MM.yyyy HH:mm");
                        worksheet.Cells[row, 4].Value = order.ClientCode;
                        worksheet.Cells[row, 5].Value = order.Services;
                        row++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                package.SaveAs(new FileInfo(filePath));
            }

            MessageBox.Show($"Файл сохранен на рабочем столе:\n{filePath}", "Готово");
        }
    }
}