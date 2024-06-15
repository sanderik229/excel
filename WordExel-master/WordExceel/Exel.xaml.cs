using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
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

namespace WordExceel
{
    /// <summary>
    /// Логика взаимодействия для Exel.xaml
    /// </summary>
    public partial class Exel : Window
    {
        private string filename2;
        public ObservableCollection<DynamicData> Data { get; set; }

        public Exel()
        {
            InitializeComponent();
            Data = new ObservableCollection<DynamicData>();
            ExcelGrid.ItemsSource = Data;
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string columnName = ColumnTbx.Text.Trim();
            if (!string.IsNullOrEmpty(columnName))
            {
                // Добавляем новую колонку в DataGrid
                var newColumn = new DataGridTextColumn
                {
                    Header = columnName,
                    Binding = new Binding($"[{columnName}]"),
                    IsReadOnly = false
                };
                ExcelGrid.Columns.Add(newColumn);

                // Добавляем новое поле в каждую строку данных
                foreach (var item in Data)
                {
                    item.Add(new KeyValuePair<string, object>(columnName, ""));
                }

                ColumnTbx.Clear();
            }
            else
            {
                MessageBox.Show("Пожалуйста, введите имя колонки", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            int m = 0;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Лист Excel (*.xlsx)|*.xlsx",
                FileName = $"Data {m++}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");


                    for (int i = 0; i < ExcelGrid.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = ExcelGrid.Columns[i].Header;
                    }


                    for (int i = 0; i < Data.Count; i++)
                    {
                        for (int j = 0; j < ExcelGrid.Columns.Count; j++)
                        {
                            var columnName = ExcelGrid.Columns[j].Header.ToString();
                            worksheet.Cells[i + 2, j + 1].Value = Data[i][columnName];
                        }
                    }


                    var fileInfo = new FileInfo(saveFileDialog.FileName);
                    package.SaveAs(fileInfo);
                }

                MessageBox.Show("Данные успешно сохранены", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public class DynamicData : ObservableCollection<KeyValuePair<string, object>>
        {
            public object this[string key]
            {
                get
                {
                    var skvp = this.FirstOrDefault(kvp => kvp.Key == key);
                    return skvp.Equals(default(KeyValuePair<string, object>)) ? null : skvp.Value;
                }
                set
                {
                    var existingKvp = this.FirstOrDefault(kvp => kvp.Key == key);
                    if (!existingKvp.Equals(default(KeyValuePair<string, object>)))
                    {
                        var index = this.IndexOf(existingKvp);
                        this[index] = new KeyValuePair<string, object>(key, value);
                    }
                    else
                    {
                        this.Add(new KeyValuePair<string, object>(key, value));
                    }
                }
            }
        }



        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            SendFile file = new SendFile(filename2);
            file.Show();

        }
    }
}
