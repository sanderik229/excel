using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
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
    /// Логика взаимодействия для ExcelEmail.xaml
    /// </summary>
    public partial class ExcelEmail : Window
    {
        private string filename2;

        public ExcelEmail(string filename)
        {
            InitializeComponent();
            LoadFile(filename);
        }

        private void LoadFile(string filename)
        {
            if (File.Exists(filename))
            {
                // EPPlus License context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filename)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var dataTable = new DataTable();

                    // Добавляем столбцы
                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }


                    for (var rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                        var row = dataTable.NewRow();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                        dataTable.Rows.Add(row);
                    }

                    ExcelDataGrid.ItemsSource = dataTable.DefaultView;
                    filename2 = filename;
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SendFile sendFile = new SendFile(filename2);
            sendFile.Show();
        }

        private void SaveFile(string filename)
        {
            // EPPlus License context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Добавляем столбцы
                for (int i = 0; i < ExcelDataGrid.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = ExcelDataGrid.Columns[i].Header;
                }

                // Добавляем строки
                for (int i = 0; i < ExcelDataGrid.Items.Count; i++)
                {
                    for (int j = 0; j < ExcelDataGrid.Columns.Count; j++)
                    {
                        TextBlock cellContent = ExcelDataGrid.Columns[j].GetCellContent(ExcelDataGrid.Items[i]) as TextBlock;
                        worksheet.Cells[i + 2, j + 1].Value = cellContent.Text;
                    }
                }

                FileInfo fi = new FileInfo(filename);
                package.SaveAs(fi);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";

            if (dlg.ShowDialog() == true)
            {
                string filename = dlg.FileName;
                SaveFile(filename);
                MessageBox.Show("Файл успешно сохранён");
            }
        }
    }
}
