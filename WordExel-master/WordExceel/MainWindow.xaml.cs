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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordExceel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
    
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {



        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Exel exel = new Exel();
            exel.Show();
            Close();
        }

        private void ButtonForOpenExel(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Лист Excel|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string filename = openFileDialog.FileName;
                ExcelEmail exel = new ExcelEmail(filename);
                exel.Show();
                Close();
            }
        }
    }
}
