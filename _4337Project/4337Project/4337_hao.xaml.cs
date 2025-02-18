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
using Microsoft.Win32; 
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.SqlClient;

namespace _4337Project
{
    /// <summary>
    /// Логика взаимодействия для _4337_hao.xaml
    /// </summary>
    public partial class _4337_hao : Window
    {
        public _4337_hao()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string connectionString = "data source=(localdb)\\MSSQLLocalDB;Initial Catalog=Status_hao;Integrated Security=True;";
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                string tableName = "Table_" + fileName;
                try
                {
                    hao_Import_4var.ImportData(filePath, connectionString, tableName);
                    MessageBox.Show("Данные успешно импортированы!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте данных: {ex.Message}");
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = "data source=(localdb)\\MSSQLLocalDB;Initial Catalog=Status_hao;Integrated Security=True;";
            string tableName = "Table_2";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.FileName = "ExportBD_hao4var.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputFilePath = saveFileDialog.FileName;

                try
                {
                    hao_Export_4var.ExportData(connectionString, tableName, outputFilePath);
                    MessageBox.Show("Данные успешно экспортированы по статусу");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте данных: {ex.Message}");
                }
            }
        }
    }
}
