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
using System.Text.RegularExpressions;

namespace ExceloDopisywarka
{
    public static class Okno
    {
        public static string nazwaPlikuExcel;
        public static Guid guidBaza;
    }
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        

        public MainWindow()
        {
            InitializeComponent();
            wczytaZakładkęButton.Visibility = System.Windows.Visibility.Hidden;
            listaZakładekComboBox.Visibility = System.Windows.Visibility.Hidden;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog();
            fileDialog.Filter = "Excel (*.xls,*.xlsx)|*.xls;*.xlsx";
            Nullable<bool> excelFile = fileDialog.ShowDialog();
            if (excelFile == true)
            {
                nazwaPlikuExcelForm.Text = "Wczytano plik: " + fileDialog.FileName;
                Okno.nazwaPlikuExcel = fileDialog.FileName;
                List<string> zakładkiExcela = wczytajlistęZakładek(fileDialog.FileName);
                nazwaPlikuExcelForm.Text = "Wczytano plik: " + fileDialog.FileName
                        + " - wybierz zakładkę do doczytania";
                wczytaZakładkęButton.Visibility = System.Windows.Visibility.Visible;
                listaZakładekComboBox.Visibility = System.Windows.Visibility.Visible;
                listaZakładekComboBox.ItemsSource = zakładkiExcela;
            }
        }

        private List<string> wczytajlistęZakładek(string nazwaPliku)
        {
            List<string> listazakładek = new List<string>();
            using (System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;"
                +"Data Source='"+nazwaPliku+"';Extended Properties='Excel 12.0 Xml;HDR=YES'"))
            {
                connection.Open();
                var infoTabela = connection.GetSchema("Tables");
                foreach (System.Data.DataRow infoTabelaWiersz in infoTabela.Rows)
                {
                    if (Regex.IsMatch(infoTabelaWiersz["TABLE_NAME"].ToString(), @"\$") == true)
                        listazakładek.Add(infoTabelaWiersz["TABLE_NAME"].ToString());
                }
            }

            return listazakładek;
        }

        private void wczytaZakładkęButton_Click(object sender, RoutedEventArgs e)
        {
            using (System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;"
                + "Data Source='" + Okno.nazwaPlikuExcel + "';Extended Properties='Excel 12.0 Xml;HDR=YES'"))
            {
                System.Data.DataSet dataset = new System.Data.DataSet();
                System.Data.OleDb.OleDbDataAdapter command = new System.Data.OleDb.OleDbDataAdapter("select * from [" 
                    + listaZakładekComboBox.SelectedValue + "]", connection);
                command.Fill(dataset);

                using (RabatyEntities1 rabaty = new RabatyEntities1())
                {
                    Dopisywarka dopisywarka = new Dopisywarka() { Dane = dataset.GetXml() };
                    rabaty.Dopisywarka.Add(dopisywarka);
                    rabaty.SaveChanges();
                }

            }
            //var c = Dataset.GetXml();
        }
    }
}
