using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace PaymentExampleApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bdEntities _context = new bdEntities();
        private User currentUser;

        public MainWindow()
        {
            InitializeComponent();
            ChartPayments.ChartAreas.Add(new ChartArea("Main"));

            var currentSeries = new Series("Payments")
            {
                IsValueShownAsLabel = true

            };
            ChartPayments.Series.Add(currentSeries);

            ComboUsers.ItemsSource = _context.User.ToList();
            ComboChartTypes.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
        }

        public SeriesChartType currentType { get; private set; }

        private ComboBox GetComboChartTypes()
        {
            return ComboChartTypes;
        }

        private SeriesChartType GetCurrentType()
        {
            return currentType;
        }


        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (ComboUsers.SelectedItem is User && ComboChartTypes.SelectedItem is SeriesChartType) { }

            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault();
                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categoriesList = _context.Category.ToList();
                foreach (var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.Name, _context.Payment.ToList().Where(p => p.User == currentUser && p.Category== category).Sum(p => p.Price * p.Num));
                }
            }
        }

        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = _context.User.ToList().OrderBy(p => p.FIO).ToList();

            var aplication = new Excel.Application();
            aplication.SheetsInNewWorkbook = allUsers.Count();

            Excel.Workbook workbook = aplication.Workbooks.Add(Type.Missing);

            int startRowIndex = 1;

            for (int i = 0; i < allUsers.Count(); i++)
            {
                Excel.Worksheet worksheet = aplication.Worksheets.Item[i + 1];
                worksheet.Name = allUsers[i].FIO;

                worksheet.Cells[1][startRowIndex] = "Payment date";
                worksheet.Cells[2][startRowIndex] = "Titel";
                worksheet.Cells[3][startRowIndex] = "Cost";
                worksheet.Cells[4][startRowIndex] = "Number";
                worksheet.Cells[5][startRowIndex] = "Summary";

                startRowIndex++;

                var usersCAtegories = allUsers[i].Payment.OrderBy(p => p.Date).GroupBy(p => p.Category).OrderBy(p => p.Key.Name);

                foreach (var groupCategory in usersCAtegories)
                {
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[i][startRowIndex], worksheet.Cells[5][startRowIndex]];
                }
            }
        }
    }
   
}
