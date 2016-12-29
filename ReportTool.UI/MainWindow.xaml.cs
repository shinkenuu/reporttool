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
using ReportTool.Core;


namespace ReportTool.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string[] EvolutionReports = new string[]
        {
            "CV Weekly",
            "CV Monthly",
            "MSRP History",
            "Weighted MSRP"
        };

        public ReportCommand Command { get; private set; }


        public MainWindow()
        {
            InitializeComponent();
            radioButton_EvolutionReport.Checked += OnEvolutionChecked;
            Command = new ReportCommand() { EarlyDate = DateTime.MinValue, LaterDate = DateTime.MaxValue };
        }

        
        private void OnEvolutionChecked(object sender, EventArgs e)
        {
            Command.ReportFormat = "evolution";
            comboBox_ReportName.ItemsSource = EvolutionReports;
            comboBox_ReportName.SelectedIndex = 0;
        }


        private void ComboBox_ReportName_SelectionChanged(object sender, EventArgs e)
        {
             Command.ReportName = comboBox_ReportName.SelectedItem.ToString().ToLower().Replace(" ", "");
        }



        private void DatePicker_EarlyDate_SelectedDateChanged(object sender, EventArgs e)
        {
            // ... Get DatePicker reference.
            var picker = sender as DatePicker;
            
            Command.EarlyDate = picker.SelectedDate ?? DateTime.MinValue;
        }



        private void DatePicker_LaterDate_SelectedDateChanged(object sender, EventArgs e)
        {
            // ... Get DatePicker reference.
            var picker = sender as DatePicker;

            Command.LaterDate = picker.SelectedDate ?? DateTime.MaxValue;
        }


        private void Button_Report_Click(object sender, EventArgs e)
        {
            if(Command.IsReady)
            {
                Program.Main(Command);
                MessageBox.Show("Report Complete");
            }
        }
    }


}
