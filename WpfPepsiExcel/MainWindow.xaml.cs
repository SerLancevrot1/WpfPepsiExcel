
using Microsoft.VisualBasic;
using System;
using System.Threading.Tasks;
using System.Windows;
using WpfPepsiExcel.buttons;

namespace WpfPepsiExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        

        public MainWindow()
        {
            InitializeComponent();
        }

        //кнопки для создания отчетов


        private async void B1_Electricity_Click(object sender, RoutedEventArgs e)
        {

            if (dateTimePicker1.Value == null || dateTimePicker2.Value == null)
            {
                MessageBox.Show("Введите число");
                return;
            }

            B1_Elecricity butt = new B1_Elecricity();
            DateTime dateTime1 = dateTimePicker1.Value.Value;
            
            DateTime dateTime2 = dateTimePicker2.Value.Value;

            await Task.Run(() => butt.method(dateTime1, dateTime2));
          
        }

        private async void B1_Water_Click(object sender, RoutedEventArgs e)
        {
            if (dateTimePicker3.Value == null || dateTimePicker4.Value == null)
            {
                MessageBox.Show("Введите число");
                return;
            }

            B1_Water butt = new B1_Water();
            DateTime dateTime3 = dateTimePicker3.Value.Value;
            DateTime dateTime4 = dateTimePicker4.Value.Value;

            await Task.Run(() => butt.method(dateTime3, dateTime4));
        }

        private async void B2_Water_Click(object sender, RoutedEventArgs e)
        {
            if (dateTimePicker3.Value == null || dateTimePicker4.Value == null)
            {
                MessageBox.Show("Введите число");
                return;
            }

            B1_Elecricity butt = new B1_Elecricity();
            DateTime dateTime1 = dateTimePicker1.Value.Value;

            DateTime dateTime2 = dateTimePicker2.Value.Value;

            await Task.Run(() => butt.method(dateTime1, dateTime2));
        }
    }
}