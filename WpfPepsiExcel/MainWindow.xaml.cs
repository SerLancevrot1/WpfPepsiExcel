
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

        //кнопка создания отчета
        private void B1_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private async void B2_Click(object sender, RoutedEventArgs e)
        {

            if (dateTimePicker1.Value == null || dateTimePicker2.Value == null)
            {
                MessageBox.Show("Введите число");
                return;
            }

            Butt2 butt2 = new Butt2();
            DateTime dateTime1 = dateTimePicker1.Value.Value;
            
            DateTime dateTime2 = dateTimePicker2.Value.Value;

            await Task.Run(() => butt2.method(dateTime1, dateTime2));
          
        }
    }
}