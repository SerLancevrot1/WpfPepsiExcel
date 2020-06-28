using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;
using System.IO;

namespace WpfPepsiExcel.buttons
{
   public class Butt2
    {
        public void method(DateTime dateTimePicker1, DateTime dateTimePicker2)
        {

            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(@"C:\Users\" +
                Environment.UserName + @"\Templates\ReportsMain.xlsm"))
                {
                    var hash = md5.ComputeHash(stream);
                    string qwer = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();


                }
            }

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(@"C:\Users\" +
                Environment.UserName + @"\Templates\Template3.xlsx");




            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            Excel.Worksheet ws2 = (Excel.Worksheet)wb.Worksheets[2];  //Worksheets[1]; "Счетчики Opr"

            ws1.Cells[2, 2] = dateTimePicker1;
            ws2.Cells[2, 2] = dateTimePicker2;

            int year = dateTimePicker1.Year; //для подключения к коллекциям
            int month = dateTimePicker1.Month;
            int year2 = dateTimePicker2.Year;
            int month2 = dateTimePicker2.Month;
            string date1 = new DateTime(year, month, 1).ToShortDateString();
            string date2 = new DateTime(year2, month2, 1).ToShortDateString();


            IMongoDatabase database = MongoConnect.Con(); //подключение к дб

            DateTime fixedDTP1 = dateTimePicker1;
            DateTime fixedDTP2 = dateTimePicker2 ;

            FilterDefinition<MongoNode> MainFilter1 =
                 Builders<MongoNode>.Filter.Gte("dateTime", fixedDTP1);
            List<MongoNode> mainList1 =
                database.GetCollection<MongoNode>(date1).Find(MainFilter1).Limit(80).ToList();

            FilterDefinition<MongoNode> MainFilter2 =
                Builders<MongoNode>.Filter.Gte("dateTime", fixedDTP2);
            List<MongoNode> mainList2 =
                database.GetCollection<MongoNode>(date2).Find(MainFilter2).Limit(80).ToList();

            for (int i = 1; i < 64; i++)
            {
                // фильтрую листы с лямбда выражениями
                List<MongoNode> list1 = mainList1.Where(x => x.ID == i).ToList();

                foreach (var j in list1)//цикл столбцов
                {
                    ws1.Cells[i + 5, 2] = j.wP_in/1000;
                    ws1.Cells[i + 5, 3] = j.WP_out/1000;
                    ws1.Cells[i + 5, 4] = j.WQ_in / 1000;// заполнение листов
                    ws1.Cells[i + 5, 5] = j.WQ_oup / 1000;
                    ws1.Cells[i + 5, 6] = j.WQ / 1000;
                    break;
                }

                List<MongoNode> list2 = mainList2.Where(x => x.ID == i).ToList();
                foreach (var j in list2)
                {
                    ws2.Cells[i + 5, 2] = j.wP_in / 1000;
                    ws2.Cells[i + 5, 3] = j.WP_out / 1000;
                    ws2.Cells[i + 5, 4] = j.WQ_in / 1000;
                    ws2.Cells[i + 5, 5] = j.WQ_oup / 1000;
                    ws2.Cells[i + 5, 6] = j.WQ / 1000;
                    break;
                }
            }

            //дата для отчета
            string dateTime = DateTime.Now.ToShortDateString();

            string mainPath = Path.FirstPath();
            string xls = ".xls";
            string fileName = "QWERT__" + dateTime + xls;
            try
            {
                wb.SaveAs(mainPath + fileName, Excel.XlFileFormat.xlWorkbookNormal);
            }
            catch
            {
                MessageBox.Show("Не удалось сохранить файл");
                wb.Close();
                return;
            }
            wb.Close();
            MessageBox.Show("Готово");
        }
    }
}
