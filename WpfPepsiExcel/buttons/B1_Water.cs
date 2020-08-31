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
    class B1_Water
    {
        public void method(DateTime dateTimePicker1, DateTime dateTimePicker2)
        {

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(@"C:\Users\" +
                Environment.UserName + @"\Templates\Water11.xlsx");

            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            
            ws1.Cells[2, 6] = dateTimePicker1;
            ws1.Cells[2, 7] = dateTimePicker2;

            int year = dateTimePicker1.Year; //для подключения к коллекциям
            int month = dateTimePicker1.Month;
            int year2 = dateTimePicker2.Year;
            int month2 = dateTimePicker2.Month;
            string date1 = new DateTime(year, month, 1).ToShortDateString();
            string date2 = new DateTime(year2, month2, 1).ToShortDateString();

            IMongoDatabase database = MongoConnect.ConWater(); //подключение к дб

            DateTime fixedDTP1 = dateTimePicker1;
            DateTime fixedDTP2 = dateTimePicker2;

            FilterDefinition<MongoNodeWater> MainFilter1 =
                 Builders<MongoNodeWater>.Filter.Gte("dateTime", fixedDTP1);
            List<MongoNodeWater> mainList1 =
                database.GetCollection<MongoNodeWater>(date1).Find(MainFilter1).Limit(80).ToList();

            FilterDefinition<MongoNodeWater> MainFilter2 =
                Builders<MongoNodeWater>.Filter.Gte("dateTime", fixedDTP2);
            List<MongoNodeWater> mainList2 =
                database.GetCollection<MongoNodeWater>(date2).Find(MainFilter2).Limit(80).ToList();

            for (int i = 1; i < 12; i++)
            {
                // фильтрую листы с лямбда выражениями
                List<MongoNodeWater> list1 = mainList1.Where(x => x.ID == i).ToList();

                foreach (var j in list1)//цикл столбцов
                {
                    ws1.Cells[i + 8, 2] = j.value;
                    break;
                }

                List<MongoNodeWater> list2 = mainList2.Where(x => x.ID == i).ToList();
                foreach (var j in list2)
                {
                    ws1.Cells[i + 8, 3] = j.value;
                    break;
                }
            }

            //дата для отчета
            string dateTime = DateTime.Now.ToShortDateString();

            string mainPath = Path.FirstPath();
            string xls = ".xls";
            string fileName = "Water__" + dateTime + xls;
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
