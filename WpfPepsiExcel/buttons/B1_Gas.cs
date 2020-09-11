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
    class B1_Gas
    {
        public void excelWorker(DateTime dateTimePicker1, DateTime dateTimePicker2)
        {

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(@"C:\Users\" +
                Environment.UserName + @"\Templates\Gfm1.xlsx");

            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            
            ws1.Cells[1, 1] = dateTimePicker1;
            ws1.Cells[1, 2] = dateTimePicker2;

            int year = dateTimePicker1.Year; //для подключения к коллекциям
            int month = dateTimePicker1.Month;
            int year2 = dateTimePicker2.Year;
            int month2 = dateTimePicker2.Month;
            string date1 = new DateTime(year, month, 1).ToShortDateString();
            string date2 = new DateTime(year2, month2, 1).ToShortDateString();

            IMongoDatabase database = MongoConnect.ConGas(); //подключение к дб

            DateTime fixedDTP1 = dateTimePicker1;
            DateTime fixedDTP2 = dateTimePicker2;

            FilterDefinition<MongoNodeGas> MainFilter1 =
                 Builders<MongoNodeGas>.Filter.Gte("dateTime", fixedDTP1);
            List<MongoNodeGas> mainList1 =
                database.GetCollection<MongoNodeGas>(date1).Find(MainFilter1).Limit(15).ToList();

            FilterDefinition<MongoNodeGas> MainFilter2 =
                Builders<MongoNodeGas>.Filter.Gte("dateTime", fixedDTP2);
            List<MongoNodeGas> mainList2 =
                database.GetCollection<MongoNodeGas>(date2).Find(MainFilter2).Limit(15).ToList();

            for (int i = 1; i < 8; i++)
            {
                // фильтрую листы с лямбда выражениями
                List<MongoNodeGas> list1 = mainList1.Where(x => x.ID == i).ToList();

                foreach (var j in list1)//цикл столбцов
                {
                    if (j.ID >= 5)
                    {
                        ws1.Cells[i + 1, 1] = (bool)j.IsWork;
                    }
                    else
                    {
                        ws1.Cells[i + 1, 1] = j.value;
                    }
                    
                    break;
                }

                List<MongoNodeGas> list2 = mainList2.Where(x => x.ID == i).ToList();
                foreach (var j in list2)
                {

                    if (j.ID >= 5)
                    {
                        ws1.Cells[i + 1, 2] = (bool)j.IsWork;
                    }
                    else
                    {
                        ws1.Cells[i + 1, 2] = j.value;
                    }

                    break;
                }
            }

            //дата для отчета
            string dateTime = DateTime.Now.ToShortDateString();

            string mainPath = Path.FirstPath();
            string xls = ".xls";
            string fileName = "Gas__" + dateTime + xls;

            // скрыть лист
            ws1.Visible = Excel.XlSheetVisibility.xlSheetHidden;

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
