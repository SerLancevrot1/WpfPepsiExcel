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
    class B2_Gas
    {
        public void excelWorker(DateTime dateTimePicker1, DateTime dateTimePicker2)
        {

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(@"C:\Users\" +
                Environment.UserName + @"\Templates\isLineWork1.xlsx");

            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            
            ws1.Cells[2, 1] = dateTimePicker1;
            

            int year = dateTimePicker1.Year; //для подключения к коллекциям
            int month = dateTimePicker1.Month;
            int year2 = dateTimePicker2.Year;
            int month2 = dateTimePicker2.Month;
            string date1 = new DateTime(year, month, 1).ToShortDateString();
            string date2 = new DateTime(year2, month2, 1).ToShortDateString();

            IMongoDatabase database = MongoConnect.ConGas(); //подключение к дб

            DateTime fixedDTP1 = dateTimePicker1;
            DateTime fixedDTP2 = dateTimePicker2;

            List<MongoNodeGas> mainList1 = database.GetCollection<MongoNodeGas>(date1).Find(
                x => x.dateTime > fixedDTP1 &
                x.dateTime < fixedDTP2)
                .ToList();


            for (int i = 1; i < 10; i++)
            {
                // фильтрую листы с лямбда выражениями
                List<MongoNodeGas> list1 = mainList1.Where(x => x.ID == i).ToList();
                int jh = 0;
                foreach (var j in list1)//цикл столбцов
                {
                    jh++;
                    switch (j.ID)
                    {
                        case 1:
                        case 2:
                        case 3:
                            break;

                        case 8:
                            ws1.Cells[jh + 1, 1] = j.dateTime.AddHours(3);
                            ws1.Cells[jh + 1, 2] = j.value;
                            break;
                        case 5:
                            ws1.Cells[jh + 1, 3] = j.IsWork;
                            break;
                        case 6:
                            ws1.Cells[jh + 1, 4] = j.IsWork;
                            break;
                        case 7:
                            ws1.Cells[jh + 1, 5] = j.IsWork;
                            break;
                    }
                    
                }

                
            }

            //дата для отчета
            string dateTime = DateTime.Now.ToShortDateString();

            string mainPath = Path.FirstPath();
            string xls = ".xls";
            string fileName = "GasIsWork__" + dateTime + xls;

            // скрыть лист
            //ws1.Visible = Excel.XlSheetVisibility.xlSheetHidden;

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
