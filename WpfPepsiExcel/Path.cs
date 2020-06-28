using System;
using System.Collections.Generic;
using System.Text;

namespace WpfPepsiExcel
{
   static class Path
    {
        internal static string FirstPath()
        {
            // путь для сохранения докумета
            string year = DateTime.Now.Year.ToString();
            String server = Environment.UserName;
            string newLocation = "C:\\Users\\" + server + "\\Document\\" + year + "\\Count Opr\\";
            bool exists = System.IO.Directory.Exists(newLocation);

            if (!exists)
            {
                System.IO.Directory.CreateDirectory(newLocation);
            }

            return newLocation;
        }
    }
}
