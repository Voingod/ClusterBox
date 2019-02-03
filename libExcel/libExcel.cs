using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace libExcel
{
    class libExcel
    {
        static void Main(string[] args)
        {
            
            string path= @"D:\F2orExcelTest.xlsx";
            string sheet= "Random";
            string value="One";
            DataTable dt = Select(path,sheet,value);
            double dataExcelGlobalFormat = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dataExcelGlobalFormat = Convert.ToDouble(dt.Rows[i][0].ToString());
                Console.WriteLine(dataExcelGlobalFormat);
            }
            Console.ReadLine();
        }

        /// <summary>
        /// Метод для считывания данных с Excel файлов
        /// </summary>
        /// <param name="path">Пусть к файлу</param>
        /// <param name="sheet">Лист для чтения</param>
        /// <param name="value">Название столбцов для чтения</param>
        /// <returns></returns>
        static DataTable Select (string path, string sheet, string value)
        {
            DataTable dt = new DataTable("Read");
            try
            {
                string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
                OleDbConnection conn = new OleDbConnection(stringcoon);
                OleDbDataAdapter da = new OleDbDataAdapter("Select " + value + " from[" + sheet + "$]", conn);
                
                da.Fill(dt);
                
            }
            catch (OleDbException ex)
            {
                int i = 1;
                int j = 2;

                 var abc = ex.ErrorCode == -2147467259 ? i=3 : j=4;

    //            Console.WriteLine("Файл {0} не знайдено. Перевірте наявність файлу {0}.\n" +
    //"Якщо файл існує, перевірте коректність введеного шляху", path);

                MessageBox.Show("Файл "+ path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\'))
                +" не знайдено. Перевірте наявність файлу " +path+".\n" +
                    "Якщо файл існує, перевірте коректність введеного шляху");
                

            }
            return dt;
        }


    }

}
