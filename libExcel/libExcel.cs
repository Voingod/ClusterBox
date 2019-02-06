using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace libExcel
{
    class libExcel
    {
        static void Main(string[] args)
        {
            
            string path= @"D:\ForExcelTest.xlsx";
            string sheet= "Random";
            string value="One,Two";

            DataTable dt = Select(path,sheet,value);


















            //Для відлагодження
            string dataExcelGlobalFormat = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dataExcelGlobalFormat = Convert.ToString(dt.Rows[i][0]);
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
                if (File.Exists(path))
                {
                    string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
                    OleDbConnection conn = new OleDbConnection(stringcoon);
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    conn.Open();
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });


                    // Показать список листов в файле
                    List<string> ExcelSheets = new List<string>();
                    List<string> ColumnInSheets = new List<string>();


                    for (int i = 0; i < schemaTable.Rows.Count; i++)
                    {
                        string str = Convert.ToString(schemaTable.Rows[i].ItemArray[2]);
                        string[] charsToRemove = new string[] { "$" };
                        foreach (string c in charsToRemove)
                        {
                            str = str.Replace(c, string.Empty);
                        }
                        ExcelSheets.Add(str);
                    }

                    if (ExcelSheets.Contains(sheet))
                    {
                        int indexsheet = ExcelSheets.IndexOf(sheet);
                        string sheet1 = (string)schemaTable.Rows[indexsheet].ItemArray[2];
                        string select = String.Format("SELECT * FROM [{0}]", sheet1);

                        OleDbCommand oleDB = new OleDbCommand(select, conn);
                        OleDbDataReader reader = oleDB.ExecuteReader();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            ColumnInSheets.Add(reader.GetName(i)); // Имя столбца
                        }

                        if (ColumnInSheets.Contains(value))
                        {
                            OleDbDataAdapter da = new OleDbDataAdapter(" Select " + value + " from[" + sheet + "$]", conn);
                            da.Fill(dt);
                        }
                        else
                        {
                            MessageBox.Show("Таблиця (лист) " + sheet + " не містить стовпчика (колонки) " + value);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) +
                            " не містить таблиці (листа) " + sheet);
                    }
                    conn.Close();
                }
                else
                {
                    MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) + " не знайдено." +
                    "Перевірте наявність файлу " + path.Remove(0, path.IndexOf('\\') + 1) + ".\n" +
                    "Якщо файл існує, перевірте коректність введеного шляху: " +
                    path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\')) + "");
                }
            }
            catch (InvalidOperationException ex)
            {
                if (ex.HResult == -2146233079)
                {
                    MessageBox.Show("Необхідно встановити додаток AccessDatabaseEngine. \n" +
                        "(Для роботи з Excel файлами як файлами бази даних)");
                }
                else
                {
                    MessageBox.Show("Необроблена помилка!!!\n" + ex.Message);
                }

            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Необроблена помилка!!!\n" + ex.Message);

            }
            catch (FormatException ex)
            {
                MessageBox.Show("Необроблена помилка!!!\n"+ex.Message);
            }
            return dt;
        }


    }

}
