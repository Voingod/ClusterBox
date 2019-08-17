﻿using System;
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
           
            string path = @"D:\ForExcelTest1.xlsx";
            string sheet = "Random";
            string value = "One,Two";

            libExcel_Work lib = new libExcel_Work(path);
            DataTable dt = lib.Select(sheet, value);
            Console.ReadLine();
        }
    }
    class libExcel_Work
    {
        private string path;
        delegate DataTable GetSelect();
        public string Path
        {
            get
            {
                return path;
            }
            set
            {
                Exception ex = new Exception();
                if (File.Exists(path))
                {
                    path = value;
                }
                else
                {
                    path = null;
                }
            }  
        }

        public libExcel_Work()
        {

        }
        public libExcel_Work(string path)
        {
            //GetSelect getSelect;
            //getSelect = Select();
            Path =this.path = path;
            
        }
        /// <summary>
        /// Метод для считывания данных с Excel файлов. Принимает три параметра, возвращает один 
        /// в типе данных DataTable
        /// path - путь к таблице Excel, 
        /// sheet - лист в таблице, 
        /// value - название столбцов
        /// </summary>
        /// <param name="sheet">Лист для чтения</param>
        /// <param name="readColumn">Название столбцов для чтения</param>
        /// <returns></returns>
        /// <exception cref="System.InvalidOperationException">Thrown when...</exception>
        /// <exception cref="System.Data.OleDb.OleDbException">Thrown when...</exception>
        /// <exception cref="System.FormatException">Thrown when...</exception>
        /// <exception cref="System.IndexOutOfRangeException">Thrown when...</exception>
        /// 
        public DataTable Select (string sheet, string readColumn)
        {
          
            OleDbConnection conn = new OleDbConnection();
            //Знаки являются разделителями, в итоге получаем массив с именами, которые передаем в метод, разделенные этими знаками
            string [] list = readColumn.Split(new Char[] { ' ', ',', '.', ':', '_' }, StringSplitOptions.RemoveEmptyEntries);

            DataTable dt = new DataTable("Read");
            try
            {
                    Console.WriteLine(Path);
                    string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
                    conn = new OleDbConnection(stringcoon);
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    conn.Open();
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });


                    // Показать список листов в файле
                    List<string> ExcelSheets = new List<string>();
                    //Показать список столбцов в определенном листе
                    List<string> ColumnInSheets = new List<string>();


                    for (int i = 0; i < schemaTable.Rows.Count; i++)
                    {
                        //В переменную записываем название листа в том виде, в котором оно храниться в схеме Excel
                        string str = Convert.ToString(schemaTable.Rows[i].ItemArray[2]);
                        str = str.Replace("$", string.Empty);
                        ExcelSheets.Add(str);
                    }

                    if (ExcelSheets.Contains(sheet))
                    {
                        int indexsheet = ExcelSheets.IndexOf(sheet);
                        string sheet1 = (string)schemaTable.Rows[indexsheet].ItemArray[2];
                        string select = String.Format("SELECT * FROM [{0}]", sheet1);
                        bool flag = true;

                        OleDbCommand oleDB = new OleDbCommand(select, conn);
                        OleDbDataReader reader = oleDB.ExecuteReader();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            ColumnInSheets.Add(reader.GetName(i)); // Имя столбца
                        }
                        int j;
                        for (j = 0; j < list.Length; j++)
                        {
                            if (!ColumnInSheets.Contains(list[j]))
                            {
                                flag = false;
                                break;
                            }
                        }
                        if (flag)
                        {
                            OleDbDataAdapter da = new OleDbDataAdapter(" Select " + readColumn + " from[" + sheet + "$]", conn);
                            da.Fill(dt);
                        }
                        else
                            MessageBox.Show("Таблиця (лист) " + sheet + " не містить стовпчика (колонки) " + list[j]);
                    }
                    else
                        MessageBox.Show("Файл " + Path.Remove(0, Path.IndexOf('\\') + 1) +
                            " не містить таблиці (листа) " + sheet);
                //}
                //else
                //{
                //        MessageBox.Show("Файл " + Path.Remove(0, Path.IndexOf('\\') + 1) + " не знайдено." +
                //        "Перевірте наявність файлу " + Path.Remove(0, Path.IndexOf('\\') + 1) + ".\n" +
                //        "Якщо файл існує, перевірте коректність введеного шляху: " +
                //        Path.Remove(Path.LastIndexOf('\\'), Path.Length - Path.LastIndexOf('\\')) + "");
                //}
            }
            catch (InvalidOperationException ex)
            {
                if (ex.HResult == -2146233079)
                    MessageBox.Show("Необхідно встановити додаток AccessDatabaseEngine. \n" +
                        "(Для роботи з Excel файлами як файлами бази даних)");
                else
                     MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message);
            }
            catch (OleDbException ex){MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message+ex.StackTrace);}
            catch (FormatException ex){MessageBox.Show("Необроблена помилка!!!\n\t"+ex.Message+ex.StackTrace);}
            catch (IndexOutOfRangeException ex){MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message + ex.StackTrace);}
            finally{conn.Close();}

            return dt;
        }

        /// <summary>
        /// Метод для получения списка листов в таблице Excel
        /// </summary>
        /// <param name="path">Путь к таблице</param>
        /// <returns>Возвращает список с содержанием всех листов по указанному пути</returns>
        public List<string> ExcelSheet (string path)
        {
            List<string> ExcelSheets = new List<string>();
            OleDbConnection conn = new OleDbConnection();
            try
            {
                string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
                conn = new OleDbConnection(stringcoon);
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                conn.Open();
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                for (int i = 0; i < schemaTable.Rows.Count; i++)
                {
                    string str = Convert.ToString(schemaTable.Rows[i].ItemArray[2]);
                        str = str.Replace("$", string.Empty);
                    ExcelSheets.Add(str);
                }
                
            }
            catch (OleDbException)
            {
                MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) + " не знайдено." +
                "Перевірте наявність файлу " + path.Remove(0, path.IndexOf('\\') + 1) + ".\n" +
                "Якщо файл існує, перевірте коректність введеного шляху: " +
                path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\')) + "");
            }
            catch (InvalidOperationException ex)
            {
                if (ex.HResult == -2146233079)
                    MessageBox.Show("Необхідно встановити додаток AccessDatabaseEngine. \n" +
                        "(Для роботи з Excel файлами як файлами бази даних)");
                else
                    MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message);
            }
            finally { conn.Close(); };
            return ExcelSheets;
        }

        /// <summary>
        /// Метод для считывания всех столбцов из таблицы Excel по заданному пути и с заданного листа. 
        /// Внутри используется метод для считывания имен листов.
        /// </summary>
        /// <param name="path">Путь к таблице Excel</param>
        /// <param name="sheet">Лист, с которого считываются имена столбцов</param>
        /// <returns>Возвращает список с именами столбцов</returns>
        public List <string> ExcelSheetColumn(string path, string sheet)
        {
            List<string> ColumnInSheets = new List<string>();
            try
            {
                List<string> List = ExcelSheet(path);

                string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
                OleDbConnection conn = new OleDbConnection(stringcoon);
                OleDbCommand cmd = new OleDbCommand
                {
                    Connection = conn
                };
                conn.Open();
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
               
                string sheet1 = (string)schemaTable.Rows[List.IndexOf(sheet)].ItemArray[2];
                string select = String.Format("SELECT * FROM [{0}]", sheet1);

                OleDbCommand oleDB = new OleDbCommand(select, conn);
                OleDbDataReader reader = oleDB.ExecuteReader();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    ColumnInSheets.Add(reader.GetName(i)); // Имя столбца
                }
                conn.Close();
            }
            catch (OleDbException)
            {
                MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) + " не знайдено." +
                "Перевірте наявність файлу " + path.Remove(0, path.IndexOf('\\') + 1) + ".\n" +
                "Якщо файл існує, перевірте коректність введеного шляху: " +
                path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\')) + "");
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
                    MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message);
                }

            }
            return ColumnInSheets;
        }

        public void Instert(string path, string sheet, string columnName, string type)
        {
            type = "int,int,int";
            columnName = "One17,One174,One255";
            string[] list = columnName.Split(new Char[] { ' ', ',', '.', ':', '_' }, StringSplitOptions.RemoveEmptyEntries);
            string[] types = type.Split(new Char[] { ' ', ',', '.', ':', '_' }, StringSplitOptions.RemoveEmptyEntries);
            string[] listtypes = new string[list.Length];
            if (list.Length!=types.Length)
            {
                Console.WriteLine("gdfghdfh");
                return;
            }
            for(int i=0;i<list.Length;i++)
            {
                listtypes[i] = list[i] +" "+ types[i];
            }
            int[] count = new int[list.Length];
            
            List<string> sheets = new List<string>();
            sheets = ExcelSheetColumn(path, "Random");

            for(int i=0;i<list.Length;i++)
            {
                count[i] = sheets.IndexOf(list[i]);
            }

            string columns = "";
            for (int i = 0; i < sheets.Count; i++)
            {
                columnName = sheets[i];
                sheets.Remove(sheets[i]);
                sheets.Add(columnName+" DOUBLE");

            }
            for(int i=0;i<count.Length;i++)
            {
                sheets.RemoveAt(count[i]);
                sheets.Insert(count[i], listtypes[i]);
            }

            foreach(string col in sheets)
            {
                Console.WriteLine(col);
            }

            for(int i=0;i<sheets.Count;i++)
            {
                columns += sheets[i] == sheets[sheets.Count - 1] ? sheets[i] + ",": sheets[i];
            }
          //  Console.WriteLine(columns);


            //string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Mode = ReadWrite;" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            //OleDbConnection conn = new OleDbConnection(stringcoon);

            //conn.Open();
            //OleDbCommand cmd = new OleDbCommand();
            //cmd.Connection = conn;

            //cmd.CommandText = "CREATE TABLE [Random123$] ("+columns+");";
            //cmd.ExecuteNonQuery();

            //conn.Close();

            //conn.Open();
            //cmd.CommandText = "INSERT INTO ["+sheet+"$]("+columnName+") VALUES(3, 'CCCC','2014-01-03');";
            ////OleDbCommand commInsert = new OleDbCommand("Insert into  [" + sheet + "$](" + columnName + ") VALUES(@name)", conn);
            ////commInsert.Parameters.AddWithValue("@name", "NewName");
            //cmd.ExecuteNonQuery();
            //conn.Close();
        }
    }

}
