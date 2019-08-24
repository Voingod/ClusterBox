#define SelectWithParametr
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
           
            string path = @"D:\ForExcelTest2.xlsx";
            string sheet = "Random";
            string value = "One,Two";

            libExcel_Work lib = new libExcel_Work(path,sheet,value);
            var a=lib.ExcelSheet();
            var b = lib.ExcelSheetColumn(a[0]);

            Console.WriteLine();
            var dt = lib.Select();
            //foreach(var c in dt)
            //    Console.WriteLine(c);
           // Console.WriteLine(dt.Rows[0][1]);
            Console.ReadLine();

        }

    }
    class libExcel_Work
    {
        readonly string sheet;
        readonly string column;
        readonly string path;
        OleDbConnection conn = new OleDbConnection();

        /// <summary>
        /// Конструктор создает объект и инициализирует его, добавляя введенный путь.
        /// </summary>
        /// <param name="path">Путь к файлу Excel</param>
        public libExcel_Work(string path)
        {
            this.path = path;
            Connection();

        }

        /// <summary>
        /// Конструктор создает объект и инициализирует его, добавляя введенный путь, название листа Excel и название столбцов
        /// </summary>
        /// <param name="path">Путь к файлу Excel</param>
        /// <param name="sheet">Лист для чтения</param>
        /// <param name="column">Название столбцов для чтения</param>
        public libExcel_Work(string path, string sheet, string column)
        {
            this.path = path;
            this.sheet = sheet;
            this.column = column;
            Connection();

        }

        /// <summary>
        ///  Создает строку подключения для всех методов даного класса. Вызывается в конструкторе
        /// </summary>
        private void Connection()
        {
            string stringcoon = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
            conn = new OleDbConnection(stringcoon);
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
        }

        /// <summary>
        /// Метод для получения списка листов в таблице Excel
        /// </summary>
        /// <returns>Возвращает список с содержанием всех листов по указанному пути</returns>
        public List<string> ExcelSheet ()
        {
            List<string> ExcelSheets = new List<string>();
            if (File.Exists(path))
            {
                try
                {
                    conn.Open();
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    for (int i = 0; i < schemaTable.Rows.Count; i++)
                    {
                        //В переменную записываем название листа в том виде, в котором оно храниться в схеме Excel
                        string str = Convert.ToString(schemaTable.Rows[i].ItemArray[2]);
                        str = str.Replace("$", string.Empty);
                        ExcelSheets.Add(str);
                    }

                }
                catch (OleDbException ex){MessageBox.Show(ex.Message + ex.StackTrace);}
                catch (InvalidOperationException ex)
                {
                    if (ex.HResult == -2146233079)
                        MessageBox.Show("Необхідно встановити додаток AccessDatabaseEngine. \n" +
                            "(Для роботи з Excel файлами як файлами бази даних)");
                    else
                        MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message);
                }
                finally { conn.Close();};
            }
            else
            {
                MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) + " не знайдено." +
                "Перевірте наявність файлу " + path.Remove(0, path.IndexOf('\\') + 1) + ".\n" +
                "Якщо файл існує, перевірте коректність введеного шляху: " +
                path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\')) + "");

            }
            return ExcelSheets;
        }

        /// <summary>
        /// Метод для считывания всех столбцов из таблицы Excel по заданному пути и с заданного листа. 
        /// </summary>
        /// <param name="sheet">Лист, с которого считываются имена столбцов</param>
        /// <returns>Возвращает список с именами столбцов</returns>
        public List <string> ExcelSheetColumn(string sheet)
        {
            List<string> ColumnInSheets = new List<string>();
            if (File.Exists(path))
            {
                try
                {
                    conn.Open();
                    string select = String.Format($"SELECT * FROM [{sheet}$]");
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    OleDbCommand oleDB = new OleDbCommand(select, conn);
                    OleDbDataReader reader = oleDB.ExecuteReader();
                    for (int i = 0; i < reader.FieldCount; i++)
                        ColumnInSheets.Add(reader.GetName(i)); // Имя столбца

                }
                catch (OleDbException ex){MessageBox.Show(ex.Message + ex.StackTrace);}
                catch (InvalidOperationException ex)
                {
                    if (ex.HResult == -2146233079)
                        MessageBox.Show("Необхідно встановити додаток AccessDatabaseEngine. \n" +
                            "(Для роботи з Excel файлами як файлами бази даних)"+ex.StackTrace);
                    else
                        MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message);
                }
                finally { conn.Close(); };
            }
            else
            {
                MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) + " не знайдено." +
                "Перевірте наявність файлу " + path.Remove(0, path.IndexOf('\\') + 1) + ".\n" +
                "Якщо файл існує, перевірте коректність введеного шляху: " +
                path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\')) + "");
            }
            return ColumnInSheets;
        }

        /// <summary>
        /// Обобщеный метод для считывания данных с Excel файлов. Служит в качестве заполнителя для
        /// методов-оберток, принимающих разные типы данных. 
        /// </summary>
        /// <param name="sheet">Лист для чтения</param>
        /// <param name="readColumn">Название столбцов для чтения</param>
        /// <returns></returns>
        /// <exception cref="System.InvalidOperationException">Thrown when...</exception>
        /// <exception cref="System.Data.OleDb.OleDbException">Thrown when...</exception>
        /// <exception cref="System.FormatException">Thrown when...</exception>
        /// <exception cref="System.IndexOutOfRangeException">Thrown when...</exception>
        /// 
        private DataTable Select<T, V>(T sheet, V readColumn)
        {
            DataTable dt = new DataTable("Read");
            if (File.Exists(path))
            {
                try
                {
                    conn.Open();
                    OleDbDataAdapter da = new OleDbDataAdapter(" Select " + readColumn + " from[" + sheet + "$]", conn);
                    da.Fill(dt);
                }
                catch (InvalidOperationException ex)
                {
                    if (ex.HResult == -2146233079)
                        MessageBox.Show("Необхідно встановити додаток AccessDatabaseEngine. \n" +
                            "(Для роботи з Excel файлами як файлами бази даних)");
                    else
                        MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message + ex.StackTrace);
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message + ex.StackTrace);
                }
                catch (FormatException ex) { MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message + ex.StackTrace); }
                catch (IndexOutOfRangeException ex) { MessageBox.Show("Необроблена помилка!!!\n\t" + ex.Message + ex.StackTrace); }
                finally { conn.Close(); }
            }
            else
            {
                MessageBox.Show("Файл " + path.Remove(0, path.IndexOf('\\') + 1) + " не знайдено." +
                "Перевірте наявність файлу " + path.Remove(0, path.IndexOf('\\') + 1) + ".\n" +
                "Якщо файл існує, перевірте коректність введеного шляху: " +
                path.Remove(path.LastIndexOf('\\'), path.Length - path.LastIndexOf('\\')) + "");
            }
            return dt;
        }

        /// <summary>
        /// Метод для считывания данных с Excel файлов. Выбирает из заданной таблицы (листа Excel)
        /// значения заданных столбцов (названия столбцов на заданном листе Excel). При заполении
        /// readColumn в качестве строки, запятая "," выступает в качестве разделителя
        /// Инициализация пути к файлу Excel, происходит в конструкторе
        /// Первая строка в Excel файле должна выступать в роле названий столбца. 
        /// </summary>
        /// <param name="sheet">Лист для чтени</param>
        /// <param name="readColumn">Название столбцов для чтения</param>
        /// <returns></returns>
        public DataTable Select(string sheet, List<string> readColumn)
        {
            string column="";
            foreach (string col in readColumn)
                column += readColumn.Count == readColumn.IndexOf(col) + 1 ? col : col + ",";

            DataTable dt = Select<string,string>(sheet, column);
            return dt;
        }

        /// <summary>
        /// Метод для считывания данных с Excel файлов. Выбирает из заданной таблицы (листа Excel)
        /// значения заданных столбцов (названия столбцов на заданном листе Excel).
        /// Инициализация пути к файлу Excel, происходит в конструкторе
        /// Первая строка в Excel файле должна выступать в роле названий столбца. 
        /// </summary>
        /// <param name="sheet">Лист для чтени</param>
        /// <param name="readColumn">Название столбцов для чтения</param>
        /// <returns></returns>
        public DataTable Select(string sheet, string readColumn)
        {
            DataTable dt = Select<string,string>(sheet, readColumn);
            return dt;
        }

        /// <summary>
        /// Метод для считывания данных с Excel файлов. Инициализация пути к файлу Excel, 
        /// листа для чтения и названия столбцов происходит в конструкторе
        /// Первая строка в Excel файле должна выступать в роле названий столбца.  
        /// </summary>
        /// <returns></returns>
        public DataTable Select()
        {
            DataTable dt = Select<string,string>(sheet, column);
            return dt;
        }


        public void Instert(string sheet, string columnName, string type)
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
            sheets = ExcelSheetColumn("Random");

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
