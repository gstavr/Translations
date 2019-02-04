using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;


namespace InsertTranslations
{
    class Program
    {

        static void Main(string[] args)
        {
            UnicodeEncoding uniencoding = new UnicodeEncoding();

            string fullPath = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
            //!Translation Folder
            string translationsPath = Path.Combine(fullPath, "Translations");
            string GeneratedPath = Path.Combine(fullPath, "Generated");
            string SQLPath = Path.Combine(fullPath, "SQL");
            //! Check paths
            Directory.CreateDirectory(translationsPath);
            Directory.CreateDirectory(GeneratedPath);
            Directory.CreateDirectory(SQLPath);
            //! End
            bool finish = true;
            while (finish)
            {
                showFiles();
                ConsoleKeyInfo key = Console.ReadKey();
                Console.WriteLine();
                if (checkKey(key) > 0 && checkKey(key) <= getFileEntries())
                {
                    CreateScriptFile(getSpecificFile(checkKey(key)));
                }
                else if (checkKey(key) == 0)
                {

                    foreach (string filepath in getAllFiles())
                    {
                        CreateScriptFile(filepath);
                    }
                    CreateSQLFile();
                    Console.WriteLine("Files Created Application will Close in 10 sec");
                    Timer t = new Timer(Exit, null, 10000, 10000);
                    finish = false;
                }
                else
                {
                    CreateSQLFile();
                }
            }
        }

        private static void CreateSQLFile()
        {
            string SQLPath = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "SQL");
            string sqlAllFilePath = Path.Combine(SQLPath, string.Format("SQLAll.sql"));
            if (File.Exists(sqlAllFilePath))
                File.Delete(sqlAllFilePath);

            string[] fileEntries = Directory.GetFiles(Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "SQL"), "*.sql");
            StringBuilder sqlScriptFinal = new StringBuilder();
            StringBuilder sqlScriptInsert = new StringBuilder();
            StringBuilder sqlScriptDynamicUpdate = new StringBuilder();
            StringBuilder sqlScriptDynamicInsert = new StringBuilder();
            StringBuilder stringLines = new StringBuilder();
            sqlScriptFinal.AppendLine("-------------------------- Custom Script ---------------------------");
            foreach (string file in fileEntries)
            {
                using (StreamReader sr = File.OpenText(file))
                {
                    string s = "";
                    while ((s = sr.ReadLine()) != null)
                    {
                        stringLines.AppendLine(s);
                    }
                    
                    if (Path.GetFileName(file).Contains("UpdateDynamic_") || Path.GetFileName(file).Contains("InsertDynamic_"))
                    {
                        if (Path.GetFileName(file).Contains("UpdateDynamic_"))
                            sqlScriptDynamicUpdate.Append(stringLines);
                        else
                            sqlScriptDynamicInsert.Append(stringLines);

                    }
                    else
                    {
                        sqlScriptInsert.Append(stringLines);
                    }
                }
                stringLines.Clear();
            }
            sqlScriptFinal.AppendLine("-------------------------- Add Custom Script Insert Scripts ---------------------------");
            sqlScriptFinal.Append(sqlScriptInsert);
            sqlScriptFinal.AppendLine("-------------------------- Add Dynamic Translations Scripts ---------------------------");
            sqlScriptFinal.Append(sqlScriptDynamicInsert);
            sqlScriptFinal.AppendLine("-------------------------- Add Dynamic Update Translations Scripts ---------------------------");
            sqlScriptFinal.Append(sqlScriptDynamicUpdate);
            sqlScriptFinal.AppendLine("PRINT('------------ Summary ------------')");
            sqlScriptFinal.AppendLine("GO");
            sqlScriptFinal.AppendLine("UPDATE st");
            sqlScriptFinal.AppendLine("SET st.TranslatedText = fd.TranslatedText");
            sqlScriptFinal.AppendLine("FROM X_StaticTranslations_FactoryDefaults fd");
            sqlScriptFinal.AppendLine("LEFT JOIN X_StaticTranslations st on fd.Cd = st.Cd and fd.[Language] = st.LanguageID");
            sqlScriptFinal.AppendLine("WHERE fd.TranslatedText <> st.TranslatedText and isnull(st.NoUpdate,0) = 0;");
            sqlScriptFinal.AppendLine("GO");
            sqlScriptFinal.AppendLine("\tINSERT INTO X_StaticTranslations (Cd, LanguageID, Category, TranslatedText)");
            sqlScriptFinal.AppendLine("SELECT fd.Cd, fd.Language, fd.Category,fd.TranslatedText");
            sqlScriptFinal.AppendLine("FROM X_StaticTranslations_FactoryDefaults fd");
            sqlScriptFinal.AppendLine("LEFT JOIN X_StaticTranslations st on fd.Cd = st.Cd and fd.[Language] = st.LanguageID");
            sqlScriptFinal.AppendLine("WHERE st.Cd is null;");
            sqlScriptFinal.AppendLine("GO");
            sqlScriptFinal.AppendLine("UPDATE o");
            sqlScriptFinal.AppendLine("SET o.VALUE = fd.VALUE");
            sqlScriptFinal.AppendLine("FROM L_Object_FactoryDefaults fd");
            sqlScriptFinal.AppendLine("LEFT JOIN L_Object o on o.ID_TABLE =fd.ID_TABLE and o.ID_LANGUAGES = fd.ID_LANGUAGES and o.TABLE_NAME = fd.TABLE_NAME");
            sqlScriptFinal.AppendLine("WHERE o.VALUE <> fd.VALUE and isnull(o.NO_UPDATE,0) = 0;");
            sqlScriptFinal.AppendLine("GO");
            sqlScriptFinal.AppendLine("\tINSERT INTO L_Object (ID_TABLE, ID_LANGUAGES, DATA_TYPE,VALUE,TABLE_NAME)");
            sqlScriptFinal.AppendLine("SELECT fd.ID_TABLE, fd.ID_LANGUAGES, fd.DATA_TYPE, fd.VALUE, fd.TABLE_NAME");
            sqlScriptFinal.AppendLine("FROM L_Object_FactoryDefaults fd");
            sqlScriptFinal.AppendLine("LEFT JOIN L_Object l on l.ID_TABLE = fd.ID_TABLE and l.ID_LANGUAGES = fd.ID_LANGUAGES and l.TABLE_NAME = fd.TABLE_NAME");
            sqlScriptFinal.AppendLine("WHERE l.ID_TABLE is null AND l.VALUE is null;");
            sqlScriptFinal.AppendLine("GO");
            sqlScriptFinal.AppendLine("PRINT 'Custom Script Completed'");



            using (StreamWriter files = new StreamWriter(Path.Combine(SQLPath, string.Format("SQLAll.sql")), false, new UTF8Encoding(false)))
            {
                files.Write(sqlScriptFinal);
                files.Close();
            }

        }

        private static void CreateScriptFile(string filePath)
        {

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader1 = ExcelReaderFactory.CreateBinaryReader(stream);
            //...
            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();
            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
            //! Move File to Generated Folder
            stream.Close();

            DataTable dt = result.Tables[0];

            //! Remove First Row Cause its Header
            dt.Rows[0].Delete();
            dt.AcceptChanges();

            if (Path.GetFileName(filePath).Contains("UpdateDynamic_") || Path.GetFileName(filePath).Contains("InsertDynamic_"))
            {
                SaveFile(filePath, dt, true);
            }
            else if (!CdExistsFunction(dt, filePath))
            {
                SaveFile(filePath, dt, false);
            }
        }


        private static void SaveFile(string filePath, DataTable dt, bool isDynamic)
        {
            StringBuilder sb = new StringBuilder();
            bool hasEmptyValues = false;
            foreach (DataRow row in dt.Rows)
            {
                DataColumnCollection columns = dt.Columns;

                if (isDynamic)
                {

                    if (Path.GetFileName(filePath).Contains("UpdateDynamic_"))
                    {

                        if (!string.IsNullOrWhiteSpace(row[2].ToString().Trim()) && !string.IsNullOrWhiteSpace(row[1].ToString().Trim()) && !hasEmptyValues)
                        {
                            //! Get Mode & Language
                            int updateMode = Convert.ToInt32(row[2].ToString().Trim());
                            int languageID = Convert.ToInt32(row[1].ToString().Trim());

                            // Check if setValue is Null or Empty 
                            if (!string.IsNullOrWhiteSpace(row[3].ToString().Trim())
                                && (!string.IsNullOrWhiteSpace(row[0].ToString().Trim()) // Check if Cd is Empty for Case 1 or 2
                                || (!string.IsNullOrWhiteSpace(row[4].ToString().Trim()) && !string.IsNullOrWhiteSpace(row[5].ToString().Trim())))) // check if TableName or ID_Table is empty for case 3
                            {
                                //! Update Modes Switch Case
                                switch (updateMode)
                                {
                                    // Update Static Translation
                                    case 1:
                                        sb.AppendFormat("UPDATE X_StaticTranslations_FactoryDefaults SET TranslatedText = N'{0}' WHERE CD = '{1}' AND [Language] = {2}", row[3].ToString().Trim(), row[0].ToString().Trim(), languageID);
                                        break;
                                    // Update Static Cd
                                    case 2:
                                        sb.AppendFormat("UPDATE X_StaticTranslations_FactoryDefaults SET [CD] = N'{0}' WHERE CD = '{1}' AND [Language] = {2}", row[3].ToString().Trim(), row[0].ToString().Trim(), languageID);
                                        break;
                                    // Update Dynamic Translation
                                    case 3:
                                        sb.AppendFormat("UPDATE L_Object_FactoryDefaults  SET [VALUE] = N'{0}' WHERE TABLE_NAME = '{1}' AND ID_TABLE = {2} AND ID_LANGUAGES = {3}", row[3].ToString().Trim(), row[4].ToString().Trim(), row[5].ToString().Trim(), languageID);
                                        break;
                                }
                                sb.AppendLine("");
                                sb.AppendLine("GO");
                            }
                            else
                            {
                                hasEmptyValues = true;
                            }
                        }
                        else
                        {
                            hasEmptyValues = true;
                        }

                    }
                    else
                    {
                        string TableName = row.IsNull(0) || string.IsNullOrWhiteSpace(row[0].ToString().Trim()) ? string.Empty : row[0].ToString().Trim();
                        string ID_Table = row.IsNull(1) || string.IsNullOrWhiteSpace(row[1].ToString().Trim()) ? string.Empty : row[1].ToString().Trim();
                        string Language = row.IsNull(2) || string.IsNullOrWhiteSpace(row[2].ToString().Trim()) ? string.Empty : row[2].ToString().Trim();
                        string Value = row.IsNull(3) || string.IsNullOrWhiteSpace(row[3].ToString().Trim()) ? string.Empty : row[3].ToString().Trim();
                        // Check if Cells Are Empty
                        if (!string.IsNullOrWhiteSpace(TableName) && !string.IsNullOrWhiteSpace(ID_Table) && !string.IsNullOrWhiteSpace(Language) && !string.IsNullOrWhiteSpace(Value) && !hasEmptyValues)
                        {
                            Value = Value.Contains("'") ? Value.Replace("'", "''") : Value;

                            sb.AppendFormat("IF NOT EXISTS (select 1 from L_Object_FactoryDefaults where [ID_TABLE] = {0} AND [ID_LANGUAGES] = {1} AND [VALUE] = '{2}') \n", Convert.ToInt32(ID_Table), Convert.ToInt32(Language), Value);
                            sb.AppendLine("BEGIN");
                            sb.AppendFormat("\tINSERT INTO [dbo].[L_Object_FactoryDefaults]([ID_TABLE],[ID_LANGUAGES],[DATA_TYPE],[VALUE],[TABLE_NAME]) VALUES ({0},{1},'TEXT','{2}','{3}')\n", Convert.ToInt32(ID_Table), Convert.ToInt32(Language), Value, TableName);
                            sb.AppendLine("END");
                            sb.AppendLine("GO");
                        }
                        else
                        {
                            hasEmptyValues = true;
                        }
                    }

                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(row[3].ToString().Trim()) && !row[3].ToString().Trim().Equals("..") && !hasEmptyValues)
                    {
                        for (int i = 1; i < 3; i++)
                        {
                            if (row[3].ToString().Length > 0)
                            {
                                sb.AppendFormat("IF NOT EXISTS (select 1 from X_StaticTranslations_FactoryDefaults where [Cd] = '{0}' AND [Language] = {1}) \n", row[3].ToString().Trim().Replace("'", "''"), i);
                                sb.AppendLine("BEGIN");
                                sb.AppendFormat("\tINSERT INTO X_StaticTranslations_FactoryDefaults ([Cd], [Language], [TranslatedText], [Category]) VALUES ('{0}', {1}, N'{2}',2) \n", row[3].ToString().Trim().Replace("'", "''"), i, row[i + 3].ToString().Trim().Replace("'", "''"));
                                sb.AppendLine("END");
                                sb.AppendLine("GO");
                            }
                        }
                    }
                    else
                    {
                        hasEmptyValues = true;
                    }
                }

            }
            sb.AppendLine("");

            if (!hasEmptyValues)
            {
                //! Saved path and Save to File
                string SQLPath = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "SQL");
                using (StreamWriter file = new StreamWriter(Path.Combine(SQLPath, string.Format(Path.GetFileName(filePath).Split('.')[0] + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".sql")), false, new UTF8Encoding(false)))
                {
                    file.WriteLine(sb.ToString());
                    file.Close();
                }
                string GeneratedFolder = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Generated");
                string destFile = System.IO.Path.Combine(GeneratedFolder, Path.GetFileName(filePath));
                // Use Path class to manipulate file and directory paths.
                File.Copy(filePath, destFile, true);
                File.Delete(filePath);
            }
            else
            {
                Console.WriteLine($"File ${Path.GetFileName(filePath)} has Empty or Null values");
            }

        }

        /// <summary>
        /// Show Files in Translations Folder
        /// </summary>
        private static void showFiles()
        {
            string translationsPath = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations");
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(translationsPath, "*.xlsx");
            if (fileEntries.Length > 0)
            {
                Console.WriteLine("Files in Translations Folder");
                int index = 1;
                foreach (string fileName in fileEntries)
                {
                    ProcessFile(fileName, index);
                    index++;
                }
                Console.WriteLine("Choose a file to create SQL Script from {0} to {1} or Press 0 for all files or Press any other key to Exit", 1, index - 1);
            }
            else
            {
                CreateSQLFile();
                Console.WriteLine("No Files Found Application will Close in 5 sec");
                Timer t = new Timer(Exit, null, 10000, 10000);
            }
        }

        private static int getFileEntries()
        {
            string[] fileEntries = Directory.GetFiles(Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations"), "*.xlsx");

            return fileEntries.Length;
        }

        private static string getSpecificFile(int index)
        {
            string[] fileEntries = Directory.GetFiles(Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations"), "*.xlsx");

            return fileEntries[index - 1];
        }

        private static string[] getAllFiles()
        {
            string[] fileEntries = Directory.GetFiles(Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations"), "*.xlsx");

            return fileEntries;
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string path, int index)
        {
            string fileName = Path.GetFileName(path);
            Console.WriteLine("{0}. file: '{1}'.", index, fileName);
        }

        public static int checkKey(ConsoleKeyInfo key)
        {
            int number = 0;
            // We check input for a Digit
            if (char.IsDigit(key.KeyChar))
            {
                number = int.Parse(key.KeyChar.ToString());
            }
            else
            {
                Console.WriteLine("Error in Key!!!!!");
                Timer t = new Timer(Exit, null, 10000, 10000);
            }

            return number;
        }

        /// <summary>
        /// Exit Application
        /// </summary>
        /// <param name="state"></param>
        private static void Exit(object state)
        {
            Environment.Exit(0);
        }

        /// <summary>
        /// Check Cds if exists
        /// If at least one Cd exist dont create the file.
        /// And export Dublicate Cds
        /// </summary>
        private static bool CdExistsFunction(DataTable dataTable, string filePath)
        {
            bool exists = false;
            List<string> Cds = GetAllCDs(dataTable);
            DataTable dt = CheckCdsToDataBase(Cds, filePath);
            if (dt.Rows.Count > 0)
            {
                SaveDublicateEntriestoFile(dt, filePath);
                exists = true;
            }


            return exists;
        }

        /// <summary>
        /// Get All File Cds
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        private static List<string> GetAllCDs(DataTable dataTable)
        {
            List<string> Cds = new List<string>();

            foreach (DataRow row in dataTable.Rows)
            {
                if (!string.IsNullOrWhiteSpace(row[3].ToString()))
                    Cds.Add(row[3].ToString());
            }

            return Cds;
        }

        private static DataTable CheckCdsToDataBase(List<string> Cds, string filePath)
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection("Data Source=dev2\\epsilon8; Initial Catalog = ess_dev; User Id = sa; Password = epsilonsa;"))
            {
                con.ConnectionString = "Data Source=dev2\\epsilon8; Initial Catalog = ess_dev; User Id = sa; Password = epsilonsa;";
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                string test = string.Format("'{0}'", string.Join(",", Cds.ToArray<string>()).Replace(",", "','"));
                cmd.CommandText = $"SELECT cd FROM X_StaticTranslations_FactoryDefaults WHERE Cd in({test}) GROUP by cd";
                SqlDataReader reader = cmd.ExecuteReader();
                dt.Load(reader);
                con.Close();
            }

            return dt;
        }

        private static void SaveDublicateEntriestoFile(DataTable Cds, string filePath)
        {
            StringBuilder sb = new StringBuilder();
            //! Write Dublicate Entries To
            string SQLPath = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations");
            using (StreamWriter file = new StreamWriter(Path.Combine(SQLPath, string.Format(Path.GetFileName(filePath).Split('.')[0] + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".txt")), false, new UTF8Encoding(false)))
            {
                foreach (DataRow Cd in Cds.Rows)
                {
                    sb.AppendLine($"{Cd[0]}");
                }

                file.WriteLine(sb.ToString());

                file.Close();
            }

            Console.WriteLine($"File: { Path.GetFileName(filePath)} has Dublicate Entries and wasn't Created");
            Console.WriteLine("");
        }

    }
}
