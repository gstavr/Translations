﻿using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;

namespace InsertTranslations {
    class Program {
        static void Main( string[] args ) {
            UnicodeEncoding uniencoding = new UnicodeEncoding();
            //string test1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string filepath = Path.GetFullPath("C:\\Users\\Mpoumpos\\Documents\\Visual Studio 2017\\Projects\\MCDA\\HelloWorld\\excel.xlsx");
            //string fullPath = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + "excel.xlsx");
            string fullPath = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
            //!Translation Folder
            string translationsPath =  Path.Combine(fullPath, "Translations");
            string GeneratedPath = Path.Combine(fullPath, "Generated");
            string SQLPath = Path.Combine(fullPath, "SQL");
            //! Check paths
            Directory.CreateDirectory( translationsPath );
            Directory.CreateDirectory( GeneratedPath );
            Directory.CreateDirectory( SQLPath );
            //! End
            bool valid = true;
            while ( valid ) {
                showFiles();
                ConsoleKeyInfo key = Console.ReadKey();
                Console.WriteLine();
                if ( checkKey( key ) > 0 && checkKey( key ) <= getFileEntries() ) {
                    CreateScriptFile( getSpecificFile( checkKey( key ) ) );
                }
            }
        }



        private static void CreateScriptFile( string filePath ) {
            ;

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            Encoding.RegisterProvider( CodePagesEncodingProvider.Instance );
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader1 = ExcelReaderFactory.CreateBinaryReader(stream);
            //...
            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();
            DataTable dt = result.Tables[0];
            StringBuilder sb = new StringBuilder();
            //! Remove First Row Cause its Header
            dt.Rows[0].Delete();
            dt.AcceptChanges();
            //! Saved path
            string SQLPath = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "SQL");

            using ( StreamWriter file = new StreamWriter( Path.Combine( SQLPath, string.Format( Path.GetFileName( filePath ).Split( '.' )[0] + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmm" ) + ".sql" ) ), false, new UTF8Encoding( false ) ) ) {
                foreach ( DataRow row in dt.Rows ) {
                    DataColumnCollection columns = dt.Columns;

                    for ( int i = 1; i < columns.Count; i++ ) {
                        sb.AppendFormat( "IF NOT EXISTS (select 1 from X_StaticTranslations_FactoryDefaults where [Cd] = '{0}' AND [Language] = {1}) \n", row[0].ToString(), i, row[i].ToString() );
                        sb.AppendLine( "BEGIN" );
                        sb.AppendFormat("\tINSERT INTO X_StaticTranslations_FactoryDefaults ([Cd], [Language], [TranslatedText], [Category]) VALUES ('{0}', {1}, N'{2}',2) \n", row[0].ToString(), i, row[i].ToString() );
                        sb.AppendLine( "END" );
                        sb.AppendLine( "GO" );
                    }
                }
                sb.AppendLine( "" );
                file.WriteLine( sb.ToString() );
            }
            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
            //! Move File to Generated Folder
            
            string GeneratedFolder = Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Generated");
            string destFile = System.IO.Path.Combine(GeneratedFolder, Path.GetFileName( filePath ));
            // Use Path class to manipulate file and directory paths.
            File.Copy( filePath, destFile, true );
            File.Delete( filePath );
        }

        /// <summary>
        /// Show Files in Translations Folder
        /// </summary>
        private static void showFiles() {
            string translationsPath =  Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations");
            // Process the list of files found in the directory.
            string [] fileEntries = Directory.GetFiles(translationsPath ,"*.xlsx");
            if ( fileEntries.Length > 0 ) {
                Console.WriteLine( "Files in Translations Folder" );
                int index = 1;
                foreach ( string fileName in fileEntries ) {
                    ProcessFile( fileName, index );
                    index++;
                }
                Console.WriteLine( "Choose a file to create SQL Script from {0} to {1} or Press any other key to Exit", 1, index - 1 );
            }
            else {
                Console.WriteLine( "No Files Found Application will Close in 5 sec" );
                Timer t = new Timer(Exit, null, 5000, 5000);
            }
        }

        private static int getFileEntries() {
            string [] fileEntries = Directory.GetFiles(Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations") ,"*.xlsx");

            return fileEntries.Length;
        }

        private static string getSpecificFile( int index ) {
            string [] fileEntries = Directory.GetFiles(Path.Combine(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory), "Translations") ,"*.xlsx");

            return fileEntries[index - 1];
        }

        // Insert logic for processing found files here.
        public static void ProcessFile( string path, int index ) {
            string fileName = Path.GetFileName( path );
            Console.WriteLine( "{0}. file: '{1}'.", index, fileName );
        }

        public static int checkKey( ConsoleKeyInfo key ) {
            int number = 0;
            // We check input for a Digit
            if ( char.IsDigit( key.KeyChar ) ) {
                number = int.Parse( key.KeyChar.ToString() );
            }
            else {
                Environment.Exit( 0 );
            }

            return number;
        }

        /// <summary>
        /// Exit Application
        /// </summary>
        /// <param name="state"></param>
        private static void Exit( object state ) {
            Environment.Exit( 0 );
        }
    }
}
