using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using ADODB;
using System.IO;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using DataStreams.ETL;

namespace ExcelUtilityLibrary
{
    public class ExcelUtility : IDisposable
    {
        /// <summary>
        /// Constant for indicating a text file
        /// </summary>
        public const short TXT = 0;
        /// <summary>
        /// Constant for indicating a csv file
        /// </summary>
        public const short CSV = 1;
        /// <summary>
        /// Constant for indicating a xls file
        /// </summary>
        public const short XLS = 2;
        /// <summary>
        /// Constant for indicating a xlsx file
        /// </summary>
        public const short XLSX = 3;
        /// <summary>
        /// Constant for indicating a pipe (|) separated text file
        /// </summary>
        public const short PIPE = 4;

        /// <summary>
        /// User for connecting to database
        /// </summary>
        protected string dbUser;
        /// <summary>
        /// Password to connect to db
        /// </summary>
        protected string dbPwd;
        /// <summary>
        /// Server to connect to
        /// </summary>
        protected string server;
        /// <summary>
        /// Catalog (database to connect to)
        /// </summary>
        protected string dbCatalog;
        /// <summary>
        /// Object to connect to the database
        /// </summary>
        protected SqlConnection dbConnection;

        /// <summary>
        /// Excel application used for creating files
        /// </summary>
        protected Excel.Application excelApp;

        /// <summary>
        /// ADODB.Connection used for report generation
        /// </summary>
        protected ADODB.Connection adodbConnection;
        /// <summary>
        /// Indicates if the user is connecting through SSID
        /// </summary>
        protected bool integratedSecurity;
        /// <summary>
        /// Empty library constructor
        /// </summary>
        public ExcelUtility() {

        }


        /// <summary>
        /// Establishes a connection to a database
        /// </summary>
        /// <param name="srv">Server to connect to database</param>
        /// <param name="dbCat">Database (catalog) to connect to</param>
        /// <param name="user">User to connect to database</param>
        /// <param name="pwd">Password to connect to db</param>
        public virtual void connectToDB(string srv, string dbCat, string user, string pwd)
        {

            if (user == null || pwd == null)
                connectToDB(srv, dbCat);
            else {
                #region Building the connection string

                server = srv;
                dbUser = user;
                dbPwd = pwd;
                dbCatalog = dbCat;
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = server;
                builder.InitialCatalog = dbCatalog;
                builder.UserID = dbUser;
                builder.Password = dbPwd;
                builder.MultipleActiveResultSets = true;
                #endregion
                integratedSecurity = false;
                #region Connecting to database
                try
                {
                    if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
                    {
                        terminateDBConnection();
                        dbConnection.ConnectionString = builder.ConnectionString;
                    }
                    else
                    {
                        dbConnection = new SqlConnection(builder.ConnectionString);

                    }
                    dbConnection.Open();
                    connectToADODB();
                }
                catch (Exception e)
                {
                    dbConnection = null;
                    throw new Exception("Error connecting to database: " + e.Message);
                }
                #endregion
            }
            
            
        }

        /// <summary>
        /// Establishes a connection to a database using integrated security
        /// </summary>
        /// <param name="srv">Server to connect to database</param>
        /// <param name="dbCat">Database (catalog) to connect to</param>
        /// <param name="user">User to connect to database</param>
        /// <param name="pwd">Password to connect to db</param>
        public virtual void connectToDB(string srv, string dbCat)
        {

            #region Building the connection string

            server = srv;
            dbUser = null;
            dbPwd = null;
            dbCatalog = dbCat;
            integratedSecurity = true;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = server;
            builder.InitialCatalog = dbCatalog;
            builder.IntegratedSecurity = true;
            builder.MultipleActiveResultSets = true;
            #endregion
            #region Connecting to database
            try
            {
                if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
                {
                    terminateDBConnection();
                    dbConnection.ConnectionString = builder.ConnectionString;
                }
                else
                {
                    dbConnection = new SqlConnection(builder.ConnectionString);

                }
                dbConnection.Open();
                connectToADODB();
            }
            catch (Exception e)
            {
                dbConnection = null;
                throw new Exception("Error connecting to database: " + e.Message);
            }
            #endregion
        }

        /// <summary>
        /// Establishes an ADODB Connection for quick file generation and reporting
        /// <remarks>This is a bug. There should be only one connection to the database, but because of slow performance for excel generation from SQLDataReader this was used</remarks>
        /// </summary>
        protected virtual void connectToADODB() {
            #region Building the connection string
            string connection = null;
            if(integratedSecurity)
                connection = String.Format("PROVIDER=SQLOLEDB;DATA SOURCE={0};INITIAL CATALOG={1};Integrated Security=SSPI", server, dbCatalog);
            else
                connection = String.Format("PROVIDER=SQLOLEDB;DATA SOURCE={0};INITIAL CATALOG={1};User ID={2};Password={3}", server, dbCatalog, dbUser, dbPwd);
            
            #endregion
            #region Connecting to database
            try
            {
                if (adodbConnection != null && adodbConnection.State != 0)
                {
                    terminateADODBConnection();
                }
                adodbConnection = new ADODB.Connection();
                adodbConnection.ConnectionString = connection;
                adodbConnection.ConnectionTimeout = 5000;
                adodbConnection.CommandTimeout = 50000;
                adodbConnection.Mode = ConnectModeEnum.adModeShareDenyNone;
                adodbConnection.Open();
            }
            catch (Exception e)
            {
                adodbConnection = null;
                throw new Exception("Error connecting to database: " + e.Message);
            }
            #endregion
        }

        /// <summary>
        /// Ends the established ADODB Connection
        /// </summary>
        protected virtual void terminateADODBConnection() {
            try
            {
                if (adodbConnection != null)
                    adodbConnection.Close();
                adodbConnection = null;
            }
            catch (Exception e) { 
            
            }
            
        }
        /// <summary>
        /// Executes a query in the ADODB:Connection object and returns the result set
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <returns>recordset with the data from the query</returns>
        protected virtual Recordset executeADODBQuery(string query) {
            if (adodbConnection != null && adodbConnection.State != 0)
            {
                ADODB.Recordset recordset = new ADODB.Recordset();
                recordset.CursorType = CursorTypeEnum.adOpenStatic;
                recordset.Open(query, adodbConnection);
                return recordset;
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a Query using an existing ADODB.Connection
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <param name="conn">ADODB Connection to be used</param>
        /// <returns>Recordset with the data from the query</returns>
        protected ADODB.Recordset executeADODBQuery(string query, ADODB.Connection conn)
        {
            if (conn != null && conn.State != 0)
            {
                ADODB.Recordset recordset = new ADODB.Recordset();
                recordset.Open(query, conn);
                if (recordset.State == 0)
                {
                    recordset.Close();
                    recordset = null;
                    throw new Exception("Query returns no records. Please verify your query");
                }
                return recordset;
            }
            else
                throw new Exception("Invalid connection");
        }

        /// <summary>
        /// Executes a Stored Procedure which has a return value
        /// </summary>
        /// <param name="comm">Command to be executed</param>
        /// <returns></returns>
        protected virtual int executeProcedure(String comm) {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                SqlCommand command = new SqlCommand(comm, dbConnection);
                command.CommandTimeout = 2000;
                SqlParameter outp = new SqlParameter("@return", SqlDbType.Int);
                outp.Direction = ParameterDirection.ReturnValue;
                command.Parameters.Add(outp);
                command.CommandType = CommandType.StoredProcedure;
                command.ExecuteNonQuery();
                int returnv = Int32 .Parse (command.Parameters["@return"].Value.ToString());
                return returnv;
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a Stored Procedure which has a return value
        /// </summary>
        /// <param name="comm">Command to be executed</param>
        /// <returns></returns>
        protected virtual int executeProcedure(String comm, params SqlParameter[] values)
        {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(comm, dbConnection)) {
                    command.CommandTimeout = 2000;
                    SqlParameter outp = new SqlParameter("@return", SqlDbType.Int);
                    outp.Direction = ParameterDirection.ReturnValue;
                    command.Parameters.AddRange(values);
                    command.Parameters.Add(outp);
                    command.CommandType = CommandType.StoredProcedure;
                    command.ExecuteNonQuery();
                    int returnv = Int32.Parse(command.Parameters["@return"].Value.ToString());
                    return returnv;
                }
                
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes an stored procedure using a particular SQL Connection and SQL parameters
        /// </summary>
        /// <param name="comm">Command string do be executed</param>
        /// <param name="conn">SQL Connection to be used</param>
        /// <param name="values">SQLParameters for the command</param>
        /// <returns>Return value from stored procedure</returns>
        protected virtual int executeProcedure(string comm, SqlConnection conn, params SqlParameter[] values) {
            if (conn != null && conn.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(comm, conn))
                {
                    command.CommandTimeout = 2000;
                    SqlParameter outp = new SqlParameter("@return", SqlDbType.Int);
                    outp.Direction = ParameterDirection.ReturnValue;
                    command.Parameters.AddRange(values);
                    command.Parameters.Add(outp);
                    command.CommandType = CommandType.StoredProcedure;
                    command.ExecuteNonQuery();
                    int returnv = Int32.Parse(command.Parameters["@return"].Value.ToString());
                    return returnv;
                }

            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a SQL command (Stored procedure) without validating result values
        /// </summary>
        /// <param name="comm">Command to be executed</param>
        public virtual void executeCommand(String comm) {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                SqlCommand command = new SqlCommand(comm, dbConnection);
                command.CommandTimeout = 2000;
                command.ExecuteNonQuery();
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// executes an SQL command using parameter values for 
        /// </summary>
        /// <param name="comm"></param>
        /// <param name="values"></param>
        public virtual void executeCommand(String comm, params SqlParameter[] values) {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(comm, dbConnection)) {
                    command.CommandTimeout = 2000;
                    //for (int i = 0; i < commandParameters.Length; i++)
                    command.Parameters.AddRange(values);
                    command.ExecuteNonQuery();
                }                
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }
        /// <summary>
        /// Executes a SQL query 
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <returns>A SqlDataReader with the rows resulting from the Query. null if there is no connection to a database</returns>
        public virtual SqlDataReader executeQuery(string query){
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed){
                using (SqlCommand command = new SqlCommand(query, dbConnection)) {
                    command.CommandTimeout = 2000;
                    return command.ExecuteReader();
                }
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a Query using the parameters provided by the user and the standard sql connection
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <param name="values">List of parameters to be added to the query</param>
        /// <returns>Datareader with query results</returns>
        public virtual SqlDataReader executeQuery(string query, params SqlParameter[] values)
        {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(query, dbConnection)) {
                    command.Parameters.AddRange(values);
                    command.CommandTimeout = 2000;
                    return command.ExecuteReader();
                }
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a Query using the parameters provided by the user and a particular sql connection
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <param name="conn">SQLConnection object to retrieve data from</param>
        /// <param name="values">List of parameters to be added to the query</param>
        /// <returns>Datareader with query results</returns>
        public virtual SqlDataReader executeQuery(string query, SqlConnection conn,params SqlParameter[] values)
        {
            if (conn != null && conn.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddRange(values);
                    command.CommandTimeout = 2000;
                    return command.ExecuteReader();
                }
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a query and returns the first value requested. For use with scalar functions or values
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <returns>Object returned by query</returns>
        public virtual object executeScalar(string query) {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(query, dbConnection))
                {
                    command.CommandTimeout = 2000;
                    return command.ExecuteScalar();
                }
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a query and returns the first value requested. For use with scalar functions or values
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <param name="values">List of parameters to be added to the query</param>
        /// <returns>Scalar returned by query</returns>
        public virtual object executeScalar(string query, params SqlParameter[] values)
        {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(query, dbConnection))
                {
                    command.Parameters.AddRange(values);
                    command.CommandTimeout = 2000;
                    return command.ExecuteScalar();
                }
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Executes a query and returns the first value requested. For use with scalar functions or values.
        /// </summary>
        /// <param name="query">Query to be executed</param>
        /// <param name="conn">SQLConnection object to retrieve data from</param>
        /// <param name="values">List of parameters to be added to the query</param>
        /// <returns>Scalar returned by query</returns>
        public virtual object executeScalar(string query, SqlConnection conn, params SqlParameter[] values)
        {
            if (conn != null && conn.State != ConnectionState.Closed)
            {
                using (SqlCommand command = new SqlCommand(query, conn))
                {
                    command.Parameters.AddRange(values);
                    command.CommandTimeout = 2000;
                    return command.ExecuteScalar();
                }
            }
            else
                throw new Exception("You are not connected to a database. Please connect to a valid database");
        }

        /// <summary>
        /// Opens an exiting excel file
        /// </summary>
        /// <param name="path">Path to file to be opened</param>
        public virtual void openExcelFile(string path) {
            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Workbooks.Open(path);
            }
            catch (Exception e) {
                closeExcelFile();
                throw e;
            }
        }

        /// <summary>
        /// Creates an new excel file with the specified extension
        /// </summary>
        /// <param name="workbook">Excel.Workbook to be opened</param>
        /// <param name="path">File path to file to be created</param>
        /// <param name="filetype">File extension of file to be created</param>
        public virtual void createExcelFile(ref Excel.Workbook workbook, string path, short filetype) {
            switch (filetype)
            {
                case XLSX: workbook.SaveAs(path);
                    break;
                case XLS: workbook.SaveAs(path, 56);
                    break;
                case TXT: workbook.SaveAs(path, -4158);
                    break;
                case CSV: workbook.SaveAs(path, 6);
                    break;
                default: throw new Exception("File type not valid");
            }
        }

        /// <summary>
        /// Creates an new excel file with the specified extension
        /// </summary>
        /// <param name="path">File path to file to be created</param>
        /// <param name="filetype">File extension of file to be created</param>
        public virtual void createExcelFile(string path, short filetype) {
            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Workbooks.Add();
                //@todo Check current culture
                switch (filetype){
                    case XLSX:  excelApp.ActiveWorkbook.SaveAs(path);
                        break;
                    case XLS:   excelApp.ActiveWorkbook.SaveAs(path, 56);
                        break;
                    case TXT:   excelApp.ActiveWorkbook.SaveAs(path, -4158);
                        break;
                    case CSV:   excelApp.ActiveWorkbook.SaveAs(path, 6);
                        break;
                    default: throw new Exception("File type not valid");
                }
                
            }
            catch (Exception e)
            {
                closeExcelFile();
                throw e;
            }
        }

        /// <summary>
        /// Closes an open excel file
        /// </summary>
        public virtual void closeExcelFile()
        {

            if (excelApp != null)
            {
                Excel.Workbook active = excelApp.ActiveWorkbook;
                if (active != null)
                {
                    active.Close();
                    Marshal.FinalReleaseComObject(active);
                    active = null;
                }
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
                excelApp = null;
            }
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();

        }

        /// <summary>
        /// Closes an Excel Application 
        /// </summary>
        /// <param name="app">Application to be closed</param>
        protected void closeExcelFile(Excel.Application app)
        {
            if (app != null)
            {
                Excel.Workbook wb = app.ActiveWorkbook;
                if (wb != null)
                {
                    wb.Close();
                    Marshal.FinalReleaseComObject(wb);
                    wb = null;
                }

                app.Quit();
                Marshal.FinalReleaseComObject(app);

                app = null;
            }


        }

        /// <summary>
        /// Writes a query to a pipefile
        /// </summary>
        /// <param name="query">Query to be exported</param>
        /// <param name="path">Path to file. If exists is overwritten</param>
        public virtual void writeQueryToPipefile(string query, string path)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                SqlDataReader reader = executeQuery(query);
                string line = "";
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    line += reader.GetName(i) + "|";
                }
                writer.WriteLine(line.Substring(0, line.LastIndexOf("|")));
                writer.Flush();
                while (reader.Read())
                {
                    line = "";
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        line += reader[i] + "|";
                    }
                    writer.WriteLine(line.Substring(0, line.LastIndexOf("|")));
                    writer.Flush();
                }
                reader.Close();
                writer.Close();
            }
        }

        /// <summary>
        /// Writes a query to textfile
        /// </summary>
        /// <param name="query">Query to be exported</param>
        /// <param name="path">Path to file. If exists is overwritten</param>
        public virtual void writeQueryToTextFile(string query, string path, char separationCharacter = '\t', params SqlParameter[] queryParameters)
        {
            using (StreamWriter writer = new StreamWriter(path, false))
            {
                SqlDataReader reader = executeQuery(query, queryParameters);
                string line = "";
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    line += String.Format("{0}{1}",reader.GetName(i), separationCharacter);
                }
                writer.WriteLine(line.Substring(0, line.LastIndexOf(separationCharacter)));
                writer.Flush();
                while (reader.Read())
                {
                    line = "";
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        line += String.Format("{0}{1}", reader[i], separationCharacter);
                    }
                    writer.WriteLine(line.Substring(0, line.LastIndexOf(separationCharacter)));
                    writer.Flush();
                }
                reader.Close();
                writer.Close();
            }
        }

        /// <summary>
        /// Writes a query to an Excel file
        /// </summary>
        /// <param name="query">Query to be exported</param>
        /// <param name="repName">Name of excel sheet. If exists data is overwritten</param>
        public virtual void writeQueryToExcel(string query, string repName) {
            //@todo replace for dbConnection != null
            //SqlDataReader reader = executeQuery(query);

            if (dbConnection != null && dbConnection.State != ConnectionState.Closed )
            {
                if (adodbConnection == null || adodbConnection.State == 0)
                    connectToADODB();
                ADODB.Recordset recordset = executeADODBQuery(query);
                if(recordset.State == 0){
                    closeExcelFile();
                    throw new Exception("Object is closed. Please send a query returning rows");
                }
                if (excelApp != null)
                {
                    try
                    {
                        fillExcelWorksheet(recordset, repName);
                        recordset.Close();
                        recordset = null;
                    }
                    catch (Exception e)
                    {
                        //reader.Close();
                        recordset.Close();
                        recordset = null;
                        closeExcelFile();
                        throw e;
                    }
                }
                //else
                //    reader.Close();
            }
            else
                throw new Exception("There is no connection to a database");
            
        }

        /// <summary>
        /// Writes a query to an Excel file
        /// </summary>
        /// <param name="query">Query to be exported</param>
        /// <param name="repName">Name of excel sheet. If exists data is overwritten</param>
        /// <param name="workbook"></param>
        /// <param name="save"></param>
        public virtual void writeQueryToExcel(string query, string repName, Excel.Workbook workbook, bool save = true)
        {
            //@todo replace for dbConnection != null
            //SqlDataReader reader = executeQuery(query);

            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                if (adodbConnection == null || adodbConnection.State == 0)
                    connectToADODB();
                ADODB.Recordset recordset = executeADODBQuery(query);
                if (recordset.State == 0)
                {
                    closeExcelFile();
                    throw new Exception("Object is closed. Please send a query returning rows");
                }
                if (workbook != null)
                {
                    try
                    {
                        fillExcelWorksheet(recordset, repName, workbook, save);
                        recordset.Close();
                    }
                    catch (Exception e)
                    {
                        //reader.Close();
                        recordset.Close();
                        closeExcelFile();
                        throw e;
                    }
                }
                //else
                //    reader.Close();
            }
            else
                throw new Exception("There is no connection to a database");

        }

        /// <summary>
        /// Fills an Excel Worksheet with query results
        /// </summary>
        /// <param name="reader">ADODB recordset with query results </param>
        /// <param name="repName">Excel Worksheet Name</param>
        /// <param name="app">Excel Workbook to paste data into</param>
        /// <param name="save">Indicates if the file should be saved upon pasting the data</param>
        protected void fillExcelWorksheet(ADODB.Recordset reader, string repName, Excel.Workbook app, bool save = true)
        {
            bool existsRep = false;
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            sheets = app.Worksheets;
            foreach (Excel.Worksheet e in sheets)
                if (String.Equals(e.Name, repName))
                {
                    existsRep = true;
                    break;
                }
            if (!existsRep)
            {

                sheet = app.Worksheets.Add();
                sheet.Name = repName;
            }
            else
            {
                sheet = sheets[repName];
                sheet.Cells.Clear();
                sheet.Activate();
                //sheets[repName].Activate();
                //app.ActiveSheet.Cells.Clear();
            }
            foreach (Excel.Worksheet e in sheets)
                if (String.Equals(e.Name, "Sheet1") || String.Equals(e.Name, "Sheet2") || String.Equals(e.Name, "Sheet3"))
                {
                    //sheets[e.Name].Delete();
                    e.Delete();
                }


            excelFileFormatting(reader, sheet);

            if(save)
                app.Save();

            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(sheet);
            sheet = null;
            sheets = null;

        }

        /// <summary>
        /// Creates the sheet for pasting data.
        /// </summary>
        /// <param name="reader">Reader containing the data</param>
        /// <param name="repName">Name of excel sheet.</param>
        protected virtual void fillExcelWorksheet(ADODB.Recordset reader, string repName) {
            bool existsRep = false;
            Excel.Sheets sheets = excelApp.Worksheets;

            foreach (Excel.Worksheet e in sheets)
                if(String.Equals(e.Name, repName)){
                    existsRep = true;
                    break;
                }
            if (!existsRep)
            {
                sheets.Add();
                Excel.Worksheet active = excelApp.ActiveSheet;
                active.Name = repName;
            }
            else {
                
                Excel.Worksheet active = sheets[repName];
                Excel.Range cells = active.Cells;

                cells.Clear();
                active.Activate();
                cells.Clear();
            }
            foreach (Excel.Worksheet e in excelApp.Sheets)
                if (String.Equals(e.Name, "Sheet1") || String.Equals(e.Name, "Sheet2") || String.Equals(e.Name, "Sheet3"))
                    excelApp.Sheets[e.Name].Delete();

            excelFileFormatting(reader);

            Excel.Workbook wbook = excelApp.ActiveWorkbook;
            wbook.Save();

        }

        /// <summary>
        /// Formats the fields from a query and pastes the data
        /// </summary>
        /// <param name="reader">Reader containing the data</param>
        protected virtual void excelFileFormatting(ADODB.Recordset reader)
        {
            Excel.Worksheet sheet = excelApp.ActiveSheet;
            Excel.Range cells = sheet.Cells;
            for (int i = 0; i < reader.Fields.Count; i++) { 
                //Console.WriteLine(reader.GetName(i));
                //@todo should this snippet of code pass to another method? (cell formatting for name fields)
                Excel.Range cell = cells[1, i + 1];
                Excel.Interior interior = cell.Interior;
                Excel.Font font = cell.Font;
                interior.ColorIndex = Excel.Constants.xlSolid;
                interior.PatternColorIndex = Excel.Constants.xlAutomatic;
                interior.ThemeColor = 5;
                interior.TintAndShade = -0.249977111117893;
                interior.PatternTintAndShade = 0;
                font.Bold = true;
                font.TintAndShade = 0;
                font.ThemeColor = 1;

                //Field name 
                cell.Value = reader.Fields[i].Name;
            }
            if (reader.RecordCount > 0)
                reader.MoveFirst();

            int row = 2;
            Excel.Range copyCell = cells[row, 1];
            copyCell.CopyFromRecordset((ADODB.Recordset)reader);
            Excel.Range column = cells.EntireColumn;
            column.AutoFit();

            copyCell = cells[1, 1];
            copyCell.Select();

        }

        /// <summary>
        /// Formats an excel worksheet
        /// </summary>
        /// <param name="reader">Query results to be pasted into the worksheet</param>
        /// <param name="app">Excel Worksheet</param>
        protected void excelFileFormatting(ADODB.Recordset reader, Excel.Worksheet app)
        {
            for (int i = 0; i < reader.Fields.Count; i++)
            {
                //Console.WriteLine(reader.GetName(i));
                //@todo should this snippet of code pass to another method? (cell formatting for name fields)

                ////Field name 

                Excel.Range cell = app.Cells[1, i + 1];
                Excel.Interior interior = cell.Interior;
                interior.ColorIndex = Excel.Constants.xlSolid;
                interior.PatternColorIndex = Excel.Constants.xlAutomatic;
                interior.ThemeColor = 5;
                interior.TintAndShade = -0.249977111117893;
                interior.PatternTintAndShade = 0;

                Excel.Font font = cell.Font;
                font.Bold = true;
                font.TintAndShade = 0;
                font.ThemeColor = 1;

                //Field name 
                cell.Value = reader.Fields[i].Name;

                //Marshal.FinalReleaseComObject(cell);
                //cell = null;
                //Marshal.FinalReleaseComObject(interior);
                //interior = null;
                //Marshal.FinalReleaseComObject(font);
                //font = null;
            }
            if (reader.RecordCount > 0)
                reader.MoveFirst();

            int row = 2;
            Excel.Range copyCell = app.Cells[row, 1];
            copyCell.CopyFromRecordset((ADODB.Recordset)reader);
            Excel.Range sheetCells = app.Cells;
            Excel.Range column = sheetCells.EntireColumn;
            column.AutoFit();

            copyCell = app.Cells[1, 1];
            copyCell.Select();

            Marshal.FinalReleaseComObject(copyCell);
            copyCell = null;
            Marshal.FinalReleaseComObject(sheetCells);
            sheetCells = null;
            Marshal.FinalReleaseComObject(column);
            column = null;
        }


        /// <summary>
        /// Loads the contents of an excel file worksheet into a table on the database. For use with excel files
        /// </summary>
        /// <param name="file">Path to the file containing the data</param>
        /// <param name="table">Table on the database we want to load from</param>
        /// <param name="sheetName">Worksheet on the file we want to load from.</param>
        public virtual void bulkLoadFileIntoTable(string file, string table, string sheetName)
        {

            string oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + file + ";" +
                                               "Extended Properties=\"Excel 12.0;HDR=YES\";";
            using (OleDbConnection oledbConnection = new OleDbConnection(oledbConnectString))
            {
                string oledbQuery = "select * from [" + sheetName + "$]";
                oledbConnection.Open();
                OleDbCommand command = new OleDbCommand(oledbQuery, oledbConnection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    using (SqlBulkCopy bulk = new SqlBulkCopy(dbConnection))
                    {
                        bulk.DestinationTableName = table;
                        bulk.BulkCopyTimeout = 2000;

                        ValidatingDataReader val = new ValidatingDataReader(dr, dbConnection, bulk);
                            
                        bulk.WriteToServer(val);
                            //bulk.WriteToServer(dr);
                        
                    }
                }
            }

        }

        /// <summary>
        /// Bulk loads a file into a table using an specific column mapping. For use with Excel files.
        /// </summary>
        /// <param name="file">path to the file where data is stored</param>
        /// <param name="table">table where the data is going to be stored</param>
        /// <param name="sheetName">sheet name where the data exists</param>
        /// <param name="mapping">List of column mappings for the file and table</param>
        public virtual void bulkLoadFileIntoTable(string file, string table, string sheetName, List<SqlBulkCopyColumnMapping> mapping)
        {
            string oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + file + ";" +
                                               "Extended Properties=\"Excel 12.0;HDR=YES\";";
            using (OleDbConnection oledbConnection = new OleDbConnection(oledbConnectString))
            {
                string oledbQuery = "select * from [" + sheetName + "$]";
                oledbConnection.Open();
                OleDbCommand command = new OleDbCommand(oledbQuery, oledbConnection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    using (SqlBulkCopy bulk = new SqlBulkCopy(dbConnection))
                    {
                        bulk.ColumnMappings.Clear();
                        for (int i = 0; i < mapping.Count; i++)
                        {
                            Console.WriteLine(mapping[i].ToString());
                            bulk.ColumnMappings.Add(mapping[i]);
                        }
                        bulk.DestinationTableName = table;
                        bulk.BulkCopyTimeout = 2000;
                        
                        bulk.WriteToServer(dr);
                        

                    }
                }
            }
        }

        /// <summary>
        /// Bulk loads a file using an existing OleDbConnection. For use with Excel Files.
        /// </summary>
        /// <param name="oledbConnection">The connection to be used for loading. The method doesn't verify the validity of the connection, so the user should prepare the
        /// connection accordingly and catch possible exceptions</param>
        /// <param name="table">Destination table on the database. Should be a valid table.</param>
        /// <param name="source">Data source. If the source is a text file, should be just the file name, not the full path. 
        ///     If the source is an excel file, it should be the name of an existing sheet</param>
        /// <param name="query">Checks is the source parameter is a query. If it is, uses that query instead of an straight select * from [source$]</param>
        public virtual void bulkLoadFileIntoTable(OleDbConnection oledbConnection, string table, string source, bool query = false) {
            string oledbQuery = (query ? source : "select * from [" + source + "$]");
            OleDbCommand command = new OleDbCommand(oledbQuery, oledbConnection);
            using (OleDbDataReader dr = command.ExecuteReader())
            {
                using (SqlBulkCopy bulk = new SqlBulkCopy(dbConnection))
                {
                    bulk.DestinationTableName = table;
                    bulk.BulkCopyTimeout = 2000;

                    
                    ValidatingDataReader val = new ValidatingDataReader(dr, dbConnection, bulk);

                    bulk.WriteToServer(val);
                        //bulk.WriteToServer(dr);
                }
            }
        }

        /// <summary>
        /// Bulk loads a file using an existing OleDbConnection and an specified column mapping.
        /// </summary>
        /// <param name="oledbConnection">The connection to be used for loading. The method doesn't verify the validity of the connection, so the user should prepare the
        /// connection accordingly and catch possible exceptions</param>
        /// <param name="table">Destination table on the database. Should be a valid table.</param>
        /// <param name="source">Data source. If the source is a text file, should be just the file name, not the full path. 
        ///     If the source is an excel file, it should be the name of an existing sheet</param>
        /// <param name="mapping">List of column mappings for the file and table</param>
        public virtual void bulkLoadFileIntoTable(OleDbConnection oledbConnection, string table, string source, List<SqlBulkCopyColumnMapping> mapping) {
            string oledbQuery = "select * from [" + source + "$]";
            OleDbCommand command = new OleDbCommand(oledbQuery, oledbConnection);
            using (OleDbDataReader dr = command.ExecuteReader())
            {
                using (SqlBulkCopy bulk = new SqlBulkCopy(dbConnection))
                {
                    bulk.ColumnMappings.Clear();
                    for (int i = 0; i < mapping.Count; i++)
                    {
                        Console.WriteLine(mapping[i].ToString());
                        bulk.ColumnMappings.Add(mapping[i]);
                    }
                    bulk.DestinationTableName = table;
                    bulk.BulkCopyTimeout = 2000;
                    
                    bulk.WriteToServer(dr);
                    

                }
            }
        }

        /// <summary>
        /// Bulk loads a file into an specified table. For use with text files
        /// </summary>
        /// <param name="filePath">full path to the file. Folder is required for creating the connection string</param>
        /// <param name="table">Destination table on the database. Should be a valid table.</param>
        /// <param name="tabDelimited">Indicates if the file is tabDelimited.</param>
        /// <param name="hdr">Indicates if the file's first row has the column names. Default true</param>
        /// <param name="delimiter">character used for delimiting columns on the file. defaults to a pipe-delimited file (|).</param>
        public virtual void bulkLoadFileIntoTable(string filePath, string table, bool tabDelimited, bool hdr = true, char delimiter = '|') {
            
            StringBuilder builder = new StringBuilder();
            builder.AppendFormat("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", filePath.Substring(0, filePath.LastIndexOf("\\")));
            builder.AppendFormat(";Extended Properties='text;HDR={0};FMT={1}'", (hdr == true ? "Yes" : "No"), 
                (tabDelimited == true ? "TabDelimited" : String.Format("Delimited({0})", delimiter)));
            string oledbConnectString =  builder.ToString();

            using (OleDbConnection oledbConnection = new OleDbConnection(oledbConnectString))
            {
                string oledbQuery = "select * from [" + filePath.Substring(filePath.LastIndexOf("\\") + 1) + "$]";
                oledbConnection.Open();
                OleDbCommand command = new OleDbCommand(oledbQuery, oledbConnection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    using (SqlBulkCopy bulk = new SqlBulkCopy(dbConnection))
                    {
                        bulk.DestinationTableName = table;
                        bulk.BulkCopyTimeout = 2000;
                        bulk.WriteToServer(dr);
                       
                    }
                }
            }
        }

        /// <summary>
        /// terminates the connection to the database
        /// </summary>
        public virtual void terminateDBConnection()
        {
            if (dbConnection != null && dbConnection.State != ConnectionState.Closed)
            {
                try
                {
                    dbConnection.Close();
                }
                catch (Exception ex)
                { 
                
                }
                terminateADODBConnection();   
            }
        }
        /// <summary>
        /// Destructor for DatabaseDataLoader
        /// </summary>
        ~ExcelUtility()
        {
            try {
                terminateDBConnection();
                closeExcelFile();
            }
            finally{
            
            }
            
        }

        /// <summary>
        /// Disposes the application
        /// </summary>
        void IDisposable.Dispose()
        {
            try
            {
                terminateDBConnection();
                closeExcelFile();
            }
            finally
            {

            }
        }
    }
}
