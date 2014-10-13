using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using Microsoft.Win32;

namespace FileSystemWatch
{
    static class AccessDatabaseLayer
    {
        /// <summary>
        /// The String to Access the Alias Database
        /// </summary>
        static public string ConnectionString
        {
            get { return "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=" + GetDatabasePath(); }
        }

        static public string GetDatabasePath()
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine.OpenSubKey(@"Software\HTE\Modular GPS\");
                return regKey.GetValue("Install_Path").ToString() + @"\Data\Identity.mdb";
            }
            catch
            {
                return @"C:\Program Files\HTE\Modular GPS\Data\Identity.mdb";
            }
        }

        /// <summary>
        /// Returns a DataTable containing the results of the sqlSearchString 
        /// </summary>
        /// <param name="sqlSearchString"></param>
        /// <returns></returns>
        static public DataTable ObtainDataTable(string sqlSearchString)
        {
            DataTable dataTable = null;

            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                OleDbDataAdapter sqlAdapter = new OleDbDataAdapter(sqlSearchString, connection);
                dataTable = new DataTable();
                sqlAdapter.Fill(dataTable);
            }

            return dataTable;
        }

        /// <summary>
        /// Returns a list of Alias Objects
        /// </summary>
        /// <param name="SqlSearchString"></param>
        /// <returns></returns>
        static public List<AliasData> BuildAlias(string SqlSearchString)
        {
            DataTable aliasTable = ObtainDataTable(SqlSearchString);
            List<AliasData> aliasList = new List<AliasData>();

            foreach (DataRow dataRow in aliasTable.Rows)
            {
                AliasData tmp = new AliasData(dataRow);
                aliasList.Add(tmp);
            }

            return aliasList;
        }

        /// <summary>
        /// Excecutes sqlScript
        /// </summary>
        /// <param name="sqlScript"></param>
        /// <returns></returns>
        static public void ExecuteScript(string sqlScript)
        {
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                OleDbCommand oleDbCommand = new OleDbCommand(sqlScript);

                connection.Open();
                oleDbCommand.Connection = connection;
                oleDbCommand.ExecuteNonQuery();
                connection.Close();
            }
        }
    }
}
