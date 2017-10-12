using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;

namespace EmailPaymentAdvice
{
    class ODBCClass : IDisposable
    {
        OdbcConnection oConnection;
        OdbcCommand oCommand;

        public ODBCClass(string DataSourceName)
        {
            //Instantiate the connection
            oConnection = new OdbcConnection("Dsn=" + DataSourceName);
            try
            {
                //Open the connection
                oConnection.Open();
                Console.WriteLine("The connection is established with the database");

            }
            catch (OdbcException caught)
            {
                Console.WriteLine(caught.Message);
                Console.Read();
            }
        }

        /// <summary>
        /// Execute SQL query
        /// </summary>
        /// <param name="Query">The SQL command to execute</param>
        /// <returns></returns>
        public OdbcCommand GetCommand(string Query)
        {
            oCommand = new OdbcCommand();
            oCommand.Connection = oConnection;
            oCommand.CommandText = Query;
            return oCommand;
        }

        /// <summary>
        /// Call stored procedure
        /// </summary>
        /// <param name="storedProcName">Procedure name including (?,?,...) for parameters.</param>
        /// <param name="parms">Array of KeyValuePairs, each containing param name and param value.</param>
        /// <returns></returns>
        public OdbcCommand GetCommand(string storedProcName, KeyValuePair<string, string>[] parms)
        {
            oCommand = new OdbcCommand("{call " + storedProcName + "}", oConnection);
            oCommand.CommandType = CommandType.StoredProcedure;
            foreach (var kvp in parms)
            {
                oCommand.Parameters.AddWithValue(kvp.Key, kvp.Value);
            }
            return oCommand;
        }



        public void Dispose()
        {
            oConnection.Close();
        }
    }
}
