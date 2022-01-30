using System;
using Microsoft.Data.Sqlite;

namespace UsingSQLite
{
    public static class MyFunctions
    {
        static SqliteConnection _connection;
        static SqliteCommand _productNameCommand;

        private static void EnsureConnection()
        {
            if (_connection == null)
            {
                _connection = new SqliteConnection(@"Data Source=C:\Work\Excel-DNA\Samples\UsingSQLiteMDS\Northwind.db");
                _connection.Open();

                _productNameCommand = new SqliteCommand("SELECT ProductName FROM Products WHERE ProductID = @ProductID", _connection);
                _productNameCommand.Parameters.Add("@ProductID", SqliteType.Integer);
            }
        }

        public static object ProductName(int productID)
        {
            try
            {
                EnsureConnection();
                _productNameCommand.Parameters["@ProductID"].Value = productID;
                return _productNameCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

    }
}
