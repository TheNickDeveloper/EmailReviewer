using System.Data.SQLite;

namespace EmailReviewer.Services
{
    public class SqliteHelper
    {
        private readonly SQLiteConnection _dbConnection;
        private SQLiteCommand _dbCommand;
        private SQLiteDataReader _dataReader;

        public SqliteHelper(string connectionString)
        {
            _dbConnection = new SQLiteConnection(connectionString);
            _dbConnection.Open();
        }

        public SQLiteDataReader CreateTable(string tableName, string[] colNames, string[] colTypes)
        {
            string queryString = "CREATE TABLE IF NOT EXISTS " + tableName + "( " + colNames[0] + " " + colTypes[0];
            for (int i = 1; i < colNames.Length; i++)
            {
                queryString += ", " + colNames[i] + " " + colTypes[i];
            }
            queryString += "  ) ";
            return ExecuteQuery(queryString);
        }

        private SQLiteDataReader ExecuteQuery(string queryString)
        {
            _dbCommand = _dbConnection.CreateCommand();
            _dbCommand.CommandText = queryString;
            _dataReader = _dbCommand.ExecuteReader();

            return _dataReader;
        }
    }
}
