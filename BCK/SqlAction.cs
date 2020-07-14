using System.Configuration;
using System.Data;
using System.Data.SqlClient;


namespace BCK
{
    class SqlAction
    {
        public static string LastError { get; private set; }

        static readonly SqlConnection sqlConn = new SqlConnection(
            ConfigurationManager
              .ConnectionStrings[
                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Name
                ].ConnectionString
            );

        public static SqlConnection GetSqlConnection()
        {
            return sqlConn;
        }

        public static bool OpenConnection()
        {
            if (sqlConn.State != ConnectionState.Open)
            {
                sqlConn.Close();
                try
                {
                    sqlConn.Open();
                }
                catch (SqlException ex)
                {
                    LastError = ex.Message;
                    return false;
                }
            }
            return true;
        }

        public static bool CloseConnection()
        {
            try
            {
                sqlConn.Close();
            }
            catch (SqlException ex)
            {
                LastError = ex.Message;
                return false;
            }
            return true;
        }

        public static SqlDataAdapter GetTable(string tableName)
        {
            if (OpenConnection())
            {
                try
                {
                    string strQuery = "SELECT * FROM " + tableName;
                    SqlDataAdapter daTable = new SqlDataAdapter(strQuery, sqlConn);
                    return daTable;
                }
                catch (SqlException ex)
                {
                    LastError = ex.Message;
                    return null as SqlDataAdapter;
                }
            }
            else
            {
                return null as SqlDataAdapter;
            }
        }

        public static SqlDataReader GetQuery(string strQuery)
        {
            if (OpenConnection())
            {
                try
                {
                    SqlDataReader drQuery = new SqlCommand(strQuery, sqlConn).ExecuteReader();
                    return drQuery;
                }
                catch (SqlException ex)
                {
                    LastError = ex.Message;
                    return null as SqlDataReader;
                }
            }
            else
            {
                return null as SqlDataReader;
            }
        }

        public static SqlDataReader GetQuery(string strQuery, params object[] objParameters)
        {
            if (OpenConnection())
            {
                try
                {
                    SqlCommand sqlCmd = new SqlCommand(strQuery, sqlConn);

                    for (int i = 0; i < objParameters.Length; i++)
                    {
                        sqlCmd.Parameters.Add(
                            new SqlParameter("@Parameter" + i.ToString(), objParameters[i])
                            );
                    }

                    SqlDataReader drQuery = sqlCmd.ExecuteReader();
                    return drQuery;
                }
                catch (SqlException ex)
                {
                    LastError = ex.Message;
                    return null as SqlDataReader;
                }
            }
            else
            {
                return null as SqlDataReader;
            }
        }

        public static bool ExecuteNonQuery(string strCmd, params object[] objParameters)
        {
            if (OpenConnection())
            {
                try
                {
                    SqlCommand sqlCmd = new SqlCommand(strCmd, sqlConn);

                    for (int i = 0; i < objParameters.Length; i++)
                    {
                        sqlCmd.Parameters.Add(
                            new SqlParameter("@Parameter" + i.ToString(), objParameters[i])
                            );
                    }

                    sqlCmd.ExecuteNonQuery();
                    CloseConnection();
                    return true;
                }
                catch (SqlException ex)
                {
                    LastError = ex.Message;
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
    }
}
