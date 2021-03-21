using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Study_string.DataBase
{
    public class DataBaseProcess
    {
        private String myConnectionString;

        public DataBaseProcess(String connectionString)
        {
            this.myConnectionString = connectionString;
        }

        public String GetAppDate()
        {
            Object r = null;
            SqlCommand command;
            SqlConnection connection;
            DateTime serverDateTime;
            String rVal = "";

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = "Select getDate()";
                r = command.ExecuteScalar();
                if (r != null)
                {
                    serverDateTime = Convert.ToDateTime(r);

                    if (serverDateTime.Hour >= 9) //작업일자는 마감시간 9시 기준임.
                    {
                        rVal = serverDateTime.ToString("yyyyMMdd");
                    }
                    else
                    {
                        rVal = serverDateTime.AddDays(-1).ToString("yyyyMMdd");
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return rVal;
        }

        private String GetMd5Hash(String input)
        {
            // Create a new instance of the MD5CryptoServiceProvider object.
            MD5 md5Hasher = MD5.Create();

            // Convert the input string to a byte array and compute the hash.
            byte[] data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input));

            // Create a new Stringbuilder to collect the bytes
            // and create a string.
            StringBuilder sBuilder = new StringBuilder();

            // Loop through each byte of the hashed data 
            // and format each one as a hexadecimal string.
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }

            // Return the hexadecimal string.
            return sBuilder.ToString();
        }

        public DataRow GetDataRow_Query(String sql)
        {
            DataTable table;
            DataRow row;
            DataSet ds;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter ad;

            table = new DataTable();
            row = null;
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                    if (table.Rows.Count > 0)
                    {
                        row = table.Rows[0];
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return row;
        }

        public DataRow GetDataRow_Query_Parameter(String sql, SqlCommand executeCommand)
        {
            DataTable table;
            DataRow row;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;

            table = new DataTable();
            row = null;
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.Text;
                executeCommand.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                    if (table.Rows.Count > 0)
                    {
                        row = table.Rows[0];
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return row;
        }

        public DataRow GetDataRow_Query_Parameter(String sql, Dictionary<String, Object> parameters)
        {
            DataTable table;
            DataRow row;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;
            SqlCommand command;

            table = new DataTable();
            row = null;
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                    if (table.Rows.Count > 0)
                    {
                        row = table.Rows[0];
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return row;
        }

        public DataRow GetDataRow_SP(String sp, SqlCommand executeCommand)
        {
            DataTable table;
            DataRow row;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;

            table = new DataTable();
            row = null;
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                    if (table.Rows.Count > 0)
                    {
                        row = table.Rows[0];
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return row;
        }

        public DataRow GetDataRow_SP(String sp, Dictionary<String, Object> parameters)
        {
            DataTable table;
            DataRow row;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;
            SqlCommand command;

            table = new DataTable();
            row = null;
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = sp;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                    if (table.Rows.Count > 0)
                    {
                        row = table.Rows[0];
                    }
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return row;
        }

        public DataTable GetDataTable_Query(String sql)
        {
            DataTable table;
            DataSet ds;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter ad;

            table = new DataTable();
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return table;
        }

        public DataTable GetDataTable_Query_Parameter(String sql, SqlCommand executeCommand)
        {
            DataTable table;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;

            table = new DataTable();
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.Text;
                executeCommand.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return table;
        }

        public DataTable GetDataTable_Query_Parameter(String sql, Dictionary<String, Object> parameters)
        {
            DataTable table;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;
            SqlCommand command;

            table = new DataTable();
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return table;
        }

        public DataTable GetDataTable_SP(String sp, SqlCommand executeCommand)
        {
            DataTable table;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;

            table = new DataTable();
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return table;
        }

        public SqlCommand GetStringList_SP(String sp, SqlCommand executeCommand)
        {
            DataTable table;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;
            List<String> list = new List<String>();
            table = new DataTable();
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;
                executeCommand.ExecuteNonQuery();
                //ds = new DataSet();
                //ad = new SqlDataAdapter();
                //ad.SelectCommand = executeCommand;
                //ad.Fill(ds);
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return executeCommand;
        }

        //public SqlCommand GetStringList_SP2(String sp, SqlCommand executeCommand)
        public DataSet GetStringList_SP2(String sp, SqlCommand executeCommand)
        {
            DataTable table;
            DataSet ds = new DataSet(); ;
            SqlDataAdapter ad;
            SqlConnection connection;
            List<String> list = new List<String>();
            table = new DataTable();
            connection = null;
            //Logger logger = new Common.Logger();

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;
                //executeCommand.ExecuteNonQuery();
                //ds = new DataSet();
                //ad = new SqlDataAdapter();
                //ad.SelectCommand = executeCommand;
                //ad.Fill(ds);
                //executeCommand.ExecuteNonQuery();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);
                if (connection != null)
                {
                    connection.Close();
                }
                return ds;
            }
            catch (Exception e)
            {
                //logger.WriteLog("Execption: [ " + e.Message + " ]"); // log 기록
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            //return executeCommand;
            return ds;
        }

        public DataTable GetDataTable_SP(String sp, Dictionary<String, Object> parameters)
        {
            DataTable table;
            DataSet ds;
            SqlDataAdapter ad;
            SqlConnection connection;
            SqlCommand command;

            table = new DataTable();
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = sp;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);

                if (ds.Tables.Count > 0)
                {
                    table = ds.Tables[0];
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return table;
        }

        public DataSet GetDataSet_Query(String sql)
        {
            DataSet ds = new DataSet();
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter ad;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return ds;
        }

        public DataSet GetDataSet_Query_Parameter(String sql, SqlCommand executeCommand)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter ad;
            SqlConnection connection;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.Text;
                executeCommand.CommandText = sql;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return ds;
        }

        public DataSet GetDataSet_Query_Parameter(String sql, Dictionary<String, Object> parameters)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter ad;
            SqlConnection connection;
            SqlCommand command;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 120;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return ds;
        }

        public DataSet GetDataSet_SP(String sp, SqlCommand executeCommand)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter ad;
            SqlConnection connection;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandTimeout = 120;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = executeCommand;
                ad.Fill(ds);
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return ds;
        }

        public DataSet GetDataSet_SP(String sp, Dictionary<String, Object> parameters)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter ad;
            SqlConnection connection;
            SqlCommand command;
            SqlParameter parameter;
            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandTimeout = 900;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = sp;

                foreach (string key in parameters.Keys)
                {
                    command.Parameters.AddWithValue(key, parameters[key]);
                }
                //foreach (KeyValuePair<string, object> kvp in parameters)
                //{
                //    command.Parameters.AddWithValue(key, value);
                //}

                ds = new DataSet();
                ad = new SqlDataAdapter();
                ad.SelectCommand = command;
                ad.Fill(ds);
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return ds;
        }

        public Boolean ExecuteNonQuery(String sql)
        {
            Boolean r = false; ;
            SqlCommand command;
            SqlConnection connection;
            Int32 rVal = 0;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;
                rVal = command.ExecuteNonQuery();
                if (rVal > 0)
                {
                    r = true;
                }
            }
            catch (Exception e)
            {
                r = false;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Boolean ExecuteNonQuery_Parameter(String sql, SqlCommand executeCommand)
        {
            Boolean r = false; ;
            SqlConnection connection;
            Int32 rVal = 0;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandType = CommandType.Text;
                executeCommand.CommandText = sql;
                rVal = executeCommand.ExecuteNonQuery();
                if (rVal > 0)
                {
                    r = true;
                }
            }
            catch (Exception e)
            {
                r = false;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Boolean ExecuteNonQuery_Parameter(String sql, Dictionary<String, Object> parameters)
        {
            Boolean r = false; ;
            SqlConnection connection;
            SqlCommand command;
            Int32 rVal = 0;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }
                rVal = command.ExecuteNonQuery();
                if (rVal > 0)
                {
                    r = true;
                }
            }
            catch (Exception e)
            {
                r = false;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Boolean ExecuteSP(String sp, SqlCommand executeCommand)
        {
            Boolean r = false;
            SqlConnection connection;
            Int32 rVal = 0;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;
                rVal = executeCommand.ExecuteNonQuery();
                if (rVal > 0)
                {
                    r = true;
                }
            }
            catch (Exception e)
            {
                r = false;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Boolean ExecuteSP(String sp, Dictionary<String, Object> parameters)
        {
            Boolean r = false;
            SqlConnection connection;
            SqlCommand command;
            Int32 rVal = 0;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = sp;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }
                rVal = command.ExecuteNonQuery();
                if (rVal > 0)
                {
                    r = true;
                }
            }
            catch (Exception e)
            {
                r = false;
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Object ExecuteScalar(String sql)
        {
            Object r = null;
            SqlCommand command;
            SqlConnection connection;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;
                r = command.ExecuteScalar();
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Object ExecuteScalar_Parameter(String sql, SqlCommand executeCommand)
        {
            Object r = null;
            SqlConnection connection;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandType = CommandType.Text;
                executeCommand.CommandText = sql;
                r = executeCommand.ExecuteScalar();
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Object ExecuteScalar_Parameter(String sql, Dictionary<String, Object> parameters)
        {
            Object r = null;
            SqlConnection connection;
            SqlCommand command;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                command.CommandText = sql;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }
                r = command.ExecuteScalar();
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Object ExecuteScalar_SP(String sp, SqlCommand executeCommand)
        {
            Object r = null;
            SqlConnection connection;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();
                executeCommand.Connection = connection;
                executeCommand.CommandType = CommandType.StoredProcedure;
                executeCommand.CommandText = sp;

                r = executeCommand.ExecuteScalar();
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

        public Object ExecuteScalar_SP(String sp, Dictionary<String, Object> parameters)
        {
            Object r = null;
            SqlConnection connection;
            SqlCommand command;

            connection = null;

            try
            {
                connection = new SqlConnection();
                connection.ConnectionString = myConnectionString;
                connection.Open();

                command = new SqlCommand();
                command.Connection = connection;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = sp;

                foreach (KeyValuePair<string, object> kvp in parameters)
                {
                    command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                }
                r = command.ExecuteScalar();
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
            return r;
        }

    }
}
