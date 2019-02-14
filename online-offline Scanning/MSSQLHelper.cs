using System;
using System.Collections;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.Common;
using System.Collections.Generic;
using System.IO;


public class MSSQLHelper
{
    //Database connect string
    //public static string connectionString = "Data Source=10.126.64.131;Initial Catalog=MESOtherData;Persist Security Info=True;User ID=XMUpdate;Password=695Xm@123#";//惠州SCBC
    //public static string connectionString = "Data Source=192.168.0.104;Initial Catalog=MESOtherData;Persist Security Info=True;User ID=arturas;Password=gvard_19";//俄罗斯PKV
    public static string connectionString = "Data Source=192.168.109.129;Initial Catalog=MESOtherData;Persist Security Info=True;User ID=XMUpdate1;Password=695Xm@123#";//test DataBase
    public MSSQLHelper()
	{

	}

    /// <summary>
    /// Execute SQL statement,return number of affected records 
    /// </summary>
    /// <param name="SQLString">SQL command</param>
    /// <returns>return execute number</returns>
    public static int ExecuteSql(string SQLString)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            using (SqlCommand cmd = new SqlCommand(SQLString, connection))
            {
                try
                {
                    connection.Open();
                    //cmd.CommandTimeout = Times;
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (System.Data.SqlClient.SqlException e)
                {
                    connection.Close();
                    throw e;
                }
            }
        }
    }
    /// <summary>
    /// Get SQL statement running result(object)
    /// </summary>
    /// <param name="SQLString">SQL command</param>
    /// <returns>object result</returns>
    public static object GetSingle(string SQLString)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            using (SqlCommand cmd = new SqlCommand(SQLString, connection))
            {
                try
                {
                    connection.Open();
                    object obj = cmd.ExecuteScalar();
                    if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                    {
                        return null;
                    }
                    else
                    {
                        return obj;
                    }
                }
                catch (System.Data.SqlClient.SqlException e)
                {
                    connection.Close();
                    throw e;
                }
            }
        }
    }

    /// <summary>
    /// Execute SQL statement,return Dataset
    /// </summary>
    /// <param name="SQLString"></param>
    /// <returns>Return dataset format</returns>
    public static DataSet Query(string SQLString)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                SqlDataAdapter command = new SqlDataAdapter(SQLString, connection);
                command.Fill(ds, "ds");
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                throw new Exception(ex.Message);
            }
            return ds;
        }
    }



}
