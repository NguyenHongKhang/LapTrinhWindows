using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace WindowsFormsApp1.DAO
{
    public class DataProvider
    {


        private static DataProvider _Instance;
        public static DataProvider Instance
        {
            get { if (_Instance == null) _Instance = new DataProvider(); return DataProvider._Instance; }
            private set { DataProvider._Instance = value; }
        }
        private DataProvider() { }

        private string connectionStr = @"Data Source=localhost\SQLEXPRESS02; Initial Catalog=CafeShop; User Id=sa; Password=sa";


        public DataTable ExecuteQuery(string query, object[] parameter = null) 
        {
            DataTable data = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionStr))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter!=null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;

                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                SqlDataAdapter adapter = new SqlDataAdapter(command);

                adapter.Fill(data);

                connection.Close();
            }

            return data;
        }
        public int ExecuteNonQuery(string query, object[] parameter = null)
        {
            int data = 0;

            using (SqlConnection connection = new SqlConnection(connectionStr))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;

                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                data = command.ExecuteNonQuery();

                connection.Close();
            }

            return data;
        }
        public object ExecuteScalarQuery(string query, object[] parameter = null)
        {
            object data = 0;

            using (SqlConnection connection = new SqlConnection(connectionStr))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;

                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                data = command.ExecuteScalar();

                connection.Close();
            }

            return data;
        }
        public int ExecuteNonQueryTransaction(string query, List<object[]> ListParameter)
        {
            int data = 0;

            using (SqlConnection connection = new SqlConnection(connectionStr))
            {
                connection.Open();

                SqlTransaction trans = connection.BeginTransaction();
                try
                {
                    foreach (object[] o1 in ListParameter)
                    {
                        SqlCommand command = new SqlCommand(query, connection, trans);
                        if (o1 != null)
                        {
                            string[] listPara = query.Split(' ');
                            int i = 0;

                            foreach (string item in listPara)
                            {
                                if (item.Contains('@'))
                                {
                                    command.Parameters.AddWithValue(item, o1[i]);
                                    i++;
                                }
                            }
                        }
                        data += command.ExecuteNonQuery();
                    }
                    trans.Commit();
                }

                catch
                {
                    trans.Rollback();
                    return 0;
                }
                connection.Close();
            }
            return data;
        }
    }
}
