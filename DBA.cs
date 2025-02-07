using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace PrintKernel
{
    public enum ConnStr
    {
        BBCMS
    }

    class DBA
    {
        public string LastError = "";

        public SqlDataReader executeParameterReader(int timerIndex)
        {
            ConnStr ConnStr = ConnStr.BBCMS;
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = ConfigurationManager.AppSettings[ConnStr.ToString()];
            conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT [C_System],[Product],[P_Type],[PrintNum] FROM [dbo].[PrintParameter] WHERE [IsEnable] = 1 AND [PNo]= @TimerIndex";
            cmd.Parameters.AddWithValue("@TimerIndex", timerIndex);
            SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            return reader;
        }
        public System.Data.DataTable GetDataTable(string sql, ConnStr ConnStr = ConnStr.BBCMS, List<InputPara> inputpara = null)
        {
            try
            {
                if (!string.IsNullOrEmpty(sql))
                {
                    System.Data.DataTable dt = new System.Data.DataTable();
                    SqlConnection conn = new SqlConnection();
                    conn.ConnectionString = ConfigurationManager.AppSettings[ConnStr.ToString()];
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandTimeout = 10;
                    conn.Open();

                    if (inputpara != null)
                    {
                        foreach (InputPara input in inputpara)
                        {
                            cmd.Parameters.Add(input.name, input.dbtype).Value = input.value;
                        }
                    }

                    using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                    {
                        a.Fill(dt);
                    }

                    conn.Close();

                    cmd.Dispose();
                    conn.Dispose();

                    return dt;
                }
            }
            catch (Exception ex)
            {
                this.LastError = ex.Message;
                return null;
            }

            return null;
        }

        //public System.Data.DataTable GetDataTable(string sql, ConnStr ConnStr = ConnStr.BBCMS, List<InputPara> inputpara = null)
        //{
        //    DataTable dt = null;
        //    if (string.IsNullOrEmpty(sql))
        //        return dt;

        //    try
        //    {
        //        using (var conn = new SqlConnection(ConfigurationManager.AppSettings[ConnStr.ToString()]))
        //        {
        //            using (var cmd = new SqlCommand(sql, conn))
        //            {
        //                cmd.CommandTimeout = 10;

        //                // 添加參數
        //                if (inputpara != null)
        //                {
        //                    foreach (InputPara input in inputpara)
        //                    {
        //                        cmd.Parameters.Add(input.name, input.dbtype).Value = input.value;
        //                    }
        //                }

        //                conn.Open();
        //                // 填充 DataTable
        //                using (var adapter = new SqlDataAdapter(cmd))
        //                {
        //                    dt = new DataTable();
        //                    adapter.Fill(dt);
        //                }
        //                conn.Close();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        this.LastError = ex.Message;
        //        return null;
        //    }
        //}

        public bool ExeCuteNonQuery(string sql, ConnStr ConnStr = ConnStr.BBCMS, List<InputPara> inputpara = null)
        {
            using (SqlConnection cn = new SqlConnection(ConfigurationManager.AppSettings[ConnStr.ToString()]))
            {
                using (SqlCommand cmd = new SqlCommand(sql, cn))
                {
                    try
                    {
                        cmd.CommandType = CommandType.Text;
                        if (inputpara != null)
                        {
                            foreach (InputPara input in inputpara)
                            {
                                cmd.Parameters.Add(input.name, input.dbtype).Value = input.value;
                            }
                        }
                        cn.Open();
                        int rowsAffected = cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        this.LastError = ex.Message;
                        return false;
                    }
                    finally
                    {
                        cn.Close();
                        cn.Dispose();
                    }
                }
            }

            return true;
        }
    }

    public class InputPara
    {
        public string name;
        public string value;
        public SqlDbType dbtype;

        public InputPara()
        {
        }

        public InputPara(string pName, string pValue, SqlDbType pType)
        {
            name = pName;
            value = pValue;
            dbtype = pType;
        }
    }
}
