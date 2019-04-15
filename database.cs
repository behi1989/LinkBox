using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace NSP_LinkBox
{
    class database
    {
        public static DataTable select(string select)
        {
            SqlConnection con = new SqlConnection();
            con.ConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\db\Linkboxdb.mdf;Integrated Security=True;Connect Timeout=15";
            SqlDataAdapter da = new SqlDataAdapter(select,con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public static void DoIMD(string IMD)
        {
            SqlConnection con = new SqlConnection();
            con.ConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\db\Linkboxdb.mdf;Integrated Security=True;Connect Timeout=15";
            SqlCommand sc = new SqlCommand(IMD,con);
            //sc.CommandText = IMD;
            //sc.Connection = con;
            con.Open();
            sc.ExecuteNonQuery();
            con.Close();
        }
    }
}
