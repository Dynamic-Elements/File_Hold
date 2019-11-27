﻿using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;

namespace Connection_Class
{
    public class Connection_Query
    {

        //string lcServer = "playgroup.database.windows.net";
        //string lcODBC = "ODBC Driver 17 for SQL Server";
        //string lcDB = "tb_HelpingHand";
        //string lcUser = "tbmaster";
        //string lcProv = "SQLOLEDB";
        //string lcPass = "Smartman55";
        // string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
        public static string lcConnectionString = "Driver={ODBC Driver 17 for SQL Server};Provider=SQLOLEDB;Server=playgroup.database.windows.net;DATABASE=tb_HelpingHand;Uid=tbmaster; Pwd=Smartman55;";
        public static OdbcConnection con;

        public static void OpenConection()
        {
            // string lcConnectionString = "Driver={ODBC Driver 17 for SQL Server};Provider=SQLOLEDB;Server=playgroup.database.windows.net;DATABASE=tb_HelpingHand;Uid=tbmaster; Pwd=Smartman55;";
            // OdbcConnection con;
            con = new OdbcConnection(lcConnectionString);
            con.Open();
        }
        public static void CloseConnection()
        {
            con.Close();
        }
        public static void ExecuteQueries(string Query_)
        {
            OdbcCommand cmd = new OdbcCommand(Query_, con);
            cmd.ExecuteNonQuery();
        }
        public static OdbcDataReader DataReader(string Query_)  // SqlDataReader
        {
            OdbcCommand cmd = new OdbcCommand(Query_, con);
            OdbcDataReader dr = cmd.ExecuteReader();  // SqlDataReader
            return dr;
        }
        public static object ShowDataInGridView(string Query_)
        {
            SqlDataAdapter dr = new SqlDataAdapter(Query_, lcConnectionString);  // SqlDataAdapter  SqlDataAdapter
            DataSet ds = new DataSet();
            dr.Fill(ds);
            object dataum = ds.Tables[0];
            return dataum;
        }
    }
}
