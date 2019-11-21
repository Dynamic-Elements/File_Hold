using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;

namespace Accounting_PL
{
    public class Conn_cl
    {

        //string lcServer = "playgroup.database.windows.net";
        //string lcODBC = "ODBC Driver 17 for SQL Server";
        //string lcDB = "tb_HelpingHand";
        //string lcUser = "tbmaster";
        //string lcProv = "SQLOLEDB";
        //string lcPass = "Smartman55";
        // string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
        string lcConnectionString = "Driver={ODBC Driver 17 for SQL Server};Provider=SQLOLEDB;Server=playgroup.database.windows.net;DATABASE=tb_HelpingHand;Uid=tbmaster; Pwd=Smartman55;";
        OdbcConnection con;

        public void OpenConection()
        {
            // string lcConnectionString = "Driver={ODBC Driver 17 for SQL Server};Provider=SQLOLEDB;Server=playgroup.database.windows.net;DATABASE=tb_HelpingHand;Uid=tbmaster; Pwd=Smartman55;";
            // OdbcConnection con;
            con = new OdbcConnection(lcConnectionString);
            con.Open();
        }
        public void CloseConnection()
        {
            con.Close();
        }
        public void ExecuteQueries(string Query_)
        {
            OdbcCommand cmd = new OdbcCommand(Query_, con);
            cmd.ExecuteNonQuery();
        }
        public OdbcDataReader DataReader(string Query_)  // SqlDataReader
        {
            OdbcCommand cmd = new OdbcCommand(Query_, con);
            OdbcDataReader dr = cmd.ExecuteReader();  // SqlDataReader
            return dr;
        }
        public object ShowDataInGridView(string Query_)
        {
            SqlDataAdapter dr = new SqlDataAdapter(Query_, lcConnectionString);  // SqlDataAdapter  SqlDataAdapter
            DataSet ds = new DataSet();
            dr.Fill(ds);
            object dataum = ds.Tables[0];
            return dataum;
        }
    }
}
