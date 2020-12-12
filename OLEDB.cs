using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace AutomatizadorDeTransferenciaDeDados
{
    public class OLEDB
    {

        public void Conecc()
        {
              string path = @"G:/20_MAINTENANCE/02-PFMAP/KDFU.xls";
              string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=Excel 8.0;";
       // string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=G:/20_MAINTENANCE/02-PFMAP/KDFU.xls;Extended Properties='Excel 8.0;HDR=YES;'";
            






            using (OleDbConnection _conexao = new OleDbConnection(connStr))
            {
                string sql = null;
                _conexao.Open();
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                myCommand.Connection = _conexao;
                sql = "SELECT * FROM [KDFU SAF TRUCK 2019]";//"Insert into [Sheet1$] (id,name) values('5','e')";
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                _conexao.Close();

            }
            



        }


    }
}
