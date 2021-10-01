using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace ReadingDataFromExcel
{
    public class Read
    {

        public  void button1_Click()
        {

            try
            {
                System.Data.OleDb.OleDbConnection Baglanti;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

                string sql = null;
                //Baglanti =  new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=‪C:\\Users\Emre\\Desktop\\SebzeMeyve-Kopya.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                Baglanti = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\Emre\\Desktop\\SebzeMeyve-Kopya1.xlsx; Extended Properties ='Excel 12.0 xml;HDR=YES;'");
                Baglanti.Open();
                myCommand.Connection = Baglanti;
                sql = "Select * from [Sayfa1$]";
                myCommand.CommandText = sql;
                System.Data.OleDb.OleDbDataReader dr = myCommand.ExecuteReader();

                /*datareader içindeki verileri okuma*/
                while (dr.Read())
                {
                    Console.WriteLine(String.Format("Name :{0} ,Barcode:{1} ve Category ID :{2}", dr[0], dr[1], dr[2].ToString()));
                }
                Baglanti.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }
    }
}
