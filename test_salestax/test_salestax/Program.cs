using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using sales_tax;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace test_salestax
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                
                string connString = "user id=mosaic;password=P@ssword123;initial catalog=mosaic_2013_projects;data source=192.168.100.19;Connect Timeout=600";
                SqlConnection myConn = new SqlConnection(connString);
                DataTable dt = new DataTable();

                string cmd = "select state, count(*) as count from BCS_21_A_Customers (nolock) where mailflag != 'FORGN' and mailflag != 'SAMP' group by state order by state";

                SqlDataAdapter da = new SqlDataAdapter(cmd, myConn);
                da.Fill(dt);
                SalesTax mytax = new SalesTax();
                mytax.RunSalesTax("99999", "1", "1", dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
            finally
            {
                Console.WriteLine("Done");
                Console.ReadLine();
            }
        }
    }
}
