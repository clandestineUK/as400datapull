using System;
using System.Collections.Generic;
using System.Text;
using cwbx;
using System.Data.Odbc;
using System.IO;
namespace BotanyBayProductUpdate
{
    class Program
    {
             
        static void Main(string[] args)
        {
            string result = string.Empty;

            StringConverter stringConverter = new StringConverterClass();

            //Define as400 system and connect
            AS400System system = new AS400System();
            system.Define("AS400");
            system.UserID = "PCS400";
            system.Password = "PCS400";
            system.IPAddress = "192.20.20.27";
            system.Connect(cwbcoServiceEnum.cwbcoServiceRemoteCmd);

            //Check connection
            if (system.IsConnected(cwbcoServiceEnum.cwbcoServiceRemoteCmd) == 1)
            {
                //Create program object and link to a system
                cwbx.Program program = new cwbx.Program();
                //  program.LibraryName = "WAKY2KOBJ";
                program.LibraryName = "SILIB";
                program.ProgramName = "BOTBAY01";
                program.system = system;
              

                //Finally call the program

                try
                {
                    // program.Call();
                  
                   using (OdbcConnection conn = new OdbcConnection("Driver={iseries access odbc driver};system=s654d1bb;uid=PCS400;pwd=PCS400;"))
                    {
                        OdbcCommand command = new OdbcCommand("select * from silib.botbay01up", conn);
                        conn.Open();
                        OdbcDataReader reader = command.ExecuteReader();
                       
                        while (reader.Read())
                        {
                            Console.WriteLine("barcode={0}", reader[0]);
                            Console.WriteLine("price={0}", reader[1]);
                            Console.WriteLine("description={0}", reader[2]);
                            Console.WriteLine("dept={0}", reader[3]);
                        }
                        reader.Close();

                        var csv = new StringBuilder();

                        var first = reader[0].ToString();
                        var second = reader[1].ToString();
                        var third = reader[2].ToString();
                        var fourth = reader[3].ToString();

                        var newLine = string.Format("{0},{1}", first, second, third, fourth);

                        string filePath = "C:\\Users\\morrise\\Desktop";

                        File.WriteAllText(filePath, csv.ToString());
                    }                                        
                }
                catch (Exception ex)
                {
                    if (system.Errors.Count > 0)
                    {
                        foreach (cwbx.Error error in system.Errors)
                        {
                            Console.WriteLine(error.Text);
                        }
                    }

                    if (program.Errors.Count > 0)
                    {
                        foreach (cwbx.Error error in program.Errors)
                        {
                            Console.WriteLine(error.Text);
                        }
                    }
                } 
            }
            system.Disconnect(cwbcoServiceEnum.cwbcoServiceAll);
            Console.WriteLine("Done");
            Console.ReadKey();
        }

      

    }
}
