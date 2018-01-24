using System;
using System.Text;
using System.Data.Odbc;
using System.IO;
using System.Net.Mail;
using cwbx;
using NLog;

namespace BotanyBayProductUpdate
{
    class Program
    {
        //Nlog
        private static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            //properties for sending mail
            MailMessage mail = new MailMessage("elliot.morris@wsg.co.uk", "elliotmorris115@gmail.com");
            SmtpClient client = new SmtpClient
            {
                Port = 25,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Host = "192.20.20.11"
            };
            Attachment fileCsv;
            logger.Info("Connection to smtp");
            //File path location
            var fileName = "BotBayUpdate " + DateTime.Now.ToString("dd-MM-yyyy") + ".csv";
            var filePath = @"C:\Users\morrise\Desktop\";

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
                cwbx.Program program = new cwbx.Program
                {
                    //program.LibraryName = "WAKY2KOBJ";
                    LibraryName = "SILIB",
                    ProgramName = "BOTBAY01",
                    system = system
                };
                logger.Info("Connection to as400"); 
                //call the program
                try
                {
                    // program.Call();
                    logger.Info("program called on as400");
                    //creating a odbc connection to the as400
                    using (OdbcConnection conn = new OdbcConnection("Driver={iseries access odbc driver};system=s654d1bb;uid=PCS400;pwd=PCS400;"))
                    {
                        OdbcCommand command = new OdbcCommand("select * from silib.botbay01up", conn);
                        conn.Open();
                        OdbcDataReader reader = command.ExecuteReader();
                        logger.Info("connection made to as400 via odbc");
                        var csv = new StringBuilder();
                        //header string
                        var header = string.Format("Barcode, Price, Description, Dept");
                        //appending the headers
                        csv.AppendLine(header);
                        logger.Info("Appending headers to the csv");
                        while (reader.Read())
                        {                            
                            Console.WriteLine("Barcode = {0} Price = {0} Description = {0} Dept = {0} ", reader[0],  reader[1], reader[2], reader[3]);
                            //reading in the values into the csv
                            var first = reader[0].ToString();
                            var second = reader[1].ToString();
                            var third = reader[2].ToString();
                            var fourth = reader[3].ToString();
                            //creating the string format
                            var newLine = string.Format("{0}, {1}, {2}, {3}", first, second, third, fourth);
                            //appending the values
                            csv.AppendLine(newLine);
                            logger.Info("Appending data to csv");
                        }
                            File.WriteAllText(filePath + fileName, csv.ToString());
                            logger.Info("File created" + fileName);
                    }                   
                 }
                    catch (Exception ex)
                    {

                    if (system.Errors.Count > 0)
                        {
                            foreach (cwbx.Error error in system.Errors)
                            {
                                Console.WriteLine(error.Text);
                                logger.Error("The error is " + error.Text + ex);
                            }
                        }
                        if (program.Errors.Count > 0)
                        {
                            foreach (cwbx.Error error in program.Errors)
                            {
                                Console.WriteLine(error.Text);
                                logger.Error("The error is " + error.Text + ex);
                            }
                        }
                    } 
                }
                if (filePath.Length >= 1)
                {
                    fileCsv = new Attachment(filePath + fileName);
                    mail.Subject = "BotBay product update ";
                    mail.Body = "Hi, \n Please see the csv file attached. ";
                    mail.Attachments.Add(fileCsv);
                    client.Send(mail);
                    Console.WriteLine("Mail sent");
                    logger.Info("Mail sent with attachment");
                }
                else
                {
                    mail.Subject = "BotBay product update (No updates)";
                    mail.Body = "There are no new upates today. ";
                    client.Send(mail);
                    Console.WriteLine("Csv is empty");
                logger.Info("Mail sent no updates so no attachment");
                }
                //closing connection
                system.Disconnect(cwbcoServiceEnum.cwbcoServiceAll);
                logger.Info("Connection has been disconneted");
                Console.ReadKey();
        }             
     }
}
