using System;
using System.Data.Odbc;
using System.IO;
using System.Net.Mail;
using cwbx;
using NLog;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace BotanyBayProductUpdate
{
    class Program
    {
        //Nlog
        private static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            //properties for sending mail
            MailMessage mail = new MailMessage("elliot.morris@wsg.co.uk", "elliot.morris@wsg.co.uk");
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
            var fileName = "DoubleTWO_3860_Products_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
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

            //Setup xlsx file
            Excel.Application oApp;
            Excel.Worksheet oSheet;
            Excel.Workbook oBook;
            oApp = new Excel.Application();
            oBook = oApp.Workbooks.Add();
            oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);

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
                  //  program.Call();
                    logger.Info("program called on as400");
                    //creating a odbc connection to the as400
                    using (OdbcConnection conn = new OdbcConnection("Driver={iseries access odbc driver};system=s654d1bb;uid=PCS400;pwd=PCS400;"))
                    {
                        OdbcCommand command = new OdbcCommand("select * from silib.botbay01up where budept = '3860' order by bubarcode", conn);
                        conn.Open();
                        OdbcDataReader reader = command.ExecuteReader();
                        logger.Info("connection made to as400 via odbc");

                        if (reader.HasRows == true)
                        {
                            if (File.Exists(filePath + fileName) == true)
                            {
                                File.Delete(filePath + fileName);
                            } 

                            int rowNumber = 1;
                            // column 1
                            oSheet.Cells[rowNumber, 1] = "Barcode";
                            oSheet.Cells[rowNumber, 1].ColumnWidth = 15;
                            oSheet.Cells[rowNumber, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                            //column 2
                            oSheet.Cells[rowNumber, 2] = "Price";
                            oSheet.Cells[rowNumber, 2].ColumnWidth = 8;
                            oSheet.Cells[rowNumber, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                            //column 3
                            oSheet.Cells[rowNumber, 3] = "Description";
                            oSheet.Cells[rowNumber, 3].ColumnWidth = 40;

                            while (reader.Read())
                            {
                                Console.WriteLine("Barcode = {0} Price = {0} Description = {0}", reader[0], reader[1], reader[2]);

                                rowNumber++;
                                //column 1
                                oSheet.Cells[rowNumber, 1] = reader[0].ToString();
                                oSheet.Cells[rowNumber, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                oSheet.Cells[rowNumber, 1].NumberFormat = "0";
                                //column 2
                                oSheet.Cells[rowNumber, 2] = reader[1].ToString();
                                oSheet.Cells[rowNumber, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                oSheet.Cells[rowNumber, 2].NumberFormat = "0.00";
                                //column 3
                                oSheet.Cells[rowNumber, 3] = reader[2].ToString();
                                //creating the string format
                                var newLine = string.Format("{0}, {1}, {2}", oSheet.Cells[rowNumber, 1], oSheet.Cells[rowNumber, 2], oSheet.Cells[rowNumber, 3]);
                            }
                          
                                oBook.SaveAs(filePath + fileName);
                                oBook.Close();
                                oApp.Quit();
                                logger.Info("File created " + fileName);
                                fileCsv = new Attachment(filePath + fileName);
                                mail.Subject = "Double TWO Product Update ";
                                mail.Body = "Double TWO Product Update.";
                                mail.Attachments.Add(fileCsv);
                                client.Send(mail);
                                Console.WriteLine("Mail sent");
                                logger.Info("Mail sent with attachment");                                               
                         }
                        else
                        {
                            logger.Info("No updates to send today " + DateTime.Now.ToString());
                        }
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
                //closing connection
                system.Disconnect(cwbcoServiceEnum.cwbcoServiceAll);
                logger.Info("Connection has been disconneted");
                Console.ReadKey();
        }             
     }
}
