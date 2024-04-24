using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Mail;

namespace cmdb
{
    class Program
    {
        static void Main(string[] args)
        {
            CMDBApplication();
        }

        private static void CMDBApplication()
        {
            try
            {
                WriteLog("Application Started at ==>" + DateTime.Now.ToString());
                DataSet ds = new DataSet();
                ds = GetCMDBdata();  // Get CMDB data set from service now instance 
                string DestinationFolder = ConfigurationSettings.AppSettings["DestinationFolder"].ToString();
                ExportDataSetToCsvFile(ds, DestinationFolder);  // Generate CSV file from Dataset to destination location 
            }
            catch (Exception ex)
            {
                WriteLog("error while processing " + ex);

            }
            finally
            {
                SendMail();
                WriteLog("Application Completed at ==>" + DateTime.Now.ToString());
            }
        }

        private static void SendMail()
        {
            try
            {
                var sendEmail = ConfigurationSettings.AppSettings["sendEmail"];

                if (sendEmail.ToLower() == "true")
                {
                    var from = ConfigurationSettings.AppSettings["mailFrom"];
                    var to = ConfigurationSettings.AppSettings["mailTo"];

                    var logFolder = ConfigurationSettings.AppSettings["LogPath"];
                    var fileName = ConfigurationSettings.AppSettings["LogFileName"].ToString() + " " + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt";
                    string filepath = logFolder + "\\" + fileName;

                    var logs = File.ReadAllText(filepath).ToLower();

                    if (!logs.Contains("error") &&
                        !logs.Contains("unable") &&
                        !logs.Contains("failed") &&
                        !logs.Contains("exception"))
                        return;

                    MailMessage mail = new MailMessage(from, to);
                    SmtpClient client = new SmtpClient();
                    client.Port = 25;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = ConfigurationSettings.AppSettings["smtpHost"];
                    mail.Subject = ConfigurationSettings.AppSettings["mailSubject"];
                    mail.Body = ConfigurationSettings.AppSettings["mailBody"];
                    //client.EnableSsl = true;



                    mail.Attachments.Add(new Attachment(logFolder + "\\" + fileName));

                    client.Send(mail);
                }
            }
            catch (Exception ex)
            {
                WriteLog("error while sending an email " + ex);
            }

        }

        private static DataSet GetCMDBdata()
        {
            DataSet myDataSet = null;
            try
            {
                string username = ConfigurationSettings.AppSettings["ServiceNowUserName"].ToString();
                string password = ConfigurationSettings.AppSettings["ServiceNowPassword"].ToString();
                string url = ConfigurationSettings.AppSettings["CMDB_url"].ToString();

                var auth = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(username + ":" + password));

                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Headers.Add("Authorization", auth);
                request.Timeout = Convert.ToInt32(ConfigurationSettings.AppSettings["ServiceTimeOut"].ToString());
                request.Method = "Get";

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    var res = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    JObject joResponse = JObject.Parse(res.ToString());
                    myDataSet = JsonConvert.DeserializeObject<DataSet>(res);   // Convert Json to data set 

                    return myDataSet;
                }
            }
            catch (Exception ex)
            {
                WriteLog("Error Occured while taking data from Service now instance ==>" + ex + "\n at " + DateTime.Now.ToString());
            }
            return myDataSet;

        }

        internal static void ExportDataSetToCsvFile(DataSet _DataSet, string DestinationCsvDirectory)
        {
            string CSVFileName = ConfigurationSettings.AppSettings["CSVFileName"].ToString();
            try
            {
                foreach (DataTable DDT in _DataSet.Tables)
                {
                    //String MyFile = DestinationCsvDirectory + "\\" + CSVFileName + DateTime.Now.ToString("MM/dd/yyyy HH:mm") + ".csv";//+ DateTime.Now.ToString("ddMMyyyyhhMMssffff")
                    String MyFile = @DestinationCsvDirectory + "\\" + CSVFileName + DateTime.Now.ToString("yyyy-MMMM-dd-hhMM tt") + ".csv";//+ DateTime.Now.ToString("ddMMyyyyhhMMssffff")
                    using (var outputFile = File.CreateText(MyFile))
                    {
                        String CsvText = string.Empty;

                        foreach (DataColumn DC in DDT.Columns)
                        {
                            if (CsvText != "")
                                CsvText = CsvText + "," + DC.ColumnName.ToString();
                            else
                                CsvText = DC.ColumnName.ToString();
                        }
                        outputFile.WriteLine(CsvText.ToString().TrimEnd(','));
                        CsvText = string.Empty;

                        foreach (DataRow DDR in DDT.Rows)
                        {
                            foreach (DataColumn DCC in DDT.Columns)
                            {
                                if (CsvText != "")
                                    CsvText = CsvText + "," + DDR[DCC.ColumnName.ToString()].ToString();
                                else
                                    CsvText = DDR[DCC.ColumnName.ToString()].ToString();
                            }
                            outputFile.WriteLine(CsvText.ToString().TrimEnd(','));
                            CsvText = string.Empty;
                        }
                        System.Threading.Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("Error Occured while generating CSV ==>" + ex + "\n at " + DateTime.Now.ToString());
            }
        }


        public static void WriteLog(string dataInfo)
        {
            try
            {
                Console.WriteLine(dataInfo);
                string FolderPath = ConfigurationSettings.AppSettings["LogPath"].ToString();

                Directory.CreateDirectory(FolderPath);

                var fileName = ConfigurationSettings.AppSettings.Get("LogFileName") + " " + DateTime.Now.ToString("dd-MMM-yyyy");

                using (StreamWriter w = File.AppendText(FolderPath + "\\" + fileName + ".txt"))
                {
                    //w.WriteLine("----------{0}-----------", DateTime.Now.ToLongTimeString());
                    w.WriteLine(DateTime.Now.ToLongTimeString() + " ==> {0}", dataInfo);
                }
            }
            catch
            {

            }
        }
    }
}
