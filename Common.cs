using ArcherExtractQlikview;
using SNOW;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ArcherCommonClass
{
    public static class Common
    {
        public static DataTable XmlToDT(string reportXML)
        {
            var ds = new DataSet();
            //ds.ReadXml(new System.IO.StringReader(reportXML));

            var xml = XElement.Parse(reportXML);

            var columns = xml.Descendants("FieldDefinition").Select(tmp => new Field { FieldID = tmp.Attribute("id").Value, FieldName = tmp.Attribute("name").Value }).ToList();

            DataTable dtData = new DataTable("Record");

            //for (int i = 0; i < ds.Tables["FieldDefinition"].Rows.Count; i++)
            for (int i = 0; i < columns.Count; i++)
            {
                if (dtData.Columns.Contains(columns[i].FieldName))
                {
                    columns[i].FieldName = columns[i].FieldName + "_" + i;
                }

                dtData.Columns.Add(columns[i].FieldName);
            }


            var columnsLvl3 = new List<string>();
            var columnsLvl4 = new List<string>();
            var columnsLvl5 = new List<string>();

            var xLanguageDetails = XElement.Parse(reportXML);

            foreach (XElement xUserNode in xLanguageDetails.Elements("Record"))
            {
                DataRow dr = dtData.NewRow();

                ReadFieldValue(xUserNode, ref ds, ref dtData, ref dr, columns);

                for (int i = 0; i < xUserNode.Elements("Record").Count(); i++)
                {
                    XElement xxUserNode = xUserNode.Elements("Record").ElementAt(i);

                    if (i != 0)
                    {
                        var copy = dr.ItemArray;
                        dr = dtData.NewRow();
                        dr.ItemArray = copy;

                        foreach (var column in columnsLvl3)
                        {
                            dr[column] = string.Empty;
                        }

                    }

                    ReadFieldValue(xxUserNode, ref ds, ref dtData, ref dr, columns);


                    for (int j = 0; j < xxUserNode.Elements("Record").Count(); j++)
                    {
                        XElement xxxUserNode = xxUserNode.Elements("Record").ElementAt(j);

                        if (j != 0)
                        {
                            var copy = dr.ItemArray;
                            dr = dtData.NewRow();
                            dr.ItemArray = copy;

                            foreach (var column in columnsLvl4)
                            {
                                dr[column] = string.Empty;
                            }
                        }

                        ReadFieldValue(xxxUserNode, ref ds, ref dtData, ref dr, columns);
                        columnsLvl3 = ReadFields(xxxUserNode, ref ds, columns);

                        for (int K = 0; K < xxxUserNode.Elements("Record").Count(); K++)
                        {
                            XElement xxxxUserNode = xxxUserNode.Elements("Record").ElementAt(K);

                            if (K != 0)
                            {
                                var copy = dr.ItemArray;
                                dr = dtData.NewRow();
                                dr.ItemArray = copy;

                                foreach (var column in columnsLvl5)
                                {
                                    dr[column] = string.Empty;
                                }
                            }

                            ReadFieldValue(xxxxUserNode, ref ds, ref dtData, ref dr, columns);
                            columnsLvl4 = ReadFields(xxxxUserNode, ref ds, columns);

                            for (int l = 0; l < xxxxUserNode.Elements("Record").Count(); l++)
                            {
                                XElement xxxxxUserNode = xxxxUserNode.Elements("Record").ElementAt(l);

                                if (l != 0)
                                {
                                    var copy = dr.ItemArray;
                                    dr = dtData.NewRow();
                                    dr.ItemArray = copy;
                                }

                                ReadFieldValue(xxxxxUserNode, ref ds, ref dtData, ref dr, columns);
                                columnsLvl5 = ReadFields(xxxxxUserNode, ref ds, columns);

                                dtData.Rows.Add(dr);
                            }

                            if (xxxxUserNode.Descendants("Record").Count() == 0)
                            {
                                dtData.Rows.Add(dr);
                            }
                        }

                        if (xxxUserNode.Descendants("Record").Count() == 0)
                        {
                            dtData.Rows.Add(dr);
                        }
                    }

                    if (xxUserNode.Descendants("Record").Count() == 0)
                    {
                        dtData.Rows.Add(dr);
                    }
                }
                if (xUserNode.Descendants("Record").Count() == 0)
                {
                    dtData.Rows.Add(dr);
                }
            }
            return dtData;
        }

        public static void XmlToDTNew(string reportXML, int loopCount, ref StringBuilder sb)
        {

            var ds = new DataSet();
            ds.ReadXml(new System.IO.StringReader(reportXML));

            if (loopCount == 0)
            {
                var columns = new List<string>();

                for (int i = 0; i < ds.Tables["FieldDefinition"].Rows.Count; i++)
                {
                    //dtData.Columns.Add(Convert.ToString(ds.Tables["FieldDefinition"].Rows[i]["name"]));
                    columns.Add("\"" + Convert.ToString(ds.Tables["FieldDefinition"].Rows[i]["name"]).Replace("\"", "\"\"") + "\"");
                    //sb.Append("\"" + Convert.ToString(ds.Tables["FieldDefinition"].Rows[i]["name"]).Replace("\"", "\"\"") + "\"");
                }

                sb.AppendLine(string.Join(",", columns.ToArray()));
            }

            var xLanguageDetails = XElement.Parse(reportXML);

            clsXml.ReadXml(ref xLanguageDetails, ref ds, ref sb, false);

        }

        static void ReadFieldValue(XElement xUserNode, ref DataSet ds, ref DataTable dt, ref DataRow dr, List<Field> fields)
        {

            string sColumnName;
            int iIndex = 0;

            foreach (XElement xd in xUserNode.Elements("Field"))
            {
                string sFieldId = "";
                foreach (XAttribute xt in xd.Attributes("id"))
                {
                    sFieldId = xt.Value;
                }

                string sType = "";

                foreach (XAttribute xt1 in xd.Attributes("type"))
                {
                    sType = xt1.Value;
                }

                //DataRow[] drFieldsArray = ds.Tables["FieldDefinition"].Select("id = '" + sFieldId + "'");
                //sColumnName = Convert.ToString(drFieldsArray[0]["Name"]);

                sColumnName = fields.First(t => t.FieldID == sFieldId).FieldName;

                if (sType == "8")
                {
                    var users = (from tmp in xd.Descendants("User")
                                 select tmp.Attribute("lastName").Value + "," + tmp.Attribute("firstName").Value
                             ).ToArray();

                    var groups = (from tmp in xd.Descendants("Group")
                                  select tmp.Value
                            ).ToArray();

                    dr[sColumnName] = string.Join(";", users.Union(groups).ToArray());
                }
                else if (sType == "1")//text
                {
                    dr[sColumnName] = Common.RemoveHtml(xd.Value);
                }
                else if (sType == "9")//cross refer
                {
                    dr[sColumnName] = string.Join(";", xd.Elements("Reference").Select(t => t.Value).ToArray());
                }
                else
                {
                    dr[sColumnName] = xd.Value;
                }
                iIndex++;
            }
        }

        static List<string> ReadFields(XElement xUserNode, ref DataSet ds, List<Field> fields)
        {
            var columns = new List<string>();

            string sColumnName;
            int iIndex = 0;
            //int colIndex = -1;

            foreach (XElement xd in xUserNode.Elements("Field"))
            {
                string sFieldId = "";
                foreach (XAttribute xt in xd.Attributes("id"))
                {
                    sFieldId = xt.Value;
                }

                string sType = "";

                foreach (XAttribute xt1 in xd.Attributes("type"))
                {
                    sType = xt1.Value;
                }

                //DataRow[] drFieldsArray = ds.Tables["FieldDefinition"].Select("id = '" + sFieldId + "'");
                //sColumnName = Convert.ToString(drFieldsArray[0]["Name"]);

                sColumnName = fields.First(t => t.FieldID == sFieldId).FieldName;

                columns.Add(sColumnName);

                iIndex++;
            }

            //for (int i = colIndex + 1; i < dt.Columns.Count; i++)
            //{
            //    dr.ItemArray[i] = string.Empty;
            //}

            return columns;
        }

        //public static void DataTableToCsv(DataTable dt, string fileName)
        //{
        //    StringBuilder sb = new StringBuilder();

        //    var columnNames = dt.Columns.Cast<DataColumn>().Select(column => "\"" + column.ColumnName.Replace("\"", "\"\"") + "\"").ToArray();
        //    sb.AppendLine(string.Join(",", columnNames));

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        var fields = row.ItemArray.Select(field => "\"" + field.ToString().Replace("\"", "\"\"") + "\"").ToArray();
        //        sb.AppendLine(string.Join(",", fields));
        //    }

        //    string strFilePath = ConfigurationManager.AppSettings["outPutfolder"];
        //    string strArchivePath = ConfigurationManager.AppSettings["archivefolder"];
        //    var exportFreq = ConfigurationManager.AppSettings["outPutFrequency"].Split(',');

        //    //string strCompletePathName = "";
        //    //string strCompleteFileName = "";

        //    SFTP sftp = new SFTP();

        //     try
        //     {
        //         sftp.ConnectTosFTP();

        //         foreach (var freq in exportFreq)
        //         {
        //             try
        //             {
        //                 if (freq == "D")
        //                 {
        //                     string strCompleteFileName1 = fileName + "_Daily.csv";
        //                     string strCompletePathName1 = strFilePath + "\\" + strCompleteFileName1;
        //                     File.WriteAllText(strCompletePathName1, sb.ToString().Trim(), Encoding.UTF8);

        //                     sftp.UploadFile(strCompleteFileName1, strCompletePathName1, false);

        //                     string strCompleteFileName2 = fileName + "_Daily_" + DateTime.Now.DayOfWeek.ToString() + ".csv";
        //                     string  strCompletePathName2 = strFilePath + "\\" + strCompleteFileName2;
        //                     //File.WriteAllText(strCompletePathName2, sb.ToString().Trim(), Encoding.UTF8);
        //                     File.Copy(strCompletePathName1, strCompletePathName2, true);

        //                     sftp.UploadFile(strCompleteFileName2, strCompletePathName2, false);

        //                     string  strCompleteFileName3 = fileName + "_Daily_" + DateTime.Now.ToString("dd-MMM-yyyy") + ".csv";
        //                     string strCompletePathName3 = strArchivePath + "\\" + strCompleteFileName3;
        //                     //File.WriteAllText(strCompletePathName3, sb.ToString().Trim(), Encoding.UTF8);
        //                     File.Copy(strCompletePathName2, strCompletePathName3, true);

        //                     sftp.UploadFile(strCompleteFileName3, strCompletePathName3, true);

        //                     File.Delete(strCompletePathName1);
        //                     File.Delete(strCompletePathName2);
        //                     File.Delete(strCompletePathName3);

        //                     Logger.WriteLog("file generated :- " + strCompletePathName3);
        //                 }
        //                 else if (freq == "W")
        //                 {
        //                     var dayOfWeek = ConfigurationManager.AppSettings["WeekdayIfWeekly"].Split(',');

        //                     if (dayOfWeek.Contains(((int)DateTime.Now.DayOfWeek).ToString()))
        //                     {
        //                         string strCompleteFileName1 = fileName + "_Weekly.csv";
        //                         string strCompletePathName1 = strFilePath + "\\" + strCompleteFileName1;
        //                         File.WriteAllText(strCompletePathName1, sb.ToString().Trim(), Encoding.UTF8);

        //                         sftp.UploadFile(strCompleteFileName1, strCompletePathName1, false);

        //                         string strCompleteFileName2 = fileName + "_Weekly_" + DateTime.Now.ToString("dd-MMM-yyyy") + ".csv";
        //                         string strCompletePathName2 = strArchivePath + "\\" + strCompleteFileName2;
        //                         //File.WriteAllText(strCompletePathName2, sb.ToString().Trim(), Encoding.UTF8);
        //                         File.Copy(strCompletePathName1, strCompletePathName2, true);

        //                         sftp.UploadFile(strCompleteFileName2, strCompletePathName2, true);

        //                         File.Delete(strCompletePathName1);
        //                         File.Delete(strCompletePathName2);

        //                         Logger.WriteLog("file generated :- " + strCompletePathName2);
        //                     }
        //                 }
        //                 else if (freq == "M")
        //                 {
        //                     var dayOfMonth = ConfigurationManager.AppSettings["DateOfMonthIfMonthly"].Split(',');

        //                     if (dayOfMonth.Any(tmp => tmp == ((int)DateTime.Now.Day).ToString()))
        //                     {
        //                         string strCompleteFileName1 = fileName + "_Monthly.csv";
        //                         string strCompletePathName1 = strFilePath + "\\" + strCompleteFileName1;
        //                         File.WriteAllText(strCompletePathName1, sb.ToString().Trim(), Encoding.UTF8);

        //                         sftp.UploadFile(strCompleteFileName1, strCompletePathName1, false);

        //                         string strCompleteFileName2 = fileName + "_Monthly_" + DateTime.Now.ToString("dd-MMM-yyyy") + ".csv";
        //                         string strCompletePathName2 = strArchivePath + "\\" + strCompleteFileName2;
        //                         //File.WriteAllText(strCompletePathName2, sb.ToString().Trim(), Encoding.UTF8);
        //                         File.Copy(strCompletePathName1, strCompletePathName2, true);

        //                         sftp.UploadFile(strCompleteFileName2, strCompletePathName2, true);

        //                         File.Delete(strCompletePathName1);
        //                         File.Delete(strCompletePathName2);

        //                         Logger.WriteLog("file generated :- " + strCompletePathName2);
        //                     }
        //                 }
        //             }
        //             catch (Exception ex)
        //             {
        //                 Logger.WriteError(ex);
        //             }
        //         }
        //     }
        //     finally
        //     {
        //         sftp.Disconnect();
        //     }

        //     //return strCompletePathName;
        //}

        public static string RemoveHtml(string html)
        {
            //html = "136^\"Donald Allers\"^\"CBIT\"^\"CBIT Treasury IT\"^\"Sustain without Enhancements\"^\"Web Browser\"^\"Low  (under 50 mil)\"^\"Scrittura is a externally developed system which is utilized to " + Environment.NewLine + " #NAME? " + Environment.NewLine + "  #NAME?- provide a workflow solution for P&L variance resolution\"^\"In-house, on AIG premises\"^\"Core\"^\"Wilton / Connecticut / United States\"^\"Low (Supports internal process, if unavailable, no significant business impact, manual process can take over)\"^\"Settle Cl";
            html = StripTagsCharArray(html);

            html = Regex.Replace(html, "<.*?>", String.Empty);

            html = Regex.Replace(html, @"<[^>]+>|&nbsp;", "").Trim();
            html = Regex.Replace(html, @"\s{2,}", " ");

            html = Regex.Replace(html, "\\<[^\\>]*\\>", String.Empty);
            html = Regex.Replace(html, "<.*?>|&.*?;", string.Empty);
            html = Regex.Replace(html, "<[^>]*>", string.Empty);
            html = Regex.Replace(html, @"<(.|\n)*?>", string.Empty);

            html = html.Replace(System.Environment.NewLine, string.Empty);
            html = Regex.Replace(html, @"\t|\n|\r", String.Empty);

            //html = html.Replace("\"", "'");

            var start = html.IndexOf("<");

            if (start < 0)
                return html;

            var end = html.IndexOf(">", html.IndexOf("<"));

            if (end < 0)
                return html;

            var length = (end - start) + 1;

            while (!(start < 0) && end > 0 && length > 0)
            {
                html = html.Remove(start, length);

                start = html.IndexOf("<");

                if (start < 0)
                    break;

                end = html.IndexOf(">", html.IndexOf("<"));
                if (end < 0)
                    break;

                length = (end - start) + 1;
            }

            //return html.TrimStart().TrimEnd();


            //<!*[^<>]*>
            //<[^>]*>

            //<([\w:]+)>(\s|&nbsp;)*</\1>
            //<\\S[^><]*>

            return html;

        }

        static string StripTagsCharArray(string source)
        {
            //source = "<p>asfsafsafsaf </p><p/>";

            char[] array = new char[source.Length];
            int arrayIndex = 0;
            bool inside = false;

            for (int i = 0; i < source.Length; i++)
            {
                char let = source[i];
                if (let == '<')
                {
                    inside = true;
                    continue;
                }
                if (let == '>')
                {
                    inside = false;
                    continue;
                }
                if (!inside)
                {
                    array[arrayIndex] = let;
                    arrayIndex++;
                }
            }
            return new string(array, 0, arrayIndex);
        }

        public static void WriteLog(string dataInfo)
        {
            try
            {
                Console.WriteLine(dataInfo);
                var logFolder = ConfigurationManager.AppSettings["Logfolder"];

                Directory.CreateDirectory(logFolder);

                var fileName = ConfigurationManager.AppSettings["LogFileName"] + DateTime.Now.ToString("dd-MMM-yyyy");

                using (StreamWriter w = File.AppendText(logFolder + "\\" + fileName + ".txt"))
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
