using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ArcherGeneral;
using SNOW.GeneralService;
using SNOW.ArcherSearchService;
using System.Data;
using System.Xml.Linq;
using SNOW;
using System.Text.RegularExpressions;
using SNOW.ArcherRecordService;
using ArcherCommonClass;

namespace ArcherGeneral
{
    public static class Archer
    {

        public static string GetSession()
        {
            var archerUser = ConfigurationManager.AppSettings["archerUser"];
            var archerPwd = ConfigurationManager.AppSettings["archerPwd"];
            var archerInstance = ConfigurationManager.AppSettings["archerInstance"];

            var login = new general();
            var session = login.CreateUserSessionFromInstance(archerUser, archerInstance, archerPwd);
            //Logger.WriteLog("session id :- " + session);

            return session;
        }

        public static List<string> GetReportsGuids()
        {
            return ConfigurationManager.AppSettings["archerReportsGuids"].Split(',').ToList();
        }

        public static List<string> GetOutputFileNames()
        {
            return ConfigurationManager.AppSettings["outPutFileName"].Split(',').ToList();
        }




        public static string GetReportXML(string SessionID)
        {
            var reportId = ConfigurationSettings.AppSettings.Get("archerReportsGuids");

            var srch = new search();
            var resultXML = srch.SearchRecordsByReport(SessionID, reportId, 1);

            decimal decRecCount;
            //Getting total number of records in report
            using (var xmlReader = XmlReader.Create(new StringReader(resultXML)))
            {
                xmlReader.MoveToContent();
                var strRecCount = xmlReader.GetAttribute("count");
                //Console.WriteLine("Total number of records: " + strRecCount);
                decRecCount = Convert.ToDecimal(strRecCount);
            }

            //Getting total number of records per page
            var readDoc = new XmlDocument();
            readDoc.LoadXml(resultXML);
            decimal decRecPerPage = readDoc.SelectNodes("/Records/Record").Count;
            decimal decTotalPageCount = 0;

            if (decRecPerPage > 0)
            {
                //Total Pagecount
                decTotalPageCount = decRecCount / decRecPerPage;
            }
            else
                decTotalPageCount = decRecCount;

            string strResultXML = string.Empty;

            for (int intCount = 1; intCount <= Math.Ceiling(decTotalPageCount); intCount++)
            {
                strResultXML = srch.SearchRecordsByReport(SessionID, reportId, intCount);

                if (intCount == 1)
                {
                    resultXML = strResultXML;
                }
                else
                {
                    var doc = new XmlDocument();
                    doc.LoadXml(strResultXML);

                    //Select and display the value of all the records
                    XmlNodeList nodeList;
                    XmlElement root = doc.DocumentElement;
                    nodeList = root.SelectNodes("/Records/Record");

                    resultXML = resultXML.Replace("</Records>", "");
                    foreach (XmlNode node in nodeList)
                    {
                        resultXML += node.OuterXml;
                    }

                    resultXML += "</Records>";
                }
            }
            return resultXML;
        }

        public static System.Data.DataTable xmlTOdt(string reportXML)
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
            //dtData.Columns.Add("RefrenceId");

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
                    dr[sColumnName] = RemoveHtml(xd.Value);
                }
                else if (sType == "9")//cross refer
                {
                    dr[sColumnName] = string.Join(";", xd.Elements("Reference").Select(t => t.Value).ToArray());
                    //dr[sColumnName] = string.Join(";", xd.Elements("id").Select(t => t.Value).ToArray());


                    dr["RefrenceId"] = string.Join(";", xd.Elements("Reference").Attributes("id").Select(t => t.Value).ToArray());

                }
                else
                {
                    dr[sColumnName] = xd.Value;
                }
                iIndex++;
            }
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

        public static void UpdateIncidentNumberArcher(string incNumber, string SessionID, string TrackingId)
        {
            try
            {
                var rb = new StringBuilder();
                var rcord = new record();

                int ModuleId = Convert.ToInt16(ConfigurationSettings.AppSettings.Get("archerVulnerabilityAppId"));
                int FieldId = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("archerIncindetId"));
                int ValueListId = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("archerValueListtId"));
                int serviceTimeout = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("archerserviceTimeout"));

                int ContentId = Convert.ToInt32(TrackingId);

                rcord.Timeout = serviceTimeout;

                // update the incident number field 
                rb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
                rb.AppendLine("<Record>");
                rb.AppendLine("<Field id=\"" + FieldId + "\" value=\"" + incNumber + "\" >");
                rb.AppendLine("</Field>");
                rb.AppendLine("<Field id=\"" + ValueListId + "\" value=\"" + ConfigurationSettings.AppSettings.Get("archerValueListtValue") + "\" >");
                rb.AppendLine("</Field>");
                rb.AppendLine("</Record>");

                if (rb != null)
                {
                    rcord.UpdateRecord(SessionID, ModuleId, ContentId, rb.ToString());
                    
                    Common.WriteLog("Incident :" + incNumber + " is created for Tracking id" + TrackingId);
                }
            }
            catch (Exception ex)
            {
                Common.WriteLog("Exception while record update in archer :" + ex);
            }

            //WriteLog("Record Id ==> " + ContentId + " has been Updated");
        }

        public static void UpdateIncidentInfoArcher(string TrackingId, DataRow dataRow, string SessionID)
        {
            var rb = new StringBuilder();
            var rcord = new record();
            
            int ModuleId = Convert.ToInt16(ConfigurationSettings.AppSettings.Get("archerVulnerabilityAppId"));
            int IncComments = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("archerInccomments"));
            int IncClosedAt = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("archerIncClosed_at"));
            int serviceTimeout = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("archerserviceTimeout"));

            int ContentId = Convert.ToInt32(TrackingId);

            rcord.Timeout = serviceTimeout;

            // update the incident number field 
            rb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            rb.AppendLine("<Record>");
            rb.AppendLine("<Field id=\"" + IncComments + "\" value=\"" + dataRow["close_notes"].ToString() + "\" >");
            rb.AppendLine("</Field>");
            rb.AppendLine("<Field id=\"" + IncClosedAt + "\" value=\"" + dataRow["closed_at"].ToString() + "\" >");
            rb.AppendLine("</Field>");
            rb.AppendLine("</Record>");

            if (rb != null)
            {
                rcord.UpdateRecord(SessionID, ModuleId, ContentId, rb.ToString());

                Common.WriteLog("Traking Id: " + TrackingId + "and Incident number: " + dataRow["number"].ToString() + "Comments has been updated in Archer");
            }
        }
    }
}
