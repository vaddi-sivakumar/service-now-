using ArcherCommonClass;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ArcherExtractQlikview
{
    class clsXml
    {
       public static void ReadXml(ref XElement element, ref DataSet ds, ref StringBuilder sb, bool isRecursive)
        {
            for (int i = 0; i < element.Elements("Record").Count(); i++)
            {
                var Node = element.Elements("Record").ElementAt(i);

                string sColumnName;
                var data = new List<string>();
               
                foreach (XElement xd in Node.Elements("Field"))
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

                    DataRow[] drFieldsArray = ds.Tables["FieldDefinition"].Select("id = '" + sFieldId + "'");
                    sColumnName = Convert.ToString(drFieldsArray[0]["Name"]);

                    if (sType == "8")
                    {
                        var users = (from tmp in xd.Descendants("User")
                                     select tmp.Attribute("lastName").Value + " " + tmp.Attribute("firstName").Value
                                 ).ToArray();

                        var groups = (from tmp in xd.Descendants("Group")
                                      select tmp.Value
                                ).ToArray();

                        data.Add(string.Join(";", users.Union(groups).ToArray()));
                    }
                    else if (sType == "1")//text
                    {
                        data.Add(Common.RemoveHtml(xd.Value).Replace(",", " "));
                    }
                    else if (sType == "9")//cross refer
                    {
                        data.Add(string.Join(";", xd.Elements("Reference").Select(t => t.Value.Replace(",", " ")).ToArray()));
                    }
                    else
                    {
                        data.Add(xd.Value.Replace(",", " "));
                    }
                }

                sb.Append(string.Join(",", data.ToArray()));

                if (Node.Elements("Record").Count() > 0)
                {
                    sb.Append(",");
                    ReadXml(ref Node, ref ds, ref sb, true);
                }
                else
                {
                    sb.AppendLine();                  
                }
                
            }
        }

    }
}
