using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Newtonsoft.Json;

namespace EPLAN_API.SAP
{
    public class SAPReader
    {
        public const String _URL_Config_Elec = "http://95.111.0.21:8080/pentaho/api/repos/%3Apublic%3ATKN_BI%3ATKN_CARACTERISTICAS_CONF_ELECTRICO.wcdf/generatedContent";

        //private Electric oElectric;
        //private Dictionary<string, string> SAPCararct;
        private string OE;

        public SAPReader(string OE)
        {
            this.OE = OE;
        }

        public Dictionary<string, string> readCaracConfigElec()
        {
            //Caracteristicas configurador
            try
            {
                CookieContainer cookies = new CookieContainer();
                HttpWebRequest request = null;
                HttpWebResponse response = null;
                string returnData = string.Empty;

                //Need to retrieve cookies first
                request = (HttpWebRequest)WebRequest.Create(new Uri(_URL_Config_Elec));
                request.Method = "GET";
                request.CookieContainer = cookies;
                using (response = (HttpWebResponse)request.GetResponse())
                {
                    //Set up the request
                    request = (HttpWebRequest)WebRequest.Create(new Uri("http://95.111.0.21:8080/pentaho/plugin/cda/api/doQuery?"));
                    request.Method = "POST";
                    request.Host = "druida:8080";
                    request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:85.0) Gecko/20100101 Firefox/85.0";
                    request.Accept = "*/*";
                    request.ContentType = "application/x-www-form-urlencoded;charset=UTF-8";
                    request.KeepAlive = true;
                    request.Referer = _URL_Config_Elec;
                    request.CookieContainer = cookies;


                    //Format the POST data
                    StringBuilder postData = new StringBuilder();
                    postData.Append(String.Concat("paramp_ordenes=", OE));
                    postData.Append(String.Concat("&paramp_ordenes_sinformato=", OE));
                    postData.Append("&path=/public/TKN_BI/TKN_CARACTERISTICAS_CONF_ELECTRICO.cda");
                    postData.Append("&dataAccessId=joined_query_all");
                    postData.Append("&outputIndexId=1");
                    postData.Append("&pageSize=0");
                    postData.Append("&pageStart=0");
                    postData.Append("&sortBy=0A");
                    postData.Append("&paramsearchBox=");

                    request.ContentLength = postData.Length;


                    //write the POST data to the stream
                    using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
                        writer.Write(postData.ToString());
                }
                using (response = (HttpWebResponse)request.GetResponse())

                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                    returnData = reader.ReadToEnd();

                Root rootData = JsonConvert.DeserializeObject<Root>(returnData);

                Dictionary<string, string> resultSet = rootData.GetFormattedResultSet()[0].AllGroupValues;

                return resultSet;
            }

            catch (Exception e)
            {
                var result = MessageBox.Show("No se puede conectar con el servidor\n¿quieres buscar un archivo de datos?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    OpenFileDialog openFileDialog1 = new OpenFileDialog();

                    openFileDialog1.InitialDirectory = "C:\\Users\\DARG\\Dropbox\\TKN\\2_EPLAN\\API\\";
                    openFileDialog1.Filter = "Text Files (*.txt)|*.txt";
                    openFileDialog1.FilterIndex = 0;
                    openFileDialog1.RestoreDirectory = true;

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        using (var reader = new StreamReader(openFileDialog1.FileName))
                        {
                            String text = reader.ReadToEnd();
                            String[] spearator = { "<br>", "\\\",\\\"&", "nbsp" };
                            String[] strlist = text.Split(spearator, StringSplitOptions.None);
                            return null;
                        }
                    }
                }

                throw;

            }


        }


    }

    public class Metadata
    {
        public string ColName { get; set; }
        public string ColType { get; set; }
        public int ColIndex { get; set; }
    }

    public class QueryInfo
    {
        public string TotalRows { get; set; }
    }

    public class ResultSetRow
    {
        public int Orden { get; set; }
        public string OE_OF { get; set; }
        public string UnidadDeObra { get; set; }
        public Dictionary<string, Dictionary<string, string>> Groups { get; set; } = new Dictionary<string, Dictionary<string, string>>();

        public string Observaciones { get; set; }

        public Dictionary<string, string> AllGroupValues
        {
            get
            {
                var allValues = new Dictionary<string, string>();

                // Iterate through each group in Groups dictionary
                foreach (var group in Groups)
                {
                    foreach (var item in group.Value)
                    {
                        // Add each key-value pair to the combined dictionary
                        allValues[item.Key] = item.Value;
                    }
                }

                return allValues;
            }
        }
    }

    public class Root
    {
        public List<Metadata> Metadata { get; set; }
        public List<List<object>> Resultset { get; set; }
        public QueryInfo QueryInfo { get; set; }

        public List<ResultSetRow> GetFormattedResultSet()
        {
            var formattedResultSet = new List<ResultSetRow>();

            foreach (var row in Resultset)
            {
                var formattedRow = new ResultSetRow
                {
                    Orden = Convert.ToInt32(row[0]),
                    OE_OF = row[1]?.ToString(),
                    UnidadDeObra = row[2]?.ToString(),
                    Observaciones = row[8]?.ToString()
                };

                // Parse and add each group to the Groups dictionary
                formattedRow.Groups["Unidad de Obra"] = GroupParser.ParseUnidadDeObraData(row[2]?.ToString() ?? string.Empty);
                formattedRow.Groups["Grupo A"] = GroupParser.ParseGroupData(row[3]?.ToString() ?? string.Empty);
                formattedRow.Groups["Grupo B"] = GroupParser.ParseGroupData(row[4]?.ToString() ?? string.Empty);
                formattedRow.Groups["Grupo C"] = GroupParser.ParseGroupData(row[5]?.ToString() ?? string.Empty);
                formattedRow.Groups["Grupo D"] = GroupParser.ParseGroupData(row[6]?.ToString() ?? string.Empty);
                formattedRow.Groups["Grupo E"] = GroupParser.ParseGroupData(row[7]?.ToString() ?? string.Empty);

                formattedResultSet.Add(formattedRow);
            }

            return formattedResultSet;
        }
    }

    public class GroupParser
    {
        public static Dictionary<string, string> ParseGroupData(string groupData)
        {
            var parsedData = new Dictionary<string, string>();

            // Remove all occurrences of "&nbsp"
            groupData = groupData.Replace("&nbsp", string.Empty);

            // Split the data by <br> tags to get each key-value entry
            var entries = groupData.Split(new string[] { "<br>" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var entry in entries)
            {
                // Further split each entry by ": " to separate key and value
                var keyValue = entry.Split(new string[] { ":" }, StringSplitOptions.None);

                if (keyValue.Length == 2)
                {
                    string key = keyValue[0].Trim();
                    string value = keyValue[1].Trim().Split(' ')[0];
                    if(value.Equals("#N/A"))
                        value=null;

                    // Add the parsed key-value pair to the dictionary
                    parsedData[key] = value;
                }
            }

            return parsedData;
        }

        public static Dictionary<string, string> ParseUnidadDeObraData(string groupData)
        {
            var parsedData = new Dictionary<string, string>();

            // Remove all occurrences of "&nbsp"
            groupData = groupData.Replace("&nbsp", string.Empty);

            // Split the data by <br> tags to get each key-value entry
            var entries = groupData.Split(new string[] { "<br>" }, StringSplitOptions.RemoveEmptyEntries);
            var escalatorData = entries[1].Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);

            //OE Name TNCR_COM_NOMBREOBRA_VBACK
            parsedData["TNCR_COM_NOMBREOBRA_VBACK"] = entries[0];

            //Model FMODELL
            if (escalatorData[0].Contains("ORINOCO"))
            {
                parsedData["FMODELL"] = "ORINOCO";
            }
            else if (escalatorData[0].Contains("TUGELA CL"))
            {
                parsedData["FMODELL"] = "TUGELA_CLASSIC";
            }
            else if (escalatorData[0].Contains("VELINO CL"))
            {
                parsedData["FMODELL"] = "VELINO_CLASSIC";
            }
            else if (escalatorData[0].Contains("TUGELA"))
            {
                parsedData["FMODELL"] = "TUGELA";
            }
            else if (escalatorData[0].Contains("VELINO"))
            {
                parsedData["FMODELL"] = "VELINO";
            }
            else if (escalatorData[0].Contains("IWALK"))
            {
                parsedData["FMODELL"] = "IWALK";
            }
            else if (escalatorData[0].Contains("IMOD"))
            {
                parsedData["FMODELL"] = "IMOD";
            }
            else if (escalatorData[0].Contains("VICTORIA"))
            {
                parsedData["FMODELL"] = "VICTORIA";
            }

            //Width
            parsedData["FBREITE"] = escalatorData[1].Replace("EK","");

            //Angle
            parsedData["FNEIGUNG"] = escalatorData[2].Trim('º');

            //High
            parsedData["FHOEHEV"] = entries[2].Split(':')[1];


            return parsedData;
        }
    }

}
