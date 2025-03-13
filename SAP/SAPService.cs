using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.IO;
using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Drawing;
using System.Security.Policy;
using System.Linq;
using System.Xml.Linq;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using EPLAN_API.SAP;
using static System.Net.WebRequestMethods;
using Eplan.EplApi.DataModel;


namespace EPLAN_API.User
{

    public class SAPService
    {
        private const string WSDL_URL = "http://av000t2p.sap.tkelevator.com:8048/sap/bc/srt/wsdl/bndg_595E15C8FB432AD8E10000000F5B4E77/wsdl11/allinone/ws_policy/document?sap-client=010";
        private const string SAP_USERNAME = "SRV_NORTE";
        private const string SAP_PASSWORD = "Hola@2023";

        public SAPService()
        {

        }

        public Dictionary<string, string> ReadSAPCaract(string OE)
        {
            string pos = "10";
            int iOE = int.Parse(OE);

            if ((iOE > 1130000000 && iOE < 1150000000) || (iOE > 1160000000 && iOE < 1170000000))
            {
                pos = "20";
            }

            // El contenido del mensaje SOAP
            string soapRequest = $@"<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/'
                                     xmlns:urn='urn:sap-com:document:sap:soap:functions:mc-style'>
                                    <soapenv:Header/>
                                       <soapenv:Body>
                                          <urn:ZCvReadSalesConf>
                                             <IDocumentNumber>{OE}</IDocumentNumber>
                                             <IItemNumber>{pos}</IItemNumber>
                                             <ILang>S</ILang>
                                          </urn:ZCvReadSalesConf>
                                       </soapenv:Body>
                                    </soapenv:Envelope>";

            // URL del endpoint
            string requestUri = "https://u000t2p.sap.tkelevator.com/sap/bc/srt/rfc/sap/z_cv_read_sales_conf/010/z_cv_read_sales_conf/z_cv_read_salcv_binding";
            //string requestUri = "https://av000t2p.sap.tkelevator.com:44348/sap/bc/srt/rfc/sap/z_cv_read_sales_conf/010/z_cv_read_sales_conf/z_cv_read_salcv_binding";

            // Crear una solicitud HTTP
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUri);
            request.Method = "POST";
            request.ContentType = "text/xml";
            request.Headers.Add("SOAPAction", "urn:sap-com:document:sap:soap:functions:mc-style/ZCvReadSalesConf"); // Asegúrate de que esta SOAPAction es la correcta

            // Agregar encabezados para la autenticación básica
            string username = "SRV_NORTE";  // Reemplaza con tu nombre de usuario
            string password = "Hola@2023";  // Reemplaza con tu contraseña
            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(username + ":" + password));
            request.Headers["Authorization"] = "Basic " + credentials;

            // Escribir el contenido del mensaje SOAP en la solicitud
            using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
            {
                writer.Write(soapRequest);
            }

            try
            {
                string responseBody;
                // Obtener la respuesta de la solicitud
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    responseBody = reader.ReadToEnd();
                    //return responseBody;
                }

                Dictionary<string, string> res = ParseSoapResponseToDictionary(responseBody);
                res["TNCR_COM_COD_PEDIDO_CLIENTE"] = OE;
                return res;

            }
            catch (WebException ex)
            {
                // Si hay un error en la respuesta, leer el error
                using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
                {
                    string errorResponse = reader.ReadToEnd();
                    //return $"Error: {ex.Message}\n{errorResponse}";
                }
                return null;
            }
        }

        public string ReadSAPBOM(string OE, Project oProject)
        {
            string pos = "10";
            int iOE = int.Parse(OE);

            if ((iOE > 1130000000 && iOE < 1150000000) || (iOE > 1160000000 && iOE < 1170000000))
            {
                pos = "20";
            }

            string soapRequest = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                                <soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:urn=""urn:sap-com:document:sap:soap:functions:mc-style"">
                                   <soapenv:Header/>
                                   <soapenv:Body>
                                      <urn:ZCvReadSalesBom>
                                         <ILang>S</ILang>
                                         <Posnr>{pos}</Posnr>
                                         <Vbeln>{OE}</Vbeln>
                                      </urn:ZCvReadSalesBom>
                                   </soapenv:Body>
                                </soapenv:Envelope>";


            // URL del endpoint basado en el nuevo WSDL
            string requestUri = "https://u000t2p.sap.tkelevator.com/sap/bc/srt/rfc/sap/zws_read_sales_bom/010/zws_read_sales_bom/zws_read_sales_bom_binding";

            // Crear una solicitud HTTP
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUri);
            request.Method = "POST";
            request.ContentType = "text/xml;charset=UTF-8";
            request.Headers.Add("SOAPAction", "\"\"");  // SOAPAction vacío
            request.KeepAlive = true;
            request.UserAgent = "Apache-HttpClient/4.5.5 (Java/16.0.2)";
            request.Accept = "gzip,deflate";

            // Agregar encabezados para la autenticación básica
            string username = "SRV_NORTE";  // Reemplaza con tu nombre de usuario
            string password = "Hola@2023";  // Reemplaza con tu contraseña
            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(username + ":" + password));
            request.Headers["Authorization"] = "Basic " + credentials;

            // Escribir el contenido del mensaje SOAP en la solicitud
            using (StreamWriter writer = new StreamWriter(request.GetRequestStream()))
            {
                writer.Write(soapRequest);
            }

            try
            {
                string responseBody;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    responseBody = reader.ReadToEnd();
                }

                // Escribir el contenido de la respuesta en un archivo
                string filePath = oProject.ProjectDirectoryPath + "\\CHECK";
                if (!Directory.Exists(filePath))
                {
                    // Crear la carpeta si no existe
                    Directory.CreateDirectory(filePath);
                }
                filePath = filePath + "\\response.txt";
                //string filePath = "C:\\Users\\diego.alvarez\\Desktop\\response.txt"; // Aquí puedes especificar la ruta donde deseas guardar el archivo

                return responseBody;
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        string errorResponse = reader.ReadToEnd();
                        Console.WriteLine("Error Response: " + errorResponse);
                    }
                }
                Console.WriteLine("Exception: " + ex.Message);

                return null;
            }

        }

        public Dictionary<string, string> ParseSoapResponseToDictionary(string soapResponse)
        {
            // Crear el diccionario donde se almacenarán los pares clave-valor
            Dictionary<string, string> result = new Dictionary<string, string>();

            // Cargar el XML
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(soapResponse);

            // Buscar todos los nodos "item"
            XmlNodeList itemNodes = xmlDoc.GetElementsByTagName("item");

            // Recorrer los nodos y extraer "Atnam" y "Atwrt" para llenar el diccionario
            foreach (XmlNode item in itemNodes)
            {
                string atnam = item["Atnam"]?.InnerText;
                string atwrt = item["Atwrt"]?.InnerText;

                if (!string.IsNullOrEmpty(atnam) && !string.IsNullOrEmpty(atwrt))
                {
                    // Agregar la pareja clave-valor al diccionario
                    result[atnam] = atwrt;
                }
            }

            return result;
        }

        public string ParseBOM(string xmlData, Project oProject)
        {
            XDocument doc = XDocument.Parse(xmlData);

            // Crear la nueva estructura XML
            XElement etbom = new XElement("EtBom");

            // Diccionario para rastrear los últimos elementos por nivel
            Dictionary<int, XElement> lastElementsByLevel = new Dictionary<int, XElement>();
            lastElementsByLevel[0] = etbom; // Nivel raíz

            // Recorrer los elementos <item> manteniendo la estructura jerárquica
            foreach (XElement item in doc.Descendants("item"))
            {
                int stufe = int.Parse(item.Element("Stufe")?.Value ?? "0");

                // Buscar el padre correcto en el nivel anterior
                XElement parent = lastElementsByLevel.ContainsKey(stufe - 1) ? lastElementsByLevel[stufe - 1] : etbom;

                // Crear una copia del <item> sin modificar el original
                XElement newItem = new XElement("item", item.Elements());

                // Agregar el item al padre correcto
                parent.Add(newItem);

                // Actualizar el diccionario con el nuevo nivel
                lastElementsByLevel[stufe] = newItem;
            }

            // Guardar el nuevo XML estructurado en un archivo
            string outputPath = oProject.ProjectDirectoryPath + "\\CHECK";
            if (!Directory.Exists(outputPath))
            {
                // Crear la carpeta si no existe
                Directory.CreateDirectory(outputPath);
            }
            outputPath = outputPath + "\\response_nested.xml";
            //string outputPath = "C:\\Users\\diego.alvarez\\Desktop\\response_nested.xml";
            etbom.Save(outputPath);

            Debug.WriteLine($"XML generado correctamente: {outputPath}");

            return etbom.ToString();
        }

        public Dictionary<string, Cable> BuscaCables(string xmlData)
        {
            XDocument doc = XDocument.Parse(xmlData);

            // Filtrar los elementos <item> que cumplan con las condiciones:
            // 1. El valor de <Ojtxp> comienza con "cable" (sin importar mayúsculas/minúsculas).
            // 2. El valor de <Ojtxp> no contiene " 1x".
            // 3. No es una lista de materiales
            var matchingItems = doc.Descendants("item")
                                    .Where(item => item.Element("Ojtxp") != null)
                                    .Where(item => item.Element("Ojtxp").Value.StartsWith("cable", StringComparison.OrdinalIgnoreCase) &&
                                                   !item.Element("Ojtxp").Value.Contains(" 1x"))
                                    .Where(item => !item.Elements("item").Any())
                                    .ToList();

            Dictionary<string, Cable> cables = new Dictionary<string, Cable>();
            int error = 1;
            // Mostrar los resultados
            foreach (var item in matchingItems)
            {

                string pattern = @"[WCP]\d+(\.[A-Za-z0-9])?";
                MatchCollection matchesT1 = Regex.Matches(item.Element("Potx1").Value, pattern);
                MatchCollection matchesT2 = Regex.Matches(item.Element("Potx2").Value, pattern);

                //Lenght
                double.TryParse(item.Element("Mngko").Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double lenght);

                //Name
                string name = item.Element("Ojtxp").Value;

                //Code
                string code = item.Element("Idnrk").Value;

                //Parent Name
                string parentname = item.Parent.Element("Ojtxp").Value;

                //Parent Code
                string parentcode = item.Parent.Element("Idnrk").Value;

                //Pos inside parent
                string posinparent = item.Element("Posnr").Value;


                //Tiene el nombre del cable en su primera pos de texto
                if (matchesT1.Count > 0)
                {
                    string IME = matchesT1[0].Value;
                    if (!IME.StartsWith("W", StringComparison.OrdinalIgnoreCase))
                        IME = "W" + IME;
                    cables.Add(IME, new Cable(IME, name, code, "-", lenght, parentname, parentcode, posinparent));
                }

                //Tiene el nombre del cable en su segunda pos de texto
                else if (matchesT2.Count > 0)
                {
                    string IME = matchesT2[0].Value;
                    if (!IME.StartsWith("W", StringComparison.OrdinalIgnoreCase))
                        IME = "W" + IME;
                    cables.Add(IME, new Cable(IME, name, code, "-", lenght, parentname, parentcode, posinparent));
                }

                else
                {
                    //Tiene el nombre en su nodo solo si la pos es la misma
                    if (item.NextNode != null)
                    {
                        if (((XElement)item.NextNode).Element("Posnr").Value == item.Element("Posnr").Value)
                        {
                            if (((XElement)item.NextNode).Elements("Potx1").Any())
                            {
                                MatchCollection matchesNext = Regex.Matches(((XElement)item.NextNode).Element("Potx1").Value, pattern);
                                if (matchesNext.Count > 0)
                                {
                                    string IME = matchesNext[0].Value;
                                    if (!IME.StartsWith("W", StringComparison.OrdinalIgnoreCase))
                                        IME = "W" + IME;
                                    cables.Add(IME, new Cable(IME, name, code, "-", lenght, parentname, parentcode, posinparent));
                                    continue;
                                }
                            }
                        }
                    }

                    //Tiene el nombre en su nodo anterior
                    if (item.PreviousNode != null)
                    {
                        if (((XElement)item.PreviousNode).Elements("Potx1").Any())
                        {
                            MatchCollection matchesPrevious = Regex.Matches(((XElement)item.PreviousNode).Element("Potx1").Value, pattern);
                            if (matchesPrevious.Count > 0)
                            {
                                string IME = matchesPrevious[0].Value;
                                if (!IME.StartsWith("W", StringComparison.OrdinalIgnoreCase))
                                    IME = "W" + IME;
                                cables.Add(IME, new Cable(IME, name, code, "-", lenght, parentname, parentcode, posinparent));
                                continue;
                            }
                        }
                    }

                    //es desconocido
                    string nameerr = "UK_" + error.ToString();
                    error += 1;
                    cables.Add(nameerr, new Cable(nameerr, name, code, "-", lenght, parentname, parentcode, posinparent));
                    //Debug.WriteLine("problema en cable:" + name + " que esta dentro de " + parentname);

                }
            }

            return cables;
        }

    }

}


