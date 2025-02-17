using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.IO;
using System.Xml;


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
            string requestUri = "https://av000t2p.sap.tkelevator.com:44348/sap/bc/srt/rfc/sap/z_cv_read_sales_conf/010/z_cv_read_sales_conf/z_cv_read_salcv_binding";

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
    }

}


