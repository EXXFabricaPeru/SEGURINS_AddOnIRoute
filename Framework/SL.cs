using Integration_IROUTE.App;
using Integration_IROUTE.Entity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Integration_IROUTE.Framework
{
    public class SL
    {
        public static string sConnectionContext = null;
        public static string serviceLayerAddress = null;
        public static SLLogin SLLoginResponse;

        public static void Connect()
        {
            try
            {
                //string server = Globals.oCompany.Server.Replace("NDB@", "").Substring(0, Globals.oCompany.Server.Replace("NDB@", "").IndexOf(":")).Trim();
                string serviceLayerAddress = Globals.ObtenerSL();
                //string serviceLayerAddress = "https://172.16.210.12:50000/b1s/v1";
                string sConnectionContextAux = null;

                try
                {
                    sConnectionContextAux = Globals.SBO_Application.Company.GetServiceLayerConnectionContext(serviceLayerAddress);
                }
                catch (System.Exception ex)
                {
                }

                if (sConnectionContextAux == null)
                    throw new Exception("No se logró establecer conexión con Service Layer");

                sConnectionContext = sConnectionContextAux;
                SL.serviceLayerAddress = serviceLayerAddress;
                SLLoginResponse = new SLLogin();
                SLLoginResponse.B1SESSION = sConnectionContext.Split(';')[0].Replace("B1SESSION=", "");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public static IRestResponse ObtenerGuia(string DocEntry)
        {
        band:
            try
            {
                if (serviceLayerAddress == null) Connect();
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(serviceLayerAddress);
                var request = new RestRequest("DeliveryNotes(" + DocEntry + ")", Method.GET);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", SLLoginResponse.B1SESSION);
                //request.AddCookie("ROUTEID", ".node0");
                return client.Execute(request);
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    serviceLayerAddress = null;
                    goto band;
                }

                dynamic errorMsj = JObject.Parse(ex.Message.Replace("'", ""));
                throw new Exception(errorMsj.error.message.value);
            }
        }
    }
}
