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
    public class IROUTE
    {
        public static IRestResponse EnviarGuia(List<Entrega> entrega)
        {
            try
            {
                var objJson = JsonConvert.SerializeObject(entrega);              
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.UrlRoute);
                var request = new RestRequest("registrar-masivo", Method.POST);
                request.AddHeader("content-type", "application/json");
                request.AddParameter("application/json", objJson, ParameterType.RequestBody);
                string credentials = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{Globals.UserRoute}:{Globals.PasswordRoute}"));
                request.AddHeader("Authorization", $"Basic {credentials}");
                return client.Execute(request);
            }
            catch (Exception ex)
            {
                dynamic errorMsj = JObject.Parse(ex.Message.Replace("'", ""));
                throw new Exception(errorMsj.error.message.value);
            }
        }
    }
}
