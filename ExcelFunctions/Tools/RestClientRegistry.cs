using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelFunctions.Tools
{
    internal class RestClientRegistry
    {
        public enum Supplier
        {
            Fixer,
            BankNBP,
            BankNBRB
        }

        class ClientRecord
        {
            public string ServerAddress { get; set; }
            public RestClient RestClient { get; set; }
            public DateTime? LastRequest { get; set;}
            public DateTime? NextRequestAllowed { get; set; }
        }
        Dictionary<Supplier, ClientRecord> _clients = new Dictionary<Supplier, ClientRecord>();
        HttpClient _httpClient;

        static readonly TimeSpan _requestTimeout = new TimeSpan(0, 30, 0);
        public RestClientRegistry() 
        {
            //https://stackoverflow.com/questions/22251689/make-https-call-using-httpclient
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            _httpClient = new HttpClient();
        }

        public void Register(Supplier supplier, string baseUrl, bool useRestClient = true)
        {
            if (_clients.ContainsKey(supplier))
            {
                throw new InvalidOperationException();
            }
            var record = new ClientRecord
            { 
                RestClient = useRestClient ? new RestClient(baseUrl) : null, 
                ServerAddress = baseUrl 
            };
            _clients[supplier] = record;
        }

        public Dictionary<string, object> Call(Supplier supplier, string getParam)
        {
            var record = _clients[supplier];
            if (record.NextRequestAllowed != null && record.NextRequestAllowed.Value > DateTime.Now)
                return null;

            if(record.RestClient != null)
            {
                RestRequest request = new RestRequest(getParam, Method.Get);
                record.LastRequest = DateTime.Now;
                var resultRaw = record.RestClient.Execute<Object>(request).Data;
                if (resultRaw != null)
                {
                    var e = resultRaw as JsonElement?;
                    if (e != null)
                    {
                        return e.Value.ToObject<Dictionary<string, object>>();
                    }
                }
                Debug.WriteLine($"Call({supplier}) at {getParam} failed");
            }
            else
            {
                try
                {
                    var uriServer = new Uri(record.ServerAddress);
                    var url = new Uri(uriServer, getParam);
                    var response = _httpClient.SendAsync(new HttpRequestMessage(HttpMethod.Get, url)).Result;
                    var resultString = response.Content.ReadAsStringAsync().Result;
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(resultString);
                    return dict;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Call({supplier}) at {getParam} failed, exception: " +ex.Message);
                }
            }

            record.NextRequestAllowed = DateTime.Now + _requestTimeout;
            return null;
        }
    }
}
