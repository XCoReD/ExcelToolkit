using Newtonsoft.Json;
using RestSharp;
using RestSharp.Serialization.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Tools;

namespace ExcelFunctions
{
    public class ExchangeRate
    {
        //RestClient _restClientNBRB = new RestClient("https://www.nbrb.by/");
        RestClient _restClientExchangeRates, _restClientExchangeRatesNBP;
        string _apiKey = "5db46eaf8fadb8c67faa0fbb9cad5595";
        HttpClient _client;

        static Dictionary<string, double> _ratesDictionnary;
        static ObjectSerializer<Dictionary<string, double>> _ratesSerializer;
        static int _ratesAdded = 0;
        //API password
        //asdfk2@#3rfLQsdf3
        /*
         * 
         * Your API Key: 5db46eaf8fadb8c67faa0fbb9cad5595
         * http://data.fixer.io/api/
         * Example Request: http://data.fixer.io/api/latest?access_key=5db46eaf8fadb8c67faa0fbb9cad5595
         * */
        public ExchangeRate()
        {
            if (_ratesSerializer == null)
                _ratesSerializer = new ObjectSerializer<Dictionary<string, double>>("ExcelToolkit", "ExchangeRates");

            if (_ratesDictionnary == null)
            {
                _ratesDictionnary = _ratesSerializer.GetSerializedObject();
                if(_ratesDictionnary  == null)
                    _ratesDictionnary = new Dictionary<string, double>();
            }
            _restClientExchangeRates = new RestClient("http://data.fixer.io");  //no https access in the base plan

            //https://stackoverflow.com/questions/22251689/make-https-call-using-httpclient
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            _client = new HttpClient();
        }

        ~ExchangeRate()
        {
            if(_ratesAdded != 0 && _ratesSerializer != null)
            {
                _ratesSerializer.SaveSerializedObject(_ratesDictionnary);
                _ratesSerializer = null;
            }
        }

        string GetHash(DateTime date, string currency, string baseCurrency)
        {
            return $"{date.Year}-{date.Month}-{date.Day}#{currency}#{baseCurrency}";
        }
        public double GetExchangeRate(DateTime date, string currency, string baseCurrency)
        {
            double result = 0.0;
            string key = GetHash(date, currency, baseCurrency);
            if (_ratesDictionnary.TryGetValue(key, out result))
                return result;

            using (EasyLog log = new EasyLog("ExcelToolkit"))
            {
                string dateString = date.ToString("yyyy-MM-dd");
                string notes = null;

                //https://data.fixer.io/api/YYYY-MM-DD?access_key=YOUR_ACCESS_KEY

                RestRequest request = new RestRequest($"/api/{dateString}?access_key={_apiKey}", Method.GET);

                var resultRaw = _restClientExchangeRates.Execute<Object>(request).Data;
                if (resultRaw != null)
                {
                    var dict = resultRaw as Dictionary<string, object>;
                    object paramValue;
                    if (dict.TryGetValue("base", out paramValue))
                    {
                        string serverBaseCurrency = paramValue as string;

                        object fieldsRaw;
                        if (dict.TryGetValue("rates", out fieldsRaw))
                        {
                            var dictRates = fieldsRaw as Dictionary<string, object>;
                            object rateRaw, rateRawBase;
                            if (dictRates.TryGetValue(currency, out rateRaw) && dictRates.TryGetValue(baseCurrency, out rateRawBase))
                            {
                                double rate = ToDouble(rateRaw);
                                double rateBase = ToDouble(rateRawBase);

                                result = rate / rateBase;

                                _ratesDictionnary.Add(key, result);
                                ++_ratesAdded;
                            }
                            else
                                notes = $"currency {currency} or {baseCurrency} is not found in the rates array";
                        }
                        else
                            notes = "rates section is not found";
                    }
                    else
                    {
                        notes = "coulnd't find base currency";
                    }
                }
                else
                    notes = "REST response is empty";

                //handling exchangerate.io - old case
                /*if (resultRaw != null)
                {
                    var dict = resultRaw as Dictionary<string, object>;
                    object fieldsRaw;
                    if (dict.TryGetValue("rates", out fieldsRaw))
                    {
                        var dictRates = fieldsRaw as Dictionary<string, object>;
                        object rateRaw;
                        if (dictRates.TryGetValue(currency, out rateRaw))
                        {
                            //self-test
//#if DEBUG
                            if (dict.TryGetValue("base", out fieldsRaw))
                            {
                                Debug.Assert(fieldsRaw as string == baseCurrency);
                                if (dict.TryGetValue("date", out fieldsRaw))
                                {
                                    var s = fieldsRaw as string;
                                    if(s != dateString)
                                    {
                                        //actual rate may be from the day before, it is OK
                                        DateTime actual;
                                        if(DateTime.TryParse(s, out actual))
                                        {
                                            if(actual > date || (actual - date).TotalDays > 2)
                                            {
                                                notes = "returned rate for date " + actual.ToString();
                                            }
                                        }
                                        else
                                        {
                                            notes = "returned date which cannot be parsed: " + s;
                                        }
                                    }
                                }
                            }
//#endif
                            result = (double)rateRaw;
                        }
                        else
                            notes = "currency is not found in the rates array in the REST response";
                    }
                    else
                        notes = "rates section is not found in the REST response";

                }
                else
                    notes = "REST response is empty";
                */

                log.Info($"DT.GetExchangeRate: date ({dateString}), currency ({currency}), baseCurrency ({baseCurrency}), result ({result}), notes ({notes})");
                log.Flush();
            }

            return result;
        }

        public double GetExchangeRateNBP(DateTime date, string currency)
        {
            double result = 0.0;
            string type = "a";  //medium rate
            string key = GetHash(date, currency, "NBP-" + type);
            if (_ratesDictionnary.TryGetValue(key, out result))
                return result;

            using (EasyLog log = new EasyLog("ExcelToolkit"))
            {
                string dateString = date.ToString("yyyy-MM-dd");
                string notes = null;

                if (_restClientExchangeRatesNBP == null)
                    _restClientExchangeRatesNBP = new RestClient("https://api.nbp.pl");
                //https://api.nbp.pl/api/exchangerates/rates/a/usd/2022-10-28/
                //{"table":"A","currency":"dolar amerykański","code":"USD","rates":[{"no":"210/A/NBP/2022","effectiveDate":"2022-10-28","mid":4.7477}]}

                RestRequest request = new RestRequest($"/api/exchangerates/rates/{type}/{currency}/{dateString}?format=json", Method.GET);

                var resultRaw = _restClientExchangeRatesNBP.Execute<Object>(request).Data;
                if (resultRaw != null)
                {
                    var dict = resultRaw as Dictionary<string, object>;
                    object paramValue;
                    if (dict.TryGetValue("code", out paramValue))
                    {
                        string currencyReturned = paramValue as string;
                        if(string.Compare(currencyReturned, currency, true) == 0)
                        {
                            object fieldsRaw;
                            if (dict.TryGetValue("rates", out fieldsRaw))
                            {
                                var records = fieldsRaw as IList<object>;
                                var record0 = records[0];
                                var dictRates = record0 as Dictionary<string, object>;
                                object noRaw, midRaw, dateRaw;
                                if (dictRates.TryGetValue("no", out noRaw) 
                                    && dictRates.TryGetValue("effectiveDate", out dateRaw)
                                    && dictRates.TryGetValue("mid", out midRaw))
                                {
                                    double rate = ToDouble(midRaw);
                                    string effectiveDate = dateRaw as string;
                                    if(effectiveDate != dateString)
                                        notes = $"asked {dateString}, returned {effectiveDate}";
                                    else
                                    {
                                        result = rate;
                                        _ratesDictionnary.Add(key, result);
                                        ++_ratesAdded;
                                    }
                                }
                                else
                                    notes = $"unknown structure of rates array";
                            }
                            else
                                notes = "'rates' section is not found";
                        }
                        else
                            notes = "no requested currency returned";
                    }
                    else
                    {
                        notes = "coulnd't find 'code' section";
                    }
                }
                else
                    notes = "REST response is empty";

                log.Info($"DT.GetExchangeRateNBP: date ({dateString}), currency ({currency}), result ({result}), notes ({notes})");
                log.Flush();
            }

            return result;
        }

        static double ToDouble(object value)
        {
            double result = 0;
            try
            {
                result = (double)value;
            }
            catch(InvalidCastException)
            {
                result = (long)value;
            }
            return result;
        }
        string HttpRequestGet(string url, EasyLog log)
        {
            try
            {
                var response = _client.SendAsync(new HttpRequestMessage(HttpMethod.Get, url)).Result;
                return response.Content.ReadAsStringAsync().Result;
            }
            catch(Exception ex)
            {
                log.Error($"HttpRequestGet at {url} failed", ex);
                return null;
            }
        }
        public double GetExchangeUSDRateNBRB(DateTime date)
        {
            double result = 0.0;
            string key = GetHash(date, "BYN-NBRB", "USD");
            if (_ratesDictionnary.TryGetValue(key, out result))
                return result;

            using (EasyLog log = new EasyLog("ExcelToolkit"))
            {
                string dateString = date.ToString("yyyy-MM-dd");
                string notes = null;

                /*
                RestRequest request = new RestRequest($"/API/ExRates/Rates/145", Method.GET);
                request.AddHeader("Accept", "application/json");
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("onDate", dateString);

                var resultRaw = _restClientNBRB.Execute<Object>(request, Method.GET).Data;
                if (resultRaw != null)
                {
                    var dict = resultRaw as Dictionary<string, object>;
                    object fieldsRaw;
                    if (dict.TryGetValue("Cur_OfficialRate", out fieldsRaw))
                    {
                        result = (double)fieldsRaw;
                    }
                    else
                        notes = "Cur_OfficialRate field is not found in the REST response";
                }
                else
                    notes = "REST response is empty";
                    */

                string url = $"https://www.nbrb.by/API/ExRates/Rates/431?onDate={dateString}";
                string response = HttpRequestGet(url, log);
                if(response == null)
                    notes = "REST response is empty";
                else
                {
                    try
                    {
                        var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(response);
                        object fieldsRaw;
                        if (dict.TryGetValue("Cur_OfficialRate", out fieldsRaw))
                        {
                            result = (double)fieldsRaw;

                            _ratesDictionnary.Add(key, result);
                            ++_ratesAdded;
                        }
                        else
                            notes = "Cur_OfficialRate field is not found in the REST response";
                    }
                    catch(JsonReaderException)
                    {
                        notes = "Got Json deserialize exception";
                    }
                }

                log.Info($"DT.ExchangeUSDRateNBRB: date ({dateString}), result ({result}), notes ({notes})");
                log.Flush();
            }

            return result;
        }

        public void Dispose()
        {
            Debug.WriteLine("Rest disposed");
        }
    }
}
