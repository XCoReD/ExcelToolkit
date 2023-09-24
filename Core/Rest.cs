using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Tools;

namespace Core
{
    public class Rest
    {
        RestClient _restClientNBRB = new RestClient("http://www.nbrb.by/");
        RestClient _restClientExchangeRates = new RestClient("https://api.exchangeratesapi.io/");
        EasyLog _log = new EasyLog();

        public double GetExchangeRate(DateTime date, string currency, string baseCurrency)
        {
            string dateString = date.ToString("yyyy-MM-dd");
            double result = 0.0;

            RestRequest request = new RestRequest($"/{dateString}?base={baseCurrency}&symbols={currency}", Method.GET);
            var resultRaw = _restClientExchangeRates.Execute<Object>(request).Data;
            if (resultRaw != null)
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
#if DEBUG
                        if (dict.TryGetValue("base", out fieldsRaw))
                        {
                            Debug.Assert(fieldsRaw as string == baseCurrency);
                            if (dict.TryGetValue("date", out fieldsRaw))
                            {
                                Debug.Assert(fieldsRaw as string == dateString);
                            }
                        }
#endif
                        result = (double)rateRaw;
                    }
                }

            }

            return result;
        }

        public double GetExchangeUSDRateNBRB(DateTime date)
        {
            _log.Info($"ScienceSoft.ExchangeUSDRateNBRB: date({date.ToShortDateString()})");
            string dateString = date.ToString("yyyy-MM-dd");
            double result = 0.0;

            RestRequest request = new RestRequest($"/API/ExRates/Rates/145?onDate={dateString}", Method.GET);
            var resultRaw = _restClientNBRB.Execute<Object>(request).Data;
            if (resultRaw != null)
            {
                var dict = resultRaw as Dictionary<string, object>;
                object fieldsRaw;
                if (dict.TryGetValue("Cur_OfficialRate", out fieldsRaw))
                {
                    result = (double)fieldsRaw;
                }
            }

            _log.Info($"ScienceSoft.ExchangeUSDRateNBRB: result: ({result})");
            _log.Flush();

            return result;

        }
    }
}
