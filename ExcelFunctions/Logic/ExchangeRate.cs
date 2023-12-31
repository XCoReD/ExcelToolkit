﻿using ExcelFunctions.Tools;
using Microsoft.Win32;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using Tools;

namespace ExcelFunctions
{
    public class ExchangeRate
    {
        RestClientRegistry _registry; 
        readonly string _fixerApiKey = "5db46eaf8fadb8c67faa0fbb9cad5595";

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

            _registry = new RestClientRegistry();
            _registry.Register(RestClientRegistry.Supplier.Fixer, "http://data.fixer.io"); //no https access in the base plan
            _registry.Register(RestClientRegistry.Supplier.BankNBP, "https://api.nbp.pl");
            _registry.Register(RestClientRegistry.Supplier.BankNBRB, "https://www.nbrb.by", false);

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

                var path = $"/api/{dateString}?access_key={_fixerApiKey}";
                var dict = _registry.Call(RestClientRegistry.Supplier.Fixer, path);
                if (dict != null)
                {
                    object paramValue;
                    if (dict.TryGetValue("base", out paramValue))
                    {
                        string serverBaseCurrency = paramValue as string;

                        object fieldsRaw;
                        if (dict.TryGetValue("rates", out fieldsRaw))
                        {
                            var j = fieldsRaw as JsonElement?;
                            if(j != null)
                            {
                                var dictRates = j.Value.ToObject<Dictionary<string, object>>();
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
                                notes = $"rates array is not found";
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
                    notes = "call failed";

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

                //https://api.nbp.pl/api/exchangerates/rates/a/usd/2022-10-28/
                //{"table":"A","currency":"dolar amerykański","code":"USD","rates":[{"no":"210/A/NBP/2022","effectiveDate":"2022-10-28","mid":4.7477}]}

                var path = $"/api/exchangerates/rates/{type}/{currency}/{dateString}?format=json";

                var dict = _registry.Call(RestClientRegistry.Supplier.BankNBP, path);
                if (dict != null)
                {
                    object paramValue;
                    if (dict.TryGetValue("code", out paramValue))
                    {
                        string currencyReturned = ToString(paramValue);
                        if (string.Compare(currencyReturned, currency, true) == 0)
                        {
                            object fieldsRaw;
                            if (dict.TryGetValue("rates", out fieldsRaw))
                            {
                                var records = (fieldsRaw as JsonElement?).Value.ToObject<IList<object>>();
                                var dictRates = (records[0] as JsonElement?).Value.ToObject<Dictionary<string, object>>();
                                object noRaw, midRaw, dateRaw;
                                if (dictRates.TryGetValue("no", out noRaw)
                                    && dictRates.TryGetValue("effectiveDate", out dateRaw)
                                    && dictRates.TryGetValue("mid", out midRaw))
                                {
                                    double rate = ToDouble(midRaw);
                                    string effectiveDate = ToString(dateRaw);
                                    if (effectiveDate != dateString)
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
                        notes = "response not recognized";
                }
                else
                    notes = "call failed";

                log.Info($"DT.GetExchangeRateNBP: date ({dateString}), currency ({currency}), result ({result}), notes ({notes})");
                log.Flush();
            }

            return result;
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

                string path = $"/API/ExRates/Rates/431?onDate={dateString}";
                var dict = _registry.Call(RestClientRegistry.Supplier.BankNBRB, path);
                if (dict != null)
                {
                    object fieldsRaw;
                    if (dict.TryGetValue("Cur_OfficialRate", out fieldsRaw))
                    {
                        result = ToDouble(fieldsRaw);

                        _ratesDictionnary.Add(key, result);
                        ++_ratesAdded;
                    }
                    else
                        notes = "Cur_OfficialRate field is not found in the REST response";
                }
                else
                    notes = "call failed";

                log.Info($"DT.ExchangeUSDRateNBRB: date ({dateString}), result ({result}), notes ({notes})");
                log.Flush();
            }

            return result;
        }

        static double ToDouble(object value)
        {
            double result = 0;
            if (value is JsonElement)
            {
                var v = (value as JsonElement?).Value;
                if (!v.TryGetDouble(out result))
                {
                    if (v.TryGetInt64(out long l))
                        result = l;
                }
            }
            else
            {
                try
                {
                    result = (double)value;
                }
                catch (InvalidCastException)
                {
                    result = (long)value;
                }
            }
            return result;
        }

        static string ToString(object value)
        {
            if (value is JsonElement)
            {
                var v = (value as JsonElement?).Value;
                return v.ToString();
            }
            else
            {
                return value as string;
            }
        }

        public void Dispose()
        {
            Debug.WriteLine("Rest disposed");
        }
    }
}
