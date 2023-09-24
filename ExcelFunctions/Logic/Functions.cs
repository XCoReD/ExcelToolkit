using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using Tools;

namespace ExcelFunctions
{

    public static class ExcelFunctions
    {
        static ExchangeRate _rest;
        /*[ExcelFunction(Name = "DT.CheckAlive", Description = "Check if plugin is alive", Category = "DT")]
        public static string Function_CheckAlive()
        {
            return "OK";
        }*/

        static Declination Convert(string verb)
        {
            Declination tc = string.Compare(verb, "DP", true) == 0 ? Declination.Dative :
                (string.Compare(verb, "RP", true) == 0 ? Declination.Genitive :
                (string.Compare(verb, "VP", true) == 0 ? Declination.Accusative :
                (string.Compare(verb, "TP", true) == 0 ? Declination.Instrumental :
                (string.Compare(verb, "PP", true) == 0 ? Declination.Prepositional :
                Declination.Nominative))));
            return tc;
        }

        static bool Convert(string date, out DateTime dateD)
        {
            dateD = default(DateTime);
            if (!string.IsNullOrEmpty(date))
            {
                double d;
                if (Double.TryParse(date, out d))
                {
                    dateD = DateTime.FromOADate(d);
                    return true;
                }

                var result = DateParser.ParseDateTime(date);
                if (result.HasValue)
                {
                    dateD = result.Value;
                    return true;
                }
            }

            return false;
        }

        [ExcelFunction(Name = "DT.ExchangeRateNBP", Description = "Get the medium exchange rate of a given currency to zloty by Narodowy Bank Polski on a given date", Category = "DT")]
        public static double Function_GetExchangeRateNBP([ExcelArgument("Date, e.g. 2019/11/29")] string date, [ExcelArgument("Currency, e.g. USD")] string currency)
        {
            DateTime dt;
            if (!Convert(date, out dt))
                return 0;
            if (string.IsNullOrEmpty(currency))
                return 0;

            if (_rest == null)
                _rest = new ExchangeRate();
            return _rest.GetExchangeRateNBP(dt, currency);
            //return 0;
        }

        [ExcelFunction(Name = "DT.ExchangeRate", Description = "Get a given currency exchange rate defined by European Central Bank in terms of a base currency rate on a given date", Category = "DT")]
        public static double Function_GetExchangeRate([ExcelArgument("Date, e.g. 2019/11/29")] string date, [ExcelArgument("Currency, e.g. USD")] string currency, [ExcelArgument("Base currency, e.g. EUR")] string baseCurrency)
        {
            DateTime dt;
            if (!Convert(date, out dt))
                return 0;
            if (string.IsNullOrEmpty(currency) || string.IsNullOrEmpty(baseCurrency))
                return 0;

            if (_rest == null)
                _rest = new ExchangeRate();
            return _rest.GetExchangeRate(dt, currency, baseCurrency);
            //return 0;
        }

        [ExcelFunction(Name = "DT.ExchangeUSDRateNBRB", Description = "Get USD rate on a given date (e.g. 2019/11/29) from NBRB as a double (e.g. 2.1135)", Category = "DT")]
        public static double Function_ExchangeUSDRateNBRB([ExcelArgument("Date, e.g. 2019/11/29")] string date)
        {
            DateTime dt;
            if (!Convert(date, out dt))
                return 0;

            if (_rest == null)
                _rest = new ExchangeRate();
            return _rest.GetExchangeUSDRateNBRB(dt);
            //return 0;
        }

        [ExcelFunction(Name = "DT.SumProp", Description = "Сумма прописью во всех вариантах", Category = "DT")]
        public static string Function_SumProp([ExcelArgument("Sum, e.g. 123.34")] double sum, [ExcelArgument("Currency, e.g. USD")] string currency, [ExcelArgument("Text case (падеж). Options: NP(=null),RP,DR")] string verb, [ExcelArgument("Capitalize first letter. TRUE or FALSE")] bool capitalizeFirst)
        {
            bool useUSD = string.Compare(currency, "USD", true) == 0 ? true : false;
            Declination tc = Convert(verb);
            return RuDateAndMoneyConverter.CurrencyToTxt(sum, tc, capitalizeFirst, useUSD);
        }

        [ExcelFunction(Name = "DT.Range", Description = "Get a range of a given number, i.e. '1 - 10', '10-100', '101-1000'", Category = "DT")]
        public static string Function_Range([ExcelArgument("Number, e.g. 123.34")] double number)
        {
            if (number < 1.0)
                return "";

            if (number < 11.0)
                return "1-10";
            if (number < 51.0)
                return "11-50";
            if (number < 201.0)
                return "51-200";
            if (number < 501.0)
                return "201-500";
            if (number < 1001.0)
                return "501-1000";
            if (number < 5001.0)
                return "1001-5000";
            if (number < 10001.0)
                return "5001-10,000";

            return "10,001+";
        }

        [ExcelFunction(Name = "DT.Decline", Description = "Склонение", Category = "DT")]
        public static string Function_Decline([ExcelArgument("Noun, e.g. Инженер-программист")] string arg, [ExcelArgument("Text case (падеж). Options: NP(=null),RP,VP,DR,TP,PP")] string verb, [ExcelArgument("Capitalize first letter. TRUE or FALSE")] bool capitalizeFirst)
        {
            Declination tc = Convert(verb);
            return RussianDeclination.Decline(arg, tc, capitalizeFirst);
        }

        [ExcelFunction(Name = "DT.DeclineFIO", Description = "Склонение Ф.И.О.", Category = "DT")]
        public static string Function_DeclineFIO([ExcelArgument("Фамилия Имя Отчество, e.g. Сидоров Петр Иванович")] string arg, [ExcelArgument("Text case (падеж). Options: NP(=null),RP,DR,TP,PP")] string verb, [ExcelArgument("Truncate I.O., e.g. Сидорова П.И. TRUE or FALSE")] bool truncateFirstNameParentName, [ExcelArgument("Reverse I.O., e.g. П.И.Сидоров. TRUE or FALSE")] bool reverseFirstNameParentName)
        {
            Declination tc = Convert(verb);
            return RussianDeclination.DeclineFIO(arg, tc, truncateFirstNameParentName, reverseFirstNameParentName);
        }

        [ExcelFunction(Name = "DT.Dummy", Description = "Test entry with no parameters, giving version string", Category = "DT")]
        public static string Function_Dummy()
        {
            string assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            return "Works OK. Version: " + assemblyVersion;
        }

        [ExcelFunction(Name = "DT.DummyDate", Description = "Test entry with date parameter. Returns date as string YYYY-MM-DD format", Category = "DT")]
        public static string Function_DummyDate([ExcelArgument("Date, e.g. 2019/11/29")] string date)
        {
            DateTime dt;
            if (!Convert(date, out dt))
                return "N/A";

            return $"{dt.Year}-{dt.Month}-{dt.Day}";
        }
    }
}
