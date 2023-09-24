using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Globalization;
using System.Diagnostics;
using System.IO;

namespace ExcelToolkit
{
    //https://gist.github.com/insideone/2c4172e6023a2b9ac07a0dbec61f4373
    static class IntegerExtend
    {
        /// <summary>
        /// Склоняет существительное в зависимости от числительного идущего перед ним.
        /// </summary>
        /// <param name="num">Число идущее перед существительным.</param>
        /// <param name="normative">Именительный падеж слова.</param>
        /// <param name="singular">Родительный падеж ед. число.</param>
        /// <param name="plural">Множественное число.</param>
        public static string Decline(this int num, string nominative, string singular, string plural)
        {
            if (num > 10 && ((num % 100) / 10) == 1) return plural;

            switch (num % 10)
            {
                case 1:
                    return nominative;
                case 2:
                case 3:
                case 4:
                    return singular;
                default: // case 0, 5-9
                    return plural;
            }
        }
    }
    //int i = 3;
    //Console.WriteLine(i.Decline("Прошёл", "Прошло", "Прошло") + " " + i.ToString() + " " + i.Decline("час", "часа", "часов"));

    public enum DisplayType
    {
        Text,
        Date, //31.10.2019
        //«31» октября 2019
        Number //23.40
    }
    public class Param
    {
        public string name;
        public string value;
        public DisplayType type;
    }
    public class DocumentProcessor
    {
        Application _word = null;
        Document _doc = null;
        string _company = null;
        string _month = null;
        string _num = null;
        Dictionary<string, Param> _paramValues;
        /*DateTime _startDate;
        DateTime _endDate;
        int _number;
        double _hours;
        double _hourRate;
        double _exchangeRate;*/

        CultureInfo _russian = new CultureInfo("ru-RU");

        /*const string _numSearchFor = "{#n#}";
        const string _dateStartFullSearchFor = "{#d-start-full#}"; //«01» октября 2019
        const string _dateEndFullSearchFor = "{#d-end-full#}"; //«31» октября 2019 
        const string _dateStartSearchFor = "{#d-start#}"; //01.10.2019
        const string _dateEndSearchFor = "{#d-end#}"; //31.10.2019 
        const string _hoursSearchFor = "{#h#}";//184
        const string _hoursGenitiveSearchFor = "{#h-h#}";//184 часов
        const string _sumUsdSearchFor = "{#$#}"; //4048
        const string _sumUsdPropSearchFor = "{#$-prop#}";//четыре тысячи сорок восемь долларов США
        const string _sumUsdPropDativeSearchFor = "{#$-prop-gen#}";//четырем тысячам сорока восьми долларам США
        const string _sumUsdGenitiveSearchFor = "{#$-units-gen#}";//долларов США
        const string _sumUsdDativeSearchFor = "{#$-units-dat#}";//долларам США
        const string _sumBynSearchFor = "{#byn#}";//8 316,21 
        const string _sumBynPropSearchFor = "{#byn-prop#}";//восемь тысяч триста шестнадцать рублей двадцать одну копейку
        const string _sumBynGenitiveSearchFor = "{#$-units#}";//белорусских рублей*/

        string[] _monthsGenitive = { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };


        public DocumentProcessor()
        {
            /*_startDate = new DateTime(2019, 11, 1);
            _endDate = new DateTime(2019, 11, 30);
            _number = 9;
            _exchangeRate = 2.1086;
            _hours = 160;
            _hourRate = 22;*/
        }

        public bool Process(string templateFileName, List<Param> arguments, out string outParams)
        {
            outParams = null;
            _word = new Application { Visible = true };

            if(ReadParams(arguments))
            {
                string basePath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
                foreach (var name in templateFileName.Split(';'))
                {
                    string templatePath = Path.Combine(basePath, name);
                    if (!File.Exists(templatePath))
                    {
                        Debug.WriteLine($"Template file {templatePath} does not exist, skipping");
                    }
                    else
                    {
                        _doc = _word.Documents.Open(FileName: templatePath, NoEncodingDialog: true, ConfirmConversions: false, Visible: true);
                        _doc.Activate();

                        if (DoProcess(arguments))
                        {
                            string outFileName = $"{name}-{_company}-{_num}-{_month}.docx";
                            string outPath = Path.Combine(basePath, outFileName);

                            //TODO: safe file name
                            _doc.SaveAs2(outPath);

                            if (!string.IsNullOrEmpty(outParams))
                                outParams += ';';
                            outParams += outFileName;
                        }
                        _doc.Close(SaveChanges: false);
                    }
                }
            }

            _word.Quit();
            _doc = null;
            _word = null;

            return true;
        }

        private bool ReadParams(List<Param> arguments)
        {
            _paramValues = new Dictionary<string, Param>(arguments.Count);
            foreach (var item in arguments)
            {
                if(!string.IsNullOrEmpty(item.name))
                    _paramValues.Add(item.name.ToLower(), item);
            }

            Param value = null;
            if (!_paramValues.TryGetValue("num", out value) || string.IsNullOrEmpty(value.value))
            {
                Debug.WriteLine($"Mandatory param \"Num\" is not set");
                return false;
            }
            _num = value.value;

            if (!_paramValues.TryGetValue("vendor", out value) || string.IsNullOrEmpty(value.value))
            {
                Debug.WriteLine($"Mandatory param \"Vendor\" is not set");
                return false;
            }
            _company = value.value;

            if (!_paramValues.TryGetValue("month", out value) || string.IsNullOrEmpty(value.value))
            {
                Debug.WriteLine($"Mandatory param \"Month\" is not set");
                return false;
            }
            _month = value.value;

            return true;
        }

        private string FormatNumericParam(Param item, string option, string verb)
        {
            double d;
            string text = null;
            //option is not empty - suppose the numeric value
            if (!double.TryParse(item.value, out d))
            {
                Debug.Assert(false);
                Debug.WriteLine($"Option \"{option}\" is given for non-numeric or non-date value. Parameter: {item.name}, value: {item.value}");
                text = item.value;
            }
            else
            {
                try
                {
                    const string inWords = "inwords";
                    const string units = "units";
                    const string hours = "hours";
                    bool useUSD = false;
                    if (string.IsNullOrEmpty(option) || option.ToLower() == "brief")
                    {
                        text = String.Format("{0:0.##}", d);
                    }
                    else 
                    if (option.ToLower() == "separated")
                    {
                        //https://www.csharp-examples.net/string-format-double/
                        text = String.Format("{0:0,0.00}", d).Replace(',', '\u00A0').Replace('.', ',');
                    }
                    else
                    if (option.IndexOf(inWords, 0, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        string currency = option.Length > inWords.Length ? option.Substring(inWords.Length) : "";
                        switch (currency.ToLower())
                        {
                            case "usd": useUSD = true; break;
                            case "byn": break;
                            case "": break;
                            default:
                                //suppose byn but warn!
                                Debug.Assert(false);
                                Debug.WriteLine($"Unknown currency code: \"{currency}\". Parameter: {item.name}, value: {item.value}");
                                break;
                        }
                        TextCase tc = string.Compare(verb, "DP", true) == 0 ? TextCase.Dative : 
                            (string.Compare(verb, "RP", true) == 0 ? TextCase.Genitive : TextCase.Nominative);
                        text = RuDateAndMoneyConverter.CurrencyToTxt(d, tc, false, useUSD);
                        //четыре тысячи сорок восемь долларов США
                    }
                    else
                    if (option.IndexOf(units, 0, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        //TODO: refactor
                        string currency = option.Length > units.Length ? option.Substring(units.Length) : "";
                        switch (currency.ToLower())
                        {
                            case "usd": useUSD = true; break;
                            case "byn": break;
                            case "": break;
                            default:
                                //suppose byn but warn!
                                Debug.Assert(false);
                                Debug.WriteLine($"Unknown currency code: \"{currency}\". Parameter: {item.name}, value: {item.value}");
                                break;
                        }

                        int whole = (int)d;
                        TextCase tc = string.Compare(verb, "DP", true) == 0 ? TextCase.Dative :
                            (string.Compare(verb, "RP", true) == 0 ? TextCase.Genitive : TextCase.Nominative);
                        if(useUSD)
                        {
                            switch(tc)
                            {
                                case TextCase.Nominative:
                                    text = whole.Decline("доллар США", "доллара США", "долларов США"); break;
                                case TextCase.Genitive:
                                    text = whole.Decline("доллара США", "доллара США", "долларов США"); break;
                                case TextCase.Dative:
                                    text = whole.Decline("доллару США", "долларам США", "долларам США"); break;
                                default:
                                    Debug.Assert(false); break;
                            }
                        }
                        else
                        {
                            switch (tc)
                            {
                                case TextCase.Nominative:
                                    text = whole.Decline("белорусский рубль", "белорусских рубля", "белорусских рублей"); break;
                                case TextCase.Genitive:
                                    text = whole.Decline("белорусский рубль", "белорусских рубля", "белорусских рублей"); break;
                                case TextCase.Dative:
                                    text = whole.Decline("белорусскому рублю", "белорусским рублям", "белорусским рублям"); break;
                                default:
                                    Debug.Assert(false); break;
                            }
                        }
                    }
                    else
                    if (String.Compare(option, hours, true) == 0)
                    {
                        int whole = (int)d;
                        if (Math.Abs((double)whole - d) < 0.01)
                            text = $"{whole} {whole.Decline("час", "часа", "часов")}";
                        else
                        {
                            int minutes = (int)(d - whole);
                            text = $"{whole} {whole.Decline("час", "часа", "часов")} {minutes} {minutes.Decline("минута", "минуты", "минут")}";
                        }
                    }
                    else
                    {
                        Debug.Assert(false);
                        Debug.WriteLine($"Unknown option: \"{option}\". Parameter: {item.name}, value: {item.value}");
                    }
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    Debug.Assert(false);
                    Debug.WriteLine($"Option \"{option}\" is given with no proper currency / verb. Parameter: {item.name}, value: {item.value}");
                    text = item.value;
                }
            }
            return text;
        }


        private bool ProcessParam(Param item, string option = null, string verb = null)
        {
            string text = "n/a";
            if (!string.IsNullOrEmpty(item.value))
            {
                if (item.type == DisplayType.Text)
                {
                    if (string.IsNullOrEmpty(option))
                        text = item.value;
                    else
                    {
                        Debug.WriteLine($"Param {item.name}, value {item.value}: specified display option \"{option}\", suppose number formatting(!?)");
                        text = FormatNumericParam(item, option, verb);
                    }
                }
                else if (item.type == DisplayType.Date)
                {
                    //date
                    double d = double.Parse(item.value);
                    DateTime conv = DateTime.FromOADate(d);
                    if (string.IsNullOrEmpty(option) || option.ToLower() == "brief")
                    {
                        text = conv.ToString("dd.MM.yyyy");
                    }
                    else if (option.ToLower() == "long")
                    {
                        text = $"{conv.Day} {_monthsGenitive[conv.Month - 1]} {conv.Year}";
                    }
                    else if (option.ToLower() == "longquoted")
                    {
                        text = $"«{conv.Day}» {_monthsGenitive[conv.Month - 1]} {conv.Year}";
                    }
                    else if (option.ToLower() == "longquoted2")
                    {
                        text = $"\"{conv.Day}\" {_monthsGenitive[conv.Month - 1]} {conv.Year}";
                    }
                    else
                    {
                        Debug.Assert(false);
                        Debug.WriteLine($"Param {item.name}, value {item.value}: specified display option \"{option}\" is not valid");
                    }
                }
                else if (item.type == DisplayType.Number)
                {
                    text = FormatNumericParam(item, option, verb);
                }
                else
                {
                    Debug.Assert(false);
                }
            }

            string fullTextToSearch = "{#" + item.name 
                + (string.IsNullOrEmpty(option) ? "" : ":" + option) 
                + (string.IsNullOrEmpty(verb) ? "" : ":" + verb)
                + "#}";

            return FindAndReplace(fullTextToSearch, text);
        }

        private bool DoProcess(List<Param> arguments)
        {
            //first process all entries in arguments
            foreach(var item in arguments)
            {
                if (!string.IsNullOrEmpty(item.name))
                {
                    if(!ProcessParam(item))
                    {
                        Debug.WriteLine($"Param \"{item.name}\" is not found in the template");
                    }
                }
            }

            //now process all derived entries in the document
            string allWords = _doc.Content.Text;
            int iStart = 0;
            for(;;)
            {
                int i = allWords.IndexOf("{#", iStart);
                if (i < 0)
                    break;

                int iEnd = allWords.IndexOf("#}", i);
                if (iEnd > 0)
                {
                    string lexem = allWords.Substring(i + 2, iEnd - i - 2);
                    int iOptions = lexem.IndexOf(':');
                    if (iOptions == -1)
                    {
                        //unknown or misprinted param
                        Debug.Assert(false);
                        Debug.WriteLine($"Unknown or misprinted param: {lexem}");
                    }
                    else
                    {
                        string option = null;
                        string verb = null;
                        //find base param
                        int iVerb = lexem.IndexOf(':', iOptions + 1);
                        if(iVerb > 0)
                        {
                            option = lexem.Substring(iOptions + 1, iVerb - iOptions - 1);
                            verb = lexem.Substring(iVerb + 1);
                        }
                        else
                            option = lexem.Substring(iOptions + 1);

                        string paramName = lexem.Substring(0, iOptions);

                        Param value = null;
                        if (!_paramValues.TryGetValue(paramName.ToLower(), out value))
                        {
                            Debug.Assert(false);
                            Debug.WriteLine($"Param value is not set: {paramName}, derived param: {lexem}");
                        }
                        else
                        {
                            ProcessParam(value, option, verb);
                        }
                    }
                    iStart = iEnd + 2;
                }
                else
                    iStart = i + 2;
            }

            return true;
        }

        //https://stackoverflow.com/questions/19252252/c-sharp-word-interop-find-and-replace-everything
        private bool FindAndReplace(object findText, object replaceWithText)
        {
            bool result;
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            result = _word.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);

            Debug.WriteLine($"Replaced {result}: {findText} => {replaceWithText}");
            return result;
        }
    }
}
