using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace ExcelFunctions
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
        public int useCount;
    }
    public class DocumentProcessor
    {
        Microsoft.Office.Interop.Word.Application _word = null;
        Document _doc = null;

        string _company = null;
        string _month = null;
        string _num = null;

        Dictionary<string, Param> _paramValues;
        ILog _log;
        bool _autoCloseWord;

        CultureInfo _russian = new CultureInfo("ru-RU");

        string[] _monthsGenitive = { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };

        public DocumentProcessor(ILog log, bool autoCloseWord)
        {
            _log = log;
            _autoCloseWord = autoCloseWord;
        }


        public bool Process(string basePath, string templateFileName, List<Param> arguments, out string outParams)
        {
            outParams = null;
            _word = new Microsoft.Office.Interop.Word.Application { Visible = true };
            bool result = false;

            if(ReadParams(arguments))
            {
                //string basePath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
                foreach (var name in templateFileName.Split(';'))
                {
                    string templatePath = Path.Combine(basePath, name);
                    if (!File.Exists(templatePath))
                    {
                        _log.Error($"Template file {templatePath} does not exist, processing skipped");
                    }
                    else
                    {
                        var e = Path.GetExtension(templatePath).ToLowerInvariant();
                        if(e != ".docx")
                        {
                            _log.Error($"Extension of template file must be docx");
                            MessageBox.Show($"Extension of template file must be docx, given {e}");
                        }
                        else
                        {
                            _doc = _word.Documents.Open(FileName: templatePath, NoEncodingDialog: true, ConfirmConversions: false, Visible: true);
                            _doc.Activate();

                            if (DoProcess(arguments))
                            {
                                string outFileName = null;
                                string outPath = GetSafeOutputFileName(basePath, name, out outFileName);
                                if (string.IsNullOrEmpty(outPath))
                                {
                                    _log.Error($"Failed to generate an output file name");
                                }
                                else
                                {
                                    _doc.SaveAs2(outPath);

                                    if (!string.IsNullOrEmpty(outParams))
                                        outParams += ';';
                                    outParams += outFileName;
                                    result = true;
                                }
                            }

                            if (_autoCloseWord)
                                _doc.Close(SaveChanges: false);
                        }
                    }
                }
            }

            if(_autoCloseWord)
                _word.Quit();

            _doc = null;
            _word = null;

            return result;
        }

        private string GetSafeOutputFileName(string basePath, string templateFileName, out string outFileName)
        {
            outFileName = null;
            string tail;
            if (!string.IsNullOrEmpty(_company) && !string.IsNullOrEmpty(_num) && !string.IsNullOrEmpty(_month))
                tail = $"{_company}-{_num}-{_month}";
            else
                tail = DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss");

            basePath = basePath + "\\Generated";
            Directory.CreateDirectory(basePath);

            for (int i = 0; i< 20; ++i)
            {
                outFileName = $"{Path.GetFileNameWithoutExtension(templateFileName)}-{tail}{(i==0?"":"("+i.ToString()+")")}.docx";
                string outPath = Path.Combine(basePath, outFileName);
                if (!File.Exists(outPath))
                    return outPath;
            }

            return null;
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
                //_log.Info($"Mandatory param \"Num\" is not set");
            }
            else
                _num = value.value;

            value = null;
            if (!_paramValues.TryGetValue("vendor", out value) || string.IsNullOrEmpty(value.value))
            {
                if (!_paramValues.TryGetValue("company", out value) || string.IsNullOrEmpty(value.value))
                {
                    if (!_paramValues.TryGetValue("customer", out value) || string.IsNullOrEmpty(value.value))
                    {
                        //_log.Info($"Mandatory param \"Vendor\"/\"Company\"/\"Customer\" is not set");
                    }
                }
            }
            if(value != null && !string.IsNullOrEmpty(value.value))
                _company = value.value;

            if (!_paramValues.TryGetValue("month", out value) || string.IsNullOrEmpty(value.value))
            {
                //_log.Info($"Mandatory param \"Month\" is not set");
            }
            else
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
                _log.Warning($"Option \"{option}\" is given for non-numeric or non-date value. Parameter: {item.name}, value: {item.value}");
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
                        text = String.Format(CultureInfo.InvariantCulture, "{0:0.##}", d);
                    }
                    else 
                    if (option.ToLower() == "separated")
                    {
                        //https://www.csharp-examples.net/string-format-double/
                        text = String.Format(CultureInfo.InvariantCulture, "{0:0,0.00}", d).Replace(',', '\u00A0').Replace('.', ',');
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
                                _log.Error($"Unknown currency code: \"{currency}\". Parameter: {item.name}, value: {item.value}");
                                break;
                        }
                        Declination tc = string.Compare(verb, "DP", true) == 0 ? Declination.Dative : 
                            (string.Compare(verb, "RP", true) == 0 ? Declination.Genitive : Declination.Nominative);
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
                                _log.Error($"Unknown currency code: \"{currency}\". Parameter: {item.name}, value: {item.value}");
                                break;
                        }

                        int whole = (int)d;
                        Declination tc = string.Compare(verb, "DP", true) == 0 ? Declination.Dative :
                            (string.Compare(verb, "RP", true) == 0 ? Declination.Genitive : Declination.Nominative);
                        if(useUSD)
                        {
                            switch(tc)
                            {
                                case Declination.Nominative:
                                    text = whole.Decline("доллар США", "доллара США", "долларов США"); break;
                                case Declination.Genitive:
                                    text = whole.Decline("доллара США", "доллара США", "долларов США"); break;
                                case Declination.Dative:
                                    text = whole.Decline("доллару США", "долларам США", "долларам США"); break;
                                default:
                                    Debug.Assert(false); break;
                            }
                        }
                        else
                        {
                            switch (tc)
                            {
                                case Declination.Nominative:
                                    text = whole.Decline("белорусский рубль", "белорусских рубля", "белорусских рублей"); break;
                                case Declination.Genitive:
                                    text = whole.Decline("белорусский рубль", "белорусских рубля", "белорусских рублей"); break;
                                case Declination.Dative:
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
                        _log.Error($"Unknown option: \"{option}\". Parameter: {item.name}, value: {item.value}");
                    }
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    Debug.Assert(false);
                    _log.Error($"Option \"{option}\" is given with no proper currency / verb. Parameter: {item.name}, value: {item.value}");
                    text = item.value;
                }
            }
            return text;
        }


        private bool ProcessParam(Param item, string option, string verb, string originalLexem)
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
                        //_log.Info($"Param {item.name}, value {item.value}: specified display option \"{option}\", suppose number formatting(!?)");
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
                        _log.Error($"Param {item.name}, value {item.value}: specified display option \"{option}\" is not valid");
                    }
                }
                else if (item.type == DisplayType.Number)
                {
                    text = FormatNumericParam(item, option, verb);
                }
                else
                {
                    Debug.Assert(false);
                    _log.Error($"Unknown display type for parameter {item.type}");
                }
            }

            if (string.IsNullOrEmpty(originalLexem))
                originalLexem = item.name
                + (string.IsNullOrEmpty(option) ? "" : ":" + option)
                + (string.IsNullOrEmpty(verb) ? "" : ":" + verb);

            string fullTextToSearch = "{#" + originalLexem + "#}";

            bool result = FindAndReplace(fullTextToSearch, text);
            if (result)
                ++item.useCount;

            return result;
        }

        private bool DoProcess(List<Param> arguments)
        {
            HashSet<string> variables = new HashSet<string>();
            //process all entries in the document, including derived, or entries with embedded spaces
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
                    string normalizedLexem = lexem.Replace(" ", "");
                    string option = null;
                    string verb = null;
                    string paramName = null;
                    int iOptions = normalizedLexem.IndexOf(':');
                    if (iOptions != -1)
                    {
                        //find base param
                        int iVerb = normalizedLexem.IndexOf(':', iOptions + 1);
                        if (iVerb > 0)
                        {
                            option = normalizedLexem.Substring(iOptions + 1, iVerb - iOptions - 1);
                            verb = normalizedLexem.Substring(iVerb + 1);
                        }
                        else
                            option = normalizedLexem.Substring(iOptions + 1);

                        paramName = normalizedLexem.Substring(0, iOptions);
                    }
                    else
                        paramName = normalizedLexem;

                    if(!variables.Contains(lexem))
                    {
                        Param value = null;
                        if (!_paramValues.TryGetValue(paramName.ToLower(), out value))
                        {
                            _log.Error($"Param value is not set: param: ({paramName}), full name: ({lexem})");
                        }
                        else
                        {
                            ProcessParam(value, option, verb, lexem);
                        }
                        variables.Add(lexem);
                    }

                    iStart = iEnd + 2;
                }
                else
                    iStart = i + 2;
            }

            //now check if all entries are used
            foreach(var item in arguments)
            {
                if(item.useCount == 0)
                {
                    if(!string.IsNullOrEmpty(item.name) && 
                        string.Compare(item.name, "doctemplate", true) != 0 &&
                        string.Compare(item.name, "docgen", true) != 0 &&
                        string.Compare(item.name, "month", true) != 0 &&
                        string.Compare(item.name, "docgendate", true) != 0)
                    {
                        _log.Warning($"Param \"{item.name}\" is not used in template. Have you missed something??");
                    }
                }
            }

            return true;
        }

        //https://stackoverflow.com/questions/19252252/c-sharp-word-interop-find-and-replace-everything
        private bool FindAndReplace(string findText, string replaceWithText)
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
            //object read_only = false;
            //object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace

            bool found = false;
            int numReplaces = 0;
            for(;;)
            {
                Range range = _word.ActiveDocument.Content;
                found = range.Find.Execute(findText);
                if (!found)
                    break;
                    //, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                range.Text = replaceWithText;
                ++numReplaces;
            }

            if (numReplaces != 0)
            {
                _log.Info($"FindAndReplace: ({numReplaces}) {findText} => {replaceWithText}");
            }

            return numReplaces != 0;
        }
    }
}
