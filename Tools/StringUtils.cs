using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Tools
{
    public static class StringUtils
    {
        public static readonly string[] PresentTimeConstants = { "present", "current time", "current", "now", "the present" };
        public static readonly char[] TrimEndCommas = { ',', '.', ' ' };
        public static char[] _deniedStringLiteralCharacters = { '/', ':', '_' };
        public static char[] DateRangeSeparators = { '-', '—', '–', ':', '_', ',','\t' };//longer dash inserted by Word somehow. Treat Tab also as datetime separator!
        public static char[] DateRangeSeparatorsNoComma = { '-', '—', '–', ':', '\t' };//longer dash inserted by Word somehow. Treat Tab also as datetime separator!
        public static char[] WhiteSpaceSeparators = { ' ', '\t' }; //ignore any others


        public static bool CompareRoughly(string a, string b, int sensitivity = 2)
        {
            int minSensivity = Math.Min(sensitivity, a.Length - 2);//otherwise XSD is equal to RAD!
            return CalcLevenshteinDistance(a.ToLower(), b.ToLower()) <= minSensivity;
        }

        public static bool CompareRoughly(string a, string[] b, int sensitivity = 2)
        {
            return GetRoughMatch(a, b, sensitivity) != null;
        }

        public static string GetRoughMatch(string a, string[] b, int sensitivity = 2)
        {
            int minSensivity = Math.Min(sensitivity, a.Length - 2);//otherwise XSD is equal to RAD!
            foreach (var v in b)
            {
                if (CalcLevenshteinDistance(a.ToLower(), v.ToLower()) <= minSensivity)
                    return v;
            }
            return null;
        }
        public static int IndexOf(string a, string[] b)
        {
            int i = 0;
            foreach (var v in b)
            {
                int result = v.IndexOf(a, StringComparison.InvariantCultureIgnoreCase);
                if (result != -1)
                    return i;
                ++i;
            }
            return -1;
        }

        public static int CalcLevenshteinDistance(string a, string b)
        {
            //https://stackoverflow.com/questions/9453731/how-to-calculate-distance-similarity-measure-of-given-2-strings
            if (String.IsNullOrEmpty(a) || String.IsNullOrEmpty(b))
                return 0;

            int lengthA = a.Length;
            int lengthB = b.Length;
            var distances = new int[lengthA + 1, lengthB + 1];
            for (int i = 0; i <= lengthA; distances[i, 0] = i++) ;
            for (int j = 0; j <= lengthB; distances[0, j] = j++) ;

            for (int i = 1; i <= lengthA; i++)
                for (int j = 1; j <= lengthB; j++)
                {
                    int cost = b[j - 1] == a[i - 1] ? 0 : 1;
                    distances[i, j] = Math.Min
                        (
                        Math.Min(distances[i - 1, j] + 1, distances[i, j - 1] + 1),
                        distances[i - 1, j - 1] + cost
                        );
                }
            return distances[lengthA, lengthB];
        }

        public static string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1).ToLower();

            return str.ToUpper();
        }

        public static string RemoveSpacesAndDashes(string text)
        {
            return text.Replace(".", "").Replace(" ", "").Replace("-", "");
        }

        public static string TakeWords(string str, int wordCount)
        {
            //https://stackoverflow.com/questions/13368345/get-first-250-words-of-a-string
            char lastChar = '\0';
            int spaceFound = 0;
            var strLen = str.Length;
            int i = 0;
            for (; i < strLen; i++)
            {
                if (str[i] == ' ' && lastChar != ' ')
                {
                    spaceFound++;
                }
                lastChar = str[i];
                if (spaceFound == wordCount)
                    break;
            }
            return str.Substring(0, i);
        }

        public static string SmartTruncate(string text, int maxLength, ILog log = null)
        {
            text = NormalizeSpaces(text);
            if (text.Length < maxLength)
                return text;

            //smart truncate, with abbreviating the common used words
            text = Regex.Replace(text, "master data", "MD", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "master data management", "MDM", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "project manager", "PM", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "team leader", "TL", RegexOptions.IgnoreCase);
            if (text.Length < maxLength)
                return text;

            text = Regex.Replace(text, "senior", "", RegexOptions.IgnoreCase);
            if (text.Length < maxLength)
                return text;

            text = text.Substring(0, maxLength - 1);
            if (log != null)
                log.Warning($"Smart truncate to {maxLength} is hard, please fix [{text}]");
            return text;
        }

        /*public static int CountWords(string s)
        {
            //another option
            //https://stackoverflow.com/questions/8784517/counting-number-of-words-in-c-sharp
            int c = 0;
            for (int i = 1; i < s.Length; i++)
            {
                if (char.IsWhiteSpace(s[i - 1]) == true)
                {
                    if (char.IsLetterOrDigit(s[i]) == true ||
                        char.IsPunctuation(s[i]))
                    {
                        c++;
                    }
                }
            }
            if (s.Length > 2)
            {
                c++;
            }
            return c;
        }*/

        public static bool IsYear(string text)
        {
            double test;
            if (!Double.TryParse(text.Replace(',', '.').Replace("+", String.Empty), out test))
                return false;

            if (test < 1990 || test > 2030)
                return false;

            return true;
        }

        public static bool IsYearsOfExperience(string text)
        {
            double test;
            if (!Double.TryParse(text, out test))
                return false;

            if (test < 0.1 || test > 30)
                return false;

            return true;
        }

        public static bool IsYearOrNow(string text)
        {
            if (IsYear(text))
                return true;

            if (StringUtils.CompareRoughly(text, "present") || StringUtils.CompareRoughly(text, "now"))
                return true;

            return false;
        }

        public static string NormalizeDate(string text)
        {
            string[] arr = text.Split(WhiteSpaceSeparators);
            if (arr.Count() != 2)
                return text;    //no date, i.e. "MS Visual Studio 2016"

            var dateFormatInfo = CultureInfo.GetCultureInfo("en-GB").DateTimeFormat;
            string[] fullNames = dateFormatInfo.MonthNames;
            string[] shortNames = dateFormatInfo.AbbreviatedMonthNames;

            for (int i = 0; i < 12; ++i)
            {
                if(CompareRoughly(arr[0], fullNames[i]) ||
                    CompareRoughly(arr[0], shortNames[i]))
                {
                    arr[0] = fullNames[i];
                    break;
                }
            }

            return String.Join(" ", arr);
        }
        public static string GetFirstDate(string text)
        {
            string[] items = text.Split(' ');
            string dateCandidate = "";
            foreach (var item in items)
            {
                if (!string.IsNullOrEmpty(dateCandidate))
                    dateCandidate += ' ';
                dateCandidate += item;

                if (DateParser.ParseDateTime(dateCandidate).HasValue)
                    return dateCandidate;
            }
            return null;
        }
        public static string GetCurrentYear()
        {
            return DateTime.Now.Year.ToString();
        }

        public static bool ContainsAlpha(this string str)
        {
            //prototype: https://stackoverflow.com/questions/1046740/how-can-i-validate-a-string-to-only-allow-alphanumeric-characters-in-it
            if (string.IsNullOrEmpty(str))
                return false;

            for (int i = 0; i < str.Length; i++)
            {
                if (char.IsLetter(str[i]))
                    return true;
            }

            return false;
        }
        public static string NormalizeSpaces(string text, bool remove = false)
        {
            //todo: do not use a canon for sparrows
            return Regex.Replace(text, @"\s+", (remove? "": " "));
        }

        public static string SkipLeadingNumber(string text)
        {
            //skip 1. xxx
            //skip 1.1. xxx
            //skip 1.1 xxx
            if(Char.IsDigit(text[0]))
            {
                int pos = text.IndexOfAny(WhiteSpaceSeparators);
                if (pos != -1 && pos < text.Length - 1)
                    text = text.Substring(pos + 1).Trim();
            }
            return text;
        }
        public static string GetTrailingYear(string text)
        {
            if (string.IsNullOrEmpty(text))
                return null;

            string cleared = text.Replace(".", " ").Trim();
            if (cleared.Length < 4)
                return null;

            string lastFour = cleared.Substring(cleared.Length - 4);
            DateTime dt;
            if (!DateTime.TryParseExact(lastFour, "yyyy", new CultureInfo("en-us"), DateTimeStyles.None, out dt))
                return null;

            return lastFour;
        }

        public static string Append(string str1, string str2, string trailingChars = ".", bool appendSpace = true, bool appendNewLine = false)
        {
            //one more bycicle
            if (!string.IsNullOrEmpty(str1))
            {
                bool appendTrailingChars = string.IsNullOrEmpty(trailingChars) || !str1.EndsWith(trailingChars);
                if (appendTrailingChars)
                    str1 += trailingChars;
                if (appendSpace)
                    str1 += ' ';
                if (appendNewLine)
                    str1 += '\n';

                str1 += str2;
            }
            else
                str1 = str2;

            return str1;
        }
        public static string AppendWithComma(string str1, string str2)
        {
            return Append(str1, str2, ",");
        }

        public static int CountWords(string test)
        {
            // https://stackoverflow.com/questions/8784517/counting-number-of-words-in-c-sharp
            int count = 0;
            bool wasInWord = false;
            bool inWord = false;

            for (int i = 0; i < test.Length; i++)
            {
                if (inWord)
                {
                    wasInWord = true;
                }

                if (Char.IsWhiteSpace(test[i]))
                {
                    if (wasInWord)
                    {
                        count++;
                        wasInWord = false;
                    }
                    inWord = false;
                }
                else
                {
                    inWord = true;
                }
            }

            // Check to see if we got out with seeing a word
            if (wasInWord)
            {
                count++;
            }

            return count;
        }

        public static bool IsMonthYear(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            DateTime dt;
            text = text.Trim();
            CultureInfo ci = new CultureInfo("en-us");
            if (DateTime.TryParseExact(text, "m.yyyy", ci, DateTimeStyles.None, out dt) ||
                DateTime.TryParseExact(text, "MMM yyyy", ci, DateTimeStyles.None, out dt) ||
                DateTime.TryParseExact(text, "MMMM yyyy", ci, DateTimeStyles.None, out dt) ||
                DateTime.TryParseExact(text, "yyyy", ci, DateTimeStyles.None, out dt))
            {
                return true;
            }
            return false;
        }

        public static bool IsMonth(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            DateTime dt;
            CultureInfo ci = new CultureInfo("en-us");
            if (DateTime.TryParseExact(text.Trim(), "MMM", ci, DateTimeStyles.None, out dt) ||
                DateTime.TryParseExact(text.Trim(), "MMMM", ci, DateTimeStyles.None, out dt))
            {
                return true;
            }
            return false;
        }

        public static bool IsMonthYearOrPresent(ref string text, ref bool tillPresent)
        {
            if (!IsMonthYear(text))
            {
                //not April 2018 or Apr 2018 or 04.2018 or 2018
                foreach (var current in PresentTimeConstants)
                {
                    int i = text.IndexOf(current, StringComparison.InvariantCultureIgnoreCase);
                    if (i == 0)
                    {
                        text = text.Substring(0, current.Length);
                        tillPresent = true;
                        return true;
                    }
                }
            }
            else
            {
                tillPresent = false;
                return true;
            }

            return false;
        }

        public static bool SplitToTextAndDateRange(ref string value, ref string dateRange, ref DateTime? lastDate)
        {
            //aim: to parse constructions like "Belarussian state university (2010-2015)" or "Manhattan project (Aug 2017-Nov 2018)" or "BSUIR, 2016"
            char[] seps = { '(', ',' };
            int i = value.IndexOfAny(seps);
            if (i != -1)
            {
                string dateCandidate = null;
                int j = value.IndexOf(')', i + 1);
                if (j != -1)
                {
                    dateCandidate = value.Substring(i + 1, j - i - 1);
                }
                else
                {
                    dateCandidate = value.Substring(i + 1);
                }

                string[] range = dateCandidate.Split(StringUtils.DateRangeSeparators);
                string graduationYear = null;
                for (int k = 0; k < range.Count(); k++)
                {
                    if (StringUtils.IsYear(range[k]))
                        graduationYear = range[k];
                }
                if (!string.IsNullOrEmpty(graduationYear))
                {
                    lastDate = DateParser.ParseDateTime(graduationYear);
                    if (lastDate.HasValue)
                    {
                        value = value.Substring(0, i).Trim();
                        dateRange = dateCandidate;
                        return true;
                    }
                }
            }
            return false;
        }

        //https://stackoverflow.com/questions/541954/how-would-you-count-occurrences-of-a-string-actually-a-char-within-a-string
        //CountChar2 actually
        public static int CountChar(this string str, char substr)
        {
            int count = 0;
            foreach (var c in str)
                if (c == substr)
                    ++count;

            return count;
        }

        public static bool IsStringLiteral(string str)
        {
            //very basic check if string looks like a human name, and does not contain any extra characters
            return str.IndexOfAny(_deniedStringLiteralCharacters) < 0;
        }

        public static bool IsHyperlink(string str)
        {
            return str.IndexOf("http", StringComparison.InvariantCultureIgnoreCase) >= 0;
        }

        //https://stackoverflow.com/questions/59217/merging-two-arrays-in-net
        public static T[] Combine<T>(params IEnumerable<T>[] items) =>
                            items.SelectMany(i => i).ToArray();

    }
}