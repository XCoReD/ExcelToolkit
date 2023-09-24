using LingvoNET;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions
{
    public class RussianDeclination
    {
        enum TruncateIO
        {
            None,
            Last,
            First
        };
        static RussianDeclination _this;
        static char[] _separators = { '-', '—', '–', ' ' };
        public static string Decline(string arg, Declination decl, bool upperFirstLetter)
        {
            if (_this == null)
                _this = new RussianDeclination();

            return _this.DeclineImpl(arg, decl, upperFirstLetter);
        }

        public static string DeclineFIO(string arg, Declination decl, bool truncateFirstNameParentName, bool reverseFirstNameParentName)
        {
            if (_this == null)
                _this = new RussianDeclination();

            TruncateIO param = truncateFirstNameParentName && reverseFirstNameParentName ? TruncateIO.First : (truncateFirstNameParentName ? TruncateIO.Last : TruncateIO.None);
            return _this.DeclineImpl(arg, decl, false, param);
        }
        RussianDeclination()
        {
            //_lingvo = new LingvoNET.Schema();
        }
        string DeclineImpl(string arg, Declination decl, bool upperFirstLetter, TruncateIO truncateIO = TruncateIO.None)
        {
            if (string.IsNullOrEmpty(arg))
                return arg;

            int iStart = 0, iPrevSeparator = -1, numWords = 0;
            string result = null, tail = null;
            Case c = ToCase(decl);
            for (int i = 0; i < arg.Length; ++i)
            {
                int si = Array.IndexOf(_separators, arg[i]);
                if (si != -1)
                {
                    bool truncate = truncateIO != TruncateIO.None && numWords >= 1;
                    if (iPrevSeparator == -1)
                    {
                        if(truncate && truncateIO == TruncateIO.First)
                            tail += DeclineWord(arg.Substring(iStart, i - iStart), c, truncate);
                        else
                            result += DeclineWord(arg.Substring(iStart, i - iStart), c, truncate);
                        ++numWords;
                    }
                    if(!truncate)
                        result += arg[i];
                    iPrevSeparator = i;
                }
                else
                {
                    if (iPrevSeparator != -1)
                    {
                        iPrevSeparator = -1;
                        iStart = i;
                    }
                }
            }

            if(iPrevSeparator == -1)
            {
                ++numWords;
                bool truncate = truncateIO != TruncateIO.None && numWords >= 1;
                var last = arg.Substring(iStart, arg.Length - iStart);
                if (truncate && truncateIO == TruncateIO.First)
                    tail += DeclineWord(last, c, truncate);
                else
                    result += DeclineWord(last, c, truncate);
            }

            if (truncateIO == TruncateIO.First)
                result = (tail + result).TrimEnd();

            if (upperFirstLetter)
            {
                return Char.ToUpper(result[0]) + result.Substring(1);
            }

            return result;
        }

        string DeclineWord(string arg, Case c, bool truncateToFirstChar)
        {
            if (string.IsNullOrEmpty(arg))
                return arg;

            if(truncateToFirstChar)
            {
                return Char.ToUpper(arg[0]) + ".";
            }

            var f = LingvoNET.Analyser.FindSimilarSourceForm(arg);

            Analyser.WordType type = f.Type;    // Analyser.WordType.Noun;
            /*if (f.Count() == 1)
                type = f.First().Type;
            else if(f.Count() > 1)
            {
                Debug.WriteLine($"Multiple word forms are get for {arg}, taking the first one");
                type = f.First().Type;
            }*/

            Gender g = Gender.M;

            switch(type)
            {
                case Analyser.WordType.Noun:
                    {
                        Noun n = LingvoNET.Nouns.FindOne(arg);
                        if (n == null)
                            n = LingvoNET.Nouns.FindSimilar(arg);

                        if (n == null)
                            return "N/A";

                        g = n.Gender;
                        return n[c];
                    }
                case Analyser.WordType.Adjective:
                    {
                        Adjective a = LingvoNET.Adjectives.FindOne(arg);
                        if (a == null)
                            a = LingvoNET.Adjectives.FindSimilar(arg);

                        if (a == null)
                            return "N/A";

                        return a[c, g];
                    }
            }
            return arg;
        }
        Case ToCase(Declination d)
        {
            switch(d)
            {
                case Declination.Accusative: return Case.Accusative;
                case Declination.Dative: return Case.Dative;
                case Declination.Genitive: return Case.Genitive;
                case Declination.Instrumental: return Case.Instrumental;
                case Declination.Nominative: return Case.Nominative;
                case Declination.Prepositional: return Case.Locative;

                default: Debug.Assert(false);
                    return Case.Undefined;
            }
        }
    }
}
