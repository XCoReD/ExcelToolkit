using ExcelFunctions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var text = RussianDeclination.Decline("Главный электрик", Declination.Dative, false);
            //var text = RussianDeclination.Decline("заяц", Declination.Dative, false);

            double d = 525.17;
            Declination tc = Declination.Dative;
            bool useUSD = true;

            var text2 = RuDateAndMoneyConverter.CurrencyToTxt(d, tc, false, useUSD);

            Debug.WriteLine(text2);


        }
    }
}
