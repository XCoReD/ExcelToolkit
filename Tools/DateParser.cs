using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

//original: ResumeParser/Utils/DateParser.cs
namespace Tools
{
    public static class DateParser
    {
        public static bool IsNow(string date)
        {
            if (StringUtils.CompareRoughly(date, "now") ||
                StringUtils.CompareRoughly(date, "nowadays") ||
                StringUtils.CompareRoughly(date, "present") ||
                StringUtils.CompareRoughly(date, "present time") ||
                StringUtils.CompareRoughly(date, "present days") ||
                StringUtils.CompareRoughly(date, "the present") ||
                StringUtils.CompareRoughly(date, "current"))
                return true;

            return false;

        }
        public static DateTime? ParseDateTime(string date)
        {
            DateTime dateTime;
            var redundantSuffixes = new []{"th", "nd", "rd"};

            date = redundantSuffixes.Aggregate(date, (current, s) => current.Replace(s, string.Empty)).Trim();

            if(IsNow(date))
                return DateTime.Now;

            int dashPos = date.IndexOf('/');
            if (dashPos > 0)
            {
                //"09 / 2014 - present"
                date = date.Replace('/', '.').Replace(" ", "");
            }

            if (DateTime.TryParse(date, out dateTime))
            {
				return DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
            }

            if (DateTime.TryParse(date, new CultureInfo("en-us"), DateTimeStyles.None, out dateTime))
            {
				return DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
            }

            if (DateTime.TryParseExact(date, "d.m.yyyy", new CultureInfo("en-us"), DateTimeStyles.None, out dateTime))
            {
                return DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
            }

            int year;
            //DT: fix "years" like 1, 2, 3 :)
			if (Int32.TryParse(date, out year) && year > 1980 && year < 2100)
			{
				return new DateTime(year, 1, 1, 0, 0, 0, DateTimeKind.Utc);
			}

            DateTime.TryParseExact(
                Regex.Replace(date, @"(\w+ \d+)\w+ (\w+ \d+)", "$1 $2"),
                "dddd d MMMM yyyy",
                DateTimeFormatInfo.InvariantInfo,
                DateTimeStyles.None, out dateTime);

			return dateTime == DateTime.MinValue ? (DateTime?)null : DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
        }

        public static bool ContainsDate(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            foreach (var item in text.Replace('–', '-').Replace(',', '-').Replace(" ", "").Replace(":", "").Split('-'))
            {
                string test = item.Trim();
                DateTime? dt = ParseDateTime(test);
                if (dt.HasValue)
                    return true;
            }

            return false;
        }
    }
}
