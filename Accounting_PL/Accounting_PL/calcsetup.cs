using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Accounting_PL
{
    class calcsetup
    {
        // The first Monday the first week of business(or available data) at this location
        // 53 week years occur the year following a fiscal year ending on the 28th of December

        int weekNum = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        int weekYearNum = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(new DateTime(DateTime.Now.Year, 12, 31), CultureInfo.CurrentCulture.DateTimeFormat.CalendarWeekRule, CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek);


    }
}
