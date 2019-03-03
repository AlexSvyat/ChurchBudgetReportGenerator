using System.Collections.Generic;
using System.Globalization;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Account Summary data, can be used for monthly or yearly data
    /// </summary>
    public class AccountSummaryData
    {
        public int MonthNumber { get; set; }
        public string MonthName {
            get
            {
                return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(MonthNumber);
            }
        }

        public int Year { get; set; }
        public List<AccountData> Accounts { get; set; }

        public AccountSummaryData()
        {
            Accounts = new List<AccountData>();
        }
    }
}
