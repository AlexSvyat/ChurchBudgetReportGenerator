using System.Collections.Generic;
using System.Globalization;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Account Summary monthly data
    /// </summary>
    public class AccountMonthlySummaryData
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

        public AccountMonthlySummaryData()
        {
            Accounts = new List<AccountData>();
        }
    }
}
