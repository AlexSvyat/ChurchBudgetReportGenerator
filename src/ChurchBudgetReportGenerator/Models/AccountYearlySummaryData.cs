using System.Collections.Generic;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Account Summary yearly data
    /// </summary>
    public class AccountYearlySummaryData
    {
        public int Year { get; set; }
        public List<AccountData> Accounts { get; set; }

        public AccountYearlySummaryData()
        {
            Accounts = new List<AccountData>();
        }
    }
}
