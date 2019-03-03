using System.Collections.Generic;
using System.Globalization;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Account Transactions data, can be used for monthly or yearly data
    /// </summary>
    public class AccountTransactionsData
    {
        public int MonthNumber { get; set; }
        public string MonthName {
            get
            {
                return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(MonthNumber);
            }
        }

        public int Year { get; set; }
        public ICollection<AccountTransactions> AccountTransactions { get; set; }

        public AccountTransactionsData()
        {
            AccountTransactions = new List<AccountTransactions>();
        }
    }

    public class AccountTransactions
    {
        public Account Account { get; set; }
        public string Description { get; set; }
        public decimal Amount { get; set; }
    }
}
