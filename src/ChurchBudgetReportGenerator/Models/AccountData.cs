namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Base class for account data
    /// </summary>
    public class AccountData
    {
        public Fund Fund { get; set; }
        public Account Account { get; set; }
        public decimal Amount { get; set; }

        public AccountData()
        {
            Account = new Account();
        }
    }

    /// <summary>
    /// Account object
    /// </summary>
    public class Account
    {        
        public Account()
        {
        }

        public Account(string accountNumberStr, string accountName, AccountType accountType)
        {
            Number = int.Parse(accountNumberStr);
            Name = accountName;
            Type = accountType;
        }

        public AccountType Type { get; set; }
        public int Number { get; set; }
        public string Name { get; set; }
    }
}