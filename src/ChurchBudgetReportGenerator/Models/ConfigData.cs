using System.Collections.Generic;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Config data class with all the Accounts and all the Funds data
    /// </summary>
    public class ConfigData
    {
        public List<string> AccountTypes { get; set; }
        public List<Account> Accounts { get; set; }
        public List<Fund> Funds { get; set; }

        public ConfigData()
        {
            AccountTypes = new List<string>();
            Accounts = new List<Account>();
            Funds = new List<Fund>();
        }
    }

    public class Fund
    {
        public Fund(string fundName)
        {
            Name = fundName;
        }

        public Fund(string fundName, decimal startAmount)
        {
            Name = fundName;
            StartingPeriodAmount = startAmount;
        }

        public string Name { get; set; }
        public decimal StartingPeriodAmount { get; set; }
    }
}
