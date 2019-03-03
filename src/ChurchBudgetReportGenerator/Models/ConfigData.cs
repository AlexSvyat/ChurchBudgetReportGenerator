using System.Collections.Generic;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Config data class with all the Accounts and all the Funds data
    /// </summary>
    public class ConfigData
    {
        public List<AccountType> AccountTypes { get; set; }
        public List<Account> Accounts { get; set; }
        public List<Fund> Funds { get; set; }

        public ConfigData()
        {
            AccountTypes = new List<AccountType>();
            Accounts = new List<Account>();
            Funds = new List<Fund>();
        }

        public void ResetFundsReportStartingAmount()
        {
            foreach (var fund in Funds)
            {
                fund.ReportStartingAmount = fund.BeginningAmount;
            }
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
            BeginningAmount = startAmount;
            ReportStartingAmount = BeginningAmount;
        }

        public string Name { get; set; }

        /// <summary>
        /// Fund Starting Amount for Report
        /// </summary>
        public decimal ReportStartingAmount { get; set; }

        /// <summary>
        /// Fund Beginning Amount at the beginning of the year
        /// </summary>
        public decimal BeginningAmount { get; set; }

    }
}
