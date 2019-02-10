using System;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Transaction Data record stored in Excel File
    /// </summary>
    public class TransactionData : AccountData
    {
        public DateTime TimeStamp { get; set; }
        public string Description { get; set; }
    }
}
