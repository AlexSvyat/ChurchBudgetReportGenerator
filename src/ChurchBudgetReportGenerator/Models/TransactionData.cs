using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Transaction Data record stored in Excel File
    /// </summary>
    public class TransactionData : AccountData
    {
        public DateTime TimeStamp { get; set; }
    }
}
