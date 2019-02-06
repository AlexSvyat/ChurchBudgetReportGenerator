using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChurchBudgetReportGenerator.Models
{
    /// <summary>
    /// Excel Number Formats for various formatting
    /// </summary>
    public static class ExcelNumberFormats
    {
        /// <summary>
        /// Number Format "Accounting" with dollar sign, i.e. "$  2,375.00" or "$  (1,000.00)"
        /// </summary>
        public static string Accounting { get; set; } = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

        /// <summary>
        /// Number Format "Accounting" without dollar sign, i.e. "1,000.00"
        /// </summary>
        public static string AccountingWithoutDollar { get; set; } = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";

        /// <summary>
        /// Number Format "Accounting" without dollar sign and without decimal fraction, i.e. "1,000"
        /// </summary>
        public static string AccountingWithoutDollarAndDecimalFraction { get; set; } = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)";
    }
}
