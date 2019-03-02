using ChurchBudgetReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ChurchBudgetReportGenerator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var filePath = args[0];

            // Set the output directory to the SampleApp folder where the app is running from. 
            Utils.OutputDir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            
            if (!File.Exists(filePath))
            {
                filePath = Path.GetFullPath(filePath);

                if (!File.Exists(filePath))
                {
                    throw new ArgumentException($"Unable to locate provide file path '{filePath}'");
                }
            }
            Console.WriteLine($"Start updating Excel file '{filePath}'");

            // Get all Funds and Accounts info
            Console.WriteLine("Getting Config Data");
            var configData = ExcelFileHandler.GetConfigDataFromWorksheet(filePath);

            // Read all transactions from Excel 
            Console.WriteLine("Getting Transactions Data");
            var transactionData = ExcelFileHandler.GetTransactionDataFromExcelFile(filePath, configData);
            if (!transactionData.Any())
            {
                throw new InvalidOperationException($"Unable to get any transactions data from file '{filePath}'");
            }

            // Group by months
            Console.WriteLine("Grouping Transactions Data by months");
            var monthlyGrouppedTransactions = GetTransactionsGrouppedByYearAndMonth(transactionData);

            // For each month, generate summary data for that month
            Console.WriteLine("Updating file with Monthly Summary Reports");
            foreach (var monthlyGroup in monthlyGrouppedTransactions)
            {
                ExcelFileHandler.UpdateExcelFileWithMonthlySummaryPLReport(filePath, monthlyGroup, configData);
                ExcelFileHandler.UpdateExcelFileWithMonthlySummaryForBulletin(filePath, monthlyGroup, configData);
            }

            // Create Yearly report
            var yearlyData = GetTransactionsGrouppedByYear(transactionData);
            ExcelFileHandler.UpdateExcelFileWithYearySummaryReport(filePath, yearlyData, configData);

            Console.WriteLine("File update completed.");
        }

        /// <summary>
        /// Return list of Account Summary data to be placed into final Excel Report
        /// </summary>
        public static List<AccountData> GetAccountDataSummary(IEnumerable<TransactionData> transactions)
        {
            var returnData = new List<AccountData>();

            // Get transactions grouped by account Number
            var accountGrouppedTransactions = transactions.GroupBy(t => t.Account.Number);

            // For each account, grouping by Funds
            foreach (var accountGrouppedTransaction in accountGrouppedTransactions)
            {
                var fundGrouppedTransactions = accountGrouppedTransaction
                    .GroupBy(t => t.Fund);

                foreach (var fundGrouppedTransaction in fundGrouppedTransactions)
                {
                    returnData.Add
                        (
                        new AccountData
                        {
                            Account = fundGrouppedTransaction.FirstOrDefault().Account,
                            Fund = fundGrouppedTransaction.FirstOrDefault().Fund,
                            Amount = fundGrouppedTransaction.Sum(t => t.Amount)
                        }
                        );
                }
            }
            return returnData;
        }

        /// <summary>
        /// Return all transactions grouped by each Month
        /// </summary>
        public static IEnumerable<AccountMonthlySummaryData> GetTransactionsGrouppedByYearAndMonth(ICollection<TransactionData> transactionData)
        {
            return transactionData.GroupBy(t => new { t.TimeStamp.Year, t.TimeStamp.Month})
                .Select(g => new AccountMonthlySummaryData()
                {
                    Year = g.Key.Year,
                    MonthNumber = g.Key.Month,
                    Accounts = GetAccountDataSummary(g.ToList())
                }
                ); 
        }

        /// <summary>
        /// Return all transactions grouped by Year
        /// </summary>
        public static IEnumerable<AccountMonthlySummaryData> GetTransactionsGrouppedByYear(ICollection<TransactionData> transactionData)
        {
            return transactionData.GroupBy(t => new { t.TimeStamp.Year})
                .Select(g => new AccountMonthlySummaryData()
                {
                    Year = g.Key.Year,
                    Accounts = GetAccountDataSummary(g.ToList())
                }
                );
        }
    }
}
