using ChurchBudgetReportGenerator.Models;
using System;
using System.Collections.Generic;
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
            var monthlyGroupedExpenses = GetExpensesGrouppedByYearAndMonth(transactionData);

            // For each month, generate summary data for that month
            Console.WriteLine("Updating file with Monthly Summary Reports");
            foreach (var monthlySummaryData in monthlyGrouppedTransactions)
            {
                var monthNameReduced = new string(monthlySummaryData.MonthName.Take(3).ToArray());
                ExcelFileHandler.UpdateExcelFileWithSummaryReport(filePath, monthlySummaryData, configData,
                    $"P&L_{monthNameReduced}",
                    $"{monthlySummaryData.MonthName} {monthlySummaryData.Year}", string.Empty, string.Empty, string.Empty);

                ExcelFileHandler.UpdateExcelFileWithMonthlySummaryForBulletin(filePath, monthlySummaryData, configData);

                var monthExpenses = monthlyGroupedExpenses.FirstOrDefault(m => m.MonthNumber == monthlySummaryData.MonthNumber);
                ExcelFileHandler.UpdateExcelFileWithMonthlyExpensesSummary(filePath, monthExpenses, configData,
                    $"Expenses_{monthNameReduced}", $"For {monthlySummaryData.MonthName} {monthlySummaryData.Year}");
            }

            // Create Yearly report
            // Reset cash funds amount for yearly
            Console.WriteLine("Updating file with Yearly Reports");
            configData.ResetFundsReportStartingAmount();
            var yearsData = GetTransactionsGrouppedByYear(transactionData);
            foreach (var yearData in yearsData)
            {
                ExcelFileHandler.UpdateExcelFileWithSummaryReport(filePath, yearData, configData,
                    $"Annual_{yearData.Year}", "Annual Report",
                    $"For Jan 1 {yearData.Year} to Dec 31 {yearData.Year}",
                    " (1/1)", " (12/31)");
            }

            Console.WriteLine("File update completed.");
        }

        /// <summary>
        /// Return list of Account Summary data to be placed into Excel Reports
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

        public static ICollection<AccountTransactions> GetAccountTransactionsSummary(ICollection<TransactionData> transactions)
        {
            var returnData = new List<AccountTransactions>();

            // Get transactions grouped by account Number
            var accountGrouppedTransactions = transactions.GroupBy(t => t.Account.Number);

            // For each account, grouping by Description
            foreach (var accountGrouppedTransaction in accountGrouppedTransactions)
            {
                var descriptionGrouppedTransactions = accountGrouppedTransaction
                    .GroupBy(t => t.Description);

                foreach (var descriptionGrouppedTransaction in descriptionGrouppedTransactions)
                {
                    returnData.Add
                        (
                        new AccountTransactions
                        {
                            Account = descriptionGrouppedTransaction.FirstOrDefault().Account,
                            Description = descriptionGrouppedTransaction.FirstOrDefault().Description,
                            Amount = descriptionGrouppedTransaction.Sum(t => t.Amount)
                        }
                        );
                }
            }
            return returnData;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="transactionData"></param>
        /// <returns></returns>
        public static IEnumerable<AccountTransactionsData> GetExpensesGrouppedByYearAndMonth(ICollection<TransactionData> transactionData)
        {
            var expenses = transactionData
                .Where(t => t.Account.Type == AccountType.Expenses)
                .ToList();

            return expenses
                .GroupBy(t => new { t.TimeStamp.Year, t.TimeStamp.Month })
                .Select(g => new AccountTransactionsData()
                {
                    Year = g.Key.Year,
                    MonthNumber = g.Key.Month,
                    AccountTransactions = GetAccountTransactionsSummary(g.ToList())
                }
                );
        }

        /// <summary>
        /// Return all transactions grouped by each Month
        /// </summary>
        public static IEnumerable<AccountSummaryData> GetTransactionsGrouppedByYearAndMonth(ICollection<TransactionData> transactionData)
        {
            return transactionData.GroupBy(t => new { t.TimeStamp.Year, t.TimeStamp.Month})
                .Select(g => new AccountSummaryData()
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
        public static IEnumerable<AccountSummaryData> GetTransactionsGrouppedByYear(ICollection<TransactionData> transactionData)
        {
            return transactionData.GroupBy(t => new { t.TimeStamp.Year})
                .Select(g => new AccountSummaryData()
                {
                    Year = g.Key.Year,
                    Accounts = GetAccountDataSummary(g.ToList())
                }
                );
        }
    }
}
