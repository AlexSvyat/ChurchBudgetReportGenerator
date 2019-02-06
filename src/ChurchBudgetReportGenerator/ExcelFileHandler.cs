using ChurchBudgetReportGenerator.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChurchBudgetReportGenerator
{
    /// <summary>
    /// Excel file handler that deals with various operations, i.e. reading/updating, etc.
    /// </summary>
    public static class ExcelFileHandler
    {
        // Row and Column numbers of various purpose
        private static int HeaderRow { get; set; } = 2;
        private static int FundHeaderRow { get; set; } = HeaderRow + 4;
        private static int FundStartAmountRow { get; set; } = FundHeaderRow + 2;
        private static int DataStartRow { get; set; } = FundStartAmountRow + 2;
        private static int FundEndAmountRow { get; set; }

        private static int AccountNumberColumn { get; set; } = 2;
        private static int AccountNameColumn { get; set; } = 3;
        private static int FundStartColumn { get; set; } = 4;

        /// <summary>
        /// Updates Excel file with Monthly Summary Report data
        /// </summary>
        public static void UpdateExcelFileWithMonthlySummaryReport(string filePath, AccountMonthlySummaryData monthlySummaryData, ConfigData configData)
        {
            // Create the workbook
            FileInfo output = Utils.GetFileInfo(filePath, false);
            ExcelPackage pck = new ExcelPackage(output);

            // Add the Content Profit and Loss (P&L) monthly report worksheet
            var sheetName = GetPLWorksheetName(monthlySummaryData.MonthName);
            if (pck.Workbook.Worksheets[sheetName] != null)
            {
                pck.Workbook.Worksheets.Delete(sheetName);
            }
            var ws = pck.Workbook.Worksheets.Add(sheetName);

            // Add Header Info
            AddHeaderRow(ws, HeaderRow, AccountNumberColumn, configData.Funds.Count(), 14,
                "Ukrainian Greek-Catholic Church \"Zarvanycia\", Seattle, WA");
            AddHeaderRow(ws, HeaderRow + 1, AccountNumberColumn, configData.Funds.Count(), 12,
                "Statement of Activities	");
            AddHeaderRow(ws, HeaderRow + 2, AccountNumberColumn, configData.Funds.Count(), 12,
                $"{monthlySummaryData.MonthName} {monthlySummaryData.Year}");

            // Columns width
            ws.Column(1).Width = 2.5;
            ws.Column(AccountNumberColumn).Width = 10;
            ws.Column(AccountNameColumn).Width = 25;
            
            var fundEndColumn = FundStartColumn;

            // Add Fund Headers
            AddCellValueBold(ws, FundStartAmountRow, AccountNameColumn, "Cash Beginning of Period:");
            foreach (var fund in configData.Funds)
            {
                ws.Column(fundEndColumn).Width = 12;
                AddCellValueCenterBold(ws, FundHeaderRow, fundEndColumn, fund.Name);
                                
                // Fund starting amount
                AddCellValueCenterBold(ws, FundStartAmountRow, fundEndColumn, fund.StartingPeriodAmount,
                    ExcelNumberFormats.Accounting);

                fundEndColumn++;
            }
            AddCellValueCenterBold(ws, FundHeaderRow, fundEndColumn, "Total");
            ws.Column(fundEndColumn).Width = 12;

            // Formula for Funds starting amount
            AddCellFormulaRightBold(ws, FundStartAmountRow, fundEndColumn,
                    $"SUM({GetColumnName(FundStartColumn - 1)}{FundStartAmountRow}:" +
                    $"{GetColumnName(fundEndColumn - 2)}{FundStartAmountRow})",
                    ExcelNumberFormats.Accounting);
            
            // Group by account types
            var rowListOfAccountTypes = new List<int>();
            var dataRow = DataStartRow;
            foreach (var accountType in configData.AccountTypes)
            {
                AddCellValueBold(ws, dataRow, AccountNumberColumn, $"{accountType}:");
                dataRow++;

                dataRow = AddAcountsTransactionsDataForType(ws, dataRow, accountType, configData, monthlySummaryData, FundStartColumn);
                rowListOfAccountTypes.Add(dataRow);

                // Add some space between types
                ws.Cells[dataRow, AccountNameColumn].Value = "";
                dataRow++;                
            }

            // Add final Net formulas
            AddCellValueLeftBold(ws, dataRow, AccountNameColumn, "Net: Income Gain / (Loss)");
            fundEndColumn = FundStartColumn;
            var netRow = dataRow;
            foreach (var fund in configData.Funds)
            {
                AddCellFormulaRightBold(ws, netRow, fundEndColumn, 
                    $"{GetColumnName(fundEndColumn-1)}{rowListOfAccountTypes.First()-1}-" +
                    $"{GetColumnName(fundEndColumn-1)}{rowListOfAccountTypes.Last()-1}",
                    ExcelNumberFormats.Accounting);
                fundEndColumn++;
            }
            AddCellFormulaRightBold(ws, netRow, fundEndColumn,
                   $"{GetColumnName(fundEndColumn - 1)}{rowListOfAccountTypes.First() - 1}-" +
                   $"{GetColumnName(fundEndColumn - 1)}{rowListOfAccountTypes.Last() - 1}",
                   ExcelNumberFormats.Accounting);
            dataRow++;
            dataRow++;

            // Add Fund End Amount of the Period
            AddCellValueBold(ws, dataRow, AccountNameColumn, "Cash End of Period:");
            fundEndColumn = FundStartColumn;
            foreach (var fund in configData.Funds)
            {
                AddCellFormulaRightBold(ws, dataRow, fundEndColumn,
                    $"{GetColumnName(fundEndColumn - 1)}{FundStartAmountRow}+" +
                    $"{GetColumnName(fundEndColumn - 1)}{netRow}",
                    ExcelNumberFormats.Accounting);
                fundEndColumn++;
            }
            AddCellFormulaRightBold(ws, dataRow, fundEndColumn,
                   $"SUM({GetColumnName(FundStartColumn - 1)}{dataRow}:" +
                   $"{GetColumnName(fundEndColumn - 2)}{dataRow})",
                   ExcelNumberFormats.Accounting);

            // Updating placeholder to store last row for Fund End Amount
            FundEndAmountRow = dataRow;

            ws.Calculate();
            pck.Save();

            fundEndColumn = FundStartColumn;
            foreach (var fund in configData.Funds)
            {
                var fundEndAmountStr = ws.Cells[FundEndAmountRow, fundEndColumn].Value.ToString();
                decimal fundEndAmount = string.IsNullOrWhiteSpace(fundEndAmountStr)
                    ? 0m
                    : decimal.Parse(fundEndAmountStr);
                fund.StartingPeriodAmount = fundEndAmount;
                fundEndColumn++;
            }
        }        

        /// <summary>
        /// Returns Config data about Accounts and Funds from the Excel file
        /// </summary>
        public static ConfigData GetConfigDataFromWorksheet(string filePath)
        {
            var returnData = new ConfigData();
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // Get specific worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Settings"];

                int row = 2; // Starting Row
                int maxRowsToProcess = 21;

                while (row < maxRowsToProcess)
                {
                    var firstColumnValue = worksheet.Cells[row, 1].Value;

                    // Get Account info
                    if (row < 4)
                    {
                        var accountTypeStr = firstColumnValue?.ToString();
                        if (!string.IsNullOrWhiteSpace(accountTypeStr))
                        {
                            returnData.AccountTypes.Add(accountTypeStr);

                            // All the Accounts are stored in columns 2 - 50
                            var accountStartColumn = 2;
                            var accountEndColumn = 50;
                            for (int i = accountStartColumn; i <= accountEndColumn; i++)
                            {
                                var columnValue = worksheet.Cells[row, i].Value;
                                var valueStr = columnValue?.ToString();

                                if (!string.IsNullOrWhiteSpace(valueStr))
                                {
                                    returnData.Accounts.Add(GetAccountData(valueStr, accountTypeStr));
                                }
                            }
                        }
                    }

                    // Get Funds info, they started after row 9
                    if (row > 8)
                    {
                        var secondColumnValue = worksheet.Cells[row, 2].Value;

                        var fundNameStr = firstColumnValue?.ToString();
                        if (!string.IsNullOrWhiteSpace(fundNameStr))
                        {
                            decimal fundStartingAmount = secondColumnValue == null
                                ? 0m
                                : decimal.Parse(secondColumnValue.ToString());
                            returnData.Funds.Add(new Fund(fundNameStr, fundStartingAmount));
                        }
                        returnData.Funds.OrderBy(f => f);
                    }
                    row++;
                }
            }
            return returnData;
        }

        /// <summary>
        /// Read Excel file and returns Transactions Data from a worksheet
        /// </summary>
        public static ICollection<TransactionData> GetTransactionDataFromExcelFile(string filePath, ConfigData configData)
        {
            var returnList = new List<TransactionData>();
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // Get specific worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Transactions"];

                var isEmptyRowFound = false;
                int row = 2; // Starting Row
                while (!isEmptyRowFound)
                {
                    var accountStr = worksheet.Cells[row, 4].Value;
                    if (accountStr == null)
                    {
                        isEmptyRowFound = true;
                    }
                    else
                    {
                        var accountSplitted = accountStr.ToString().Split(':');
                        var accountNumber = accountSplitted[0];
                        var accountName = accountSplitted[1];
                        var accountType = worksheet.Cells[row, 3].Value.ToString();
                        var accountFundFromTransaction = worksheet.Cells[row, 2].Value.ToString();

                        // Get and make sure that Fund name matches from the Config data
                        var fund = configData.Funds
                            .FirstOrDefault(f => f.Name.Equals(accountFundFromTransaction, StringComparison.InvariantCultureIgnoreCase));

                        if (fund == null)
                        {
                            throw new InvalidOperationException($"Unable to find a fund name '{accountFundFromTransaction}' " +
                                $"for a transaction account number '{accountNumber}' and type '{accountType}'");
                        }

                        returnList.Add(new TransactionData
                        {
                            TimeStamp = worksheet.Cells[row, 1].GetValue<DateTime>(),
                            Fund = fund,
                            Account = new Account(accountNumber, accountName, accountType),                            
                            Amount = decimal.Parse(worksheet.Cells[row, 5].Value.ToString())
                        });
                    }
                    row++;
                }
            }
            return returnList;
        }
               
        private static Account GetAccountData(string accountStr, string accountType = null)
        {
            var accountSplitted = accountStr.ToString().Split(':');
            return new Account(accountSplitted[0], accountSplitted[1], accountType);
        }

        private static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
            {
                value += letters[index / letters.Length - 1];
            }
            value += letters[index % letters.Length];

            return value;
        }

        private static int AddAcountsTransactionsDataForType(ExcelWorksheet ws,
            int row, string accountType, ConfigData configData, AccountMonthlySummaryData monthlySummaryData,
            int fundStartColumn)
        {
            var accountNumberColumn = 2;
            var accountNameColumn = 3;
            var fundStartRow = row;
            var fundAccountEndColumn = 0;

            // Add a record for each account
            foreach (var configAccount in configData.Accounts.Where(a => a.Type == accountType))
            {
                // For each Fund find the right amount from transactions
                fundAccountEndColumn = fundStartColumn;
                foreach (var fund in configData.Funds)
                {
                    ws.Cells[row, fundAccountEndColumn].Value = fund;

                    var accountFromTransaction = monthlySummaryData
                        .Accounts.FirstOrDefault(a => a.Account.Number == configAccount.Number
                        && a.Fund == fund);

                    decimal accountAmount = 0m;
                    if (accountFromTransaction?.Amount > 0)
                    {
                        accountAmount = accountFromTransaction.Amount;
                    }
                    AddCellValueRight(ws, row, accountNumberColumn, configAccount.Number);
                    AddCellValueLeft(ws, row, accountNameColumn, configAccount.Name);
                    AddCellValueRight(ws, row, fundAccountEndColumn, accountAmount,
                        ExcelNumberFormats.AccountingWithoutDollarAndDecimalFraction);

                    fundAccountEndColumn++;
                }
                // Add Total formula at the end of row
                AddCellFormulaRightBold(ws, row, fundAccountEndColumn,
                    $"SUM({GetColumnName(fundStartColumn - 1)}{row}:" +
                    $"{GetColumnName(fundAccountEndColumn - 2)}{row})", ExcelNumberFormats.AccountingWithoutDollar);

                row++;
            }

            // Adding total row at the bottom of account type
            AddCellValueCenterBold(ws, row, accountNameColumn, $"Total {accountType}:");
            var formulaColumn = fundStartColumn;
            foreach (var fund in configData.Funds)
            {
                // Formula for each Fund column on the bottom
                AddCellFormulaRightItaliDoubleTop(ws, row, formulaColumn,
                    $"SUM({GetColumnName(formulaColumn - 1)}{fundStartRow}:" +
                    $"{GetColumnName(formulaColumn - 1)}{row -1})", ExcelNumberFormats.AccountingWithoutDollarAndDecimalFraction);
                formulaColumn++;
            }

            // Last Formula at the bottom on very right of account type
            AddCellFormulaRightBoldItalicDoubleTop(ws, row, formulaColumn,
                    $"SUM({GetColumnName(formulaColumn - 1)}{fundStartRow}:" +
                    $"{GetColumnName(formulaColumn - 1)}{row - 1})", ExcelNumberFormats.Accounting);
            row++;

            return row;
        }

        private static void AddCellValueBold(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.Font.Bold = true;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellValueCenterBold(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[row, column].Style.Font.Bold = true;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellValueCenter(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellValueLeft(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellValueLeftBold(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells[row, column].Style.Font.Bold = true;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellValueRight(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellValueRightBold(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[row, column].Style.Font.Bold = true;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellFormulaRightBold(ExcelWorksheet ws, int row, int column, string formula, string numberFormat = null)
        {
            ws.Cells[row, column].Formula = formula;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[row, column].Style.Font.Bold = true;
            AddCellNumberFormat(ws, row, column, numberFormat);                      
        }
    
        private static void AddCellFormulaRightBoldItalicDoubleTop(ExcelWorksheet ws, int row, int column, string formula, string numberFormat = null)
        {
            ws.Cells[row, column].Formula = formula;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[row, column].Style.Font.Bold = true;
            ws.Cells[row, column].Style.Font.Italic = true;
            ws.Cells[row, column].Style.Border.Top.Style = ExcelBorderStyle.Double;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellFormulaRightItaliDoubleTop(ExcelWorksheet ws, int row, int column, string formula, string numberFormat = null)
        {
            ws.Cells[row, column].Formula = formula;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[row, column].Style.Font.Italic = true;
            ws.Cells[row, column].Style.Border.Top.Style = ExcelBorderStyle.Double;
            AddCellNumberFormat(ws, row, column, numberFormat);
        }

        private static void AddCellNumberFormat(ExcelWorksheet ws, int row, int column, string numberFormat)
        {
            // Apply Accounting Format
            if (!string.IsNullOrWhiteSpace(numberFormat))
            {
                ws.Cells[row, column].Style.Numberformat.Format = numberFormat;
            }
        }

        private static void AddHeaderRow(ExcelWorksheet ws, int row, int column, int amountOfCellToMerge, int fontSize, string cellValue)
        {
            ws.Cells[row, column].Value = cellValue;
            var cellRange = $"{GetColumnName(column - 1)}{row}:" +
                $"{GetColumnName(column + amountOfCellToMerge + 1)}{row}";
            ws.Cells[cellRange].Merge = true;
            ws.Cells[cellRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells[row, column].Style.Font.Name = "Copperplate Gothic Bold";
            ws.Cells[row, column].Style.Font.Size = fontSize;
        }

        private static string GetPLWorksheetName(string monthName)
        {
            return $"P&L_{new string(monthName.Take(3).ToArray())}";
        }
    }
}