using ChurchBudgetReportGenerator.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ChurchBudgetReportGenerator
{
    /// <summary>
    /// Excel file handler that deals with various operations, i.e. reading/updating, etc.
    /// </summary>
    public static class ExcelFileHandler
    {
        /// <summary>
        /// Updates Excel file with Monthly Summary Profit and Loss (PL) Report data
        /// </summary>
        public static void UpdateExcelFileWithMonthlySummaryPLReport(string filePath, AccountMonthlySummaryData monthlySummaryData, ConfigData configData)
        {

            // Row and Column numbers of for this report
            var headerRow = 2;
            var fundHeaderRow = headerRow + 4;
            var fundStartAmountRow = fundHeaderRow + 2;
            var dataStartRow = fundStartAmountRow + 2;
            int fundEndAmountRow;

            var accountNumberColumn = 2;
            var accountNameColumn = 3;
            var fundStartColumn = 4;

            // Open Excel file
            FileInfo output = Utils.GetFileInfo(filePath, false);
            ExcelPackage pck = new ExcelPackage(output);

            // Create the workbook
            // Add the Content Profit and Loss (P&L) monthly report worksheet

            // WorkSheet name should look like "P&L_Jan", "P&L_Feb", etc
            var sheetName =  $"P&L_{new string(monthlySummaryData.MonthName.Take(3).ToArray())}";            
            if (pck.Workbook.Worksheets[sheetName] != null)
            {
                pck.Workbook.Worksheets.Delete(sheetName);
            }
            var ws = pck.Workbook.Worksheets.Add(sheetName);

            // Add Header Info
            AddHeaderRow(ws, headerRow, accountNumberColumn, configData.Funds.Count(), 14,
                "Ukrainian Greek-Catholic Church \"Zarvanycia\", Seattle, WA");
            AddHeaderRow(ws, headerRow + 1, accountNumberColumn, configData.Funds.Count(), 12,
                "Statement of Activities	");
            AddHeaderRow(ws, headerRow + 2, accountNumberColumn, configData.Funds.Count(), 12,
                $"{monthlySummaryData.MonthName} {monthlySummaryData.Year}");

            // Columns width
            ws.Column(1).Width = 2.5;
            ws.Column(accountNumberColumn).Width = 10;
            ws.Column(accountNameColumn).Width = 25;
            
            var fundEndColumn = fundStartColumn;

            // Add Fund Headers
            var amountColumnWidth = 14;
            AddCellValueBold(ws, fundStartAmountRow, accountNameColumn, "Cash Beginning of Period:");
            foreach (var fund in configData.Funds)
            {
                ws.Column(fundEndColumn).Width = amountColumnWidth;
                AddCellValueCenterBold(ws, fundHeaderRow, fundEndColumn, fund.Name);
                                
                // Fund starting amount
                AddCellValueCenterBold(ws, fundStartAmountRow, fundEndColumn, fund.StartingPeriodAmount,
                    ExcelNumberFormats.Accounting);

                fundEndColumn++;
            }
            AddCellValueCenterBold(ws, fundHeaderRow, fundEndColumn, "Total");
            ws.Column(fundEndColumn).Width = amountColumnWidth;

            // Formula for Funds starting amount
            AddCellFormulaRightBold(ws, fundStartAmountRow, fundEndColumn,
                    $"SUM({GetColumnName(fundStartColumn - 1)}{fundStartAmountRow}:" +
                    $"{GetColumnName(fundEndColumn - 2)}{fundStartAmountRow})",
                    ExcelNumberFormats.Accounting);
            
            // Group by account types
            var rowListOfAccountTypes = new List<int>();
            var dataRow = dataStartRow;
            foreach (var accountType in configData.AccountTypes)
            {
                AddCellValueBold(ws, dataRow, accountNumberColumn, $"{accountType}:");
                dataRow++;

                dataRow = AddAcountsTransactionsDataForType(ws, dataRow, accountType, configData, monthlySummaryData, fundStartColumn);
                rowListOfAccountTypes.Add(dataRow);

                // Add some space between types
                ws.Cells[dataRow, accountNameColumn].Value = "";
                dataRow++;                
            }

            // Add final Net formulas
            AddCellValueLeftBold(ws, dataRow, accountNameColumn, "Net: Income Gain / (Loss)");
            fundEndColumn = fundStartColumn;
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
            AddCellValueBold(ws, dataRow, accountNameColumn, "Cash End of Period:");
            fundEndColumn = fundStartColumn;
            foreach (var fund in configData.Funds)
            {
                AddCellFormulaRightBold(ws, dataRow, fundEndColumn,
                    $"{GetColumnName(fundEndColumn - 1)}{fundStartAmountRow}+" +
                    $"{GetColumnName(fundEndColumn - 1)}{netRow}",
                    ExcelNumberFormats.Accounting);
                fundEndColumn++;
            }
            AddCellFormulaRightBold(ws, dataRow, fundEndColumn,
                   $"SUM({GetColumnName(fundStartColumn - 1)}{dataRow}:" +
                   $"{GetColumnName(fundEndColumn - 2)}{dataRow})",
                   ExcelNumberFormats.Accounting);

            // Updating placeholder to store last row for Fund End Amount
            fundEndAmountRow = dataRow;

            ws.Calculate();
            pck.Save();

            fundEndColumn = fundStartColumn;
            foreach (var fund in configData.Funds)
            {
                var fundEndAmountStr = ws.Cells[fundEndAmountRow, fundEndColumn].Value.ToString();
                decimal fundEndAmount = string.IsNullOrWhiteSpace(fundEndAmountStr)
                    ? 0m
                    : decimal.Parse(fundEndAmountStr);
                fund.StartingPeriodAmount = fundEndAmount;
                fundEndColumn++;
            }
        }

        /// <summary>
        /// Updates Excel file with Monthly Summary Report for Bulletin
        /// </summary>
        public static void UpdateExcelFileWithMonthlySummaryForBulletin(string filePath, AccountMonthlySummaryData monthlySummaryData, ConfigData configData)
        {
            // Row and Column numbers of for this report
            var headerRow = 1;
            var accountTypeHeaderRow = headerRow + 1;
            var fundHeaderRow = accountTypeHeaderRow + 1;

            var accountAmountStartRow = fundHeaderRow + 1;
            int accountAmountEndRow = accountAmountStartRow;

            var accountNameStartColumn = 1;
            var accountNameColumnWidth = 20;
            var accountAmountColumnWidth = 14;
            var fundNameStartColumn = accountNameStartColumn + 1;

            var accountTypeSpacerColumnWidth = 4;

            // Open Excel file
            FileInfo output = Utils.GetFileInfo(filePath, false);
            ExcelPackage pck = new ExcelPackage(output);

            // Create the workbook
            // WorkSheet name should look like "Bulletin_Jan", "Bulletin_Feb", etc
            var sheetName = $"Bulletin_{new string(monthlySummaryData.MonthName.Take(3).ToArray())}";
            if (pck.Workbook.Worksheets[sheetName] != null)
            {
                pck.Workbook.Worksheets.Delete(sheetName);
            }
            var ws = pck.Workbook.Worksheets.Add(sheetName);

            // Add Header Info
            var fundsWithNonZeroAccountAmounts = monthlySummaryData.Accounts.Where(a => a.Amount > 0)
                .GroupBy(a => a.Fund).Count();
            AddHeaderRow(ws, headerRow, accountNameStartColumn, fundsWithNonZeroAccountAmounts + 2, 12,
                $"Financial activity for {monthlySummaryData.MonthName} {monthlySummaryData.Year}");
            
            // Add Account Type Headers
            var accountTypeColumn = accountNameStartColumn;
            var fundNameColumn = fundNameStartColumn;
            var accountNameAdded = false;   // We need to add Account Name per Fund only at the account adding step
            foreach (var accountType in configData.AccountTypes)
            {
                // Summarize transactions data per account types and funds for non-zero amounts
                var typeSummary = monthlySummaryData.Accounts
                    .Where(a => a.Account.Type == accountType)
                    .Where(a => a.Amount > 0)
                    .GroupBy(a => a.Fund);
                
                AddHeaderRow(ws, accountTypeHeaderRow, accountTypeColumn, typeSummary.Count() - 2, 10, $"{accountType}");
                accountNameAdded = false;

                // For each Fund add its data
                foreach (var fundData in typeSummary)
                {
                    // Add Fund Name Header
                    AddCellValueCenterUnderline(ws, fundHeaderRow, fundNameColumn, fundData.Key.Name);
                    ws.Column(fundNameColumn).Width = accountAmountColumnWidth;

                    // Add Accounts
                    var accountRow = accountAmountStartRow;
                    foreach (var accountData in fundData)
                    {
                        // Add Account Name only if it's not added at the beginning of the Account Type
                        if (!accountNameAdded)
                        {
                            var accountNameColumn = fundNameColumn - 1;
                            AddCellValueLeft(ws, accountRow, accountNameColumn, accountData.Account.Name);
                            ws.Column(accountNameColumn).Width = accountNameColumnWidth;
                        }                        
                        
                        // Add Amount
                        AddCellValueRight(ws, accountRow, fundNameColumn, accountData.Amount,
                            ExcelNumberFormats.Accounting);

                        if (accountAmountEndRow < accountRow)
                        {
                            accountAmountEndRow = accountRow;
                        }
                        accountRow++;
                    }
                    accountNameAdded = true;
                    fundNameColumn++;
                }
                accountTypeColumn = accountTypeColumn + typeSummary.Count() + 2;

                // Add spacer column
                ws.Column(accountTypeColumn - 1).Width = accountTypeSpacerColumnWidth;
                AddCellValueCenter(ws, accountTypeHeaderRow, accountTypeColumn - 1, "");
                
                fundNameColumn = accountTypeColumn + 1;                                
            }

            // Add Total Sum Functions at the very last row            
            var formulaRow = accountAmountEndRow + 1;
            AddCellValueLeftBold(ws, formulaRow, fundNameStartColumn - 1, "Total");
            for (int c = fundNameStartColumn; c < fundNameColumn; c++)
            {
                var firstRowCellValue = ws.Cells[accountAmountStartRow, c].Value;
                if (firstRowCellValue != null && !string.IsNullOrWhiteSpace(firstRowCellValue.ToString()))
                {
                    if(int.TryParse(firstRowCellValue.ToString(), out int firstRowCellValueInt))
                    {
                        AddCellFormulaRightBold(ws, formulaRow, c,
                           $"SUM({GetColumnName(c - 1)}{accountAmountStartRow}:" +
                           $"{GetColumnName(c - 1)}{accountAmountEndRow})",
                           ExcelNumberFormats.Accounting);
                    }
                }
            }

            // Apply Table formatting
            //using (ExcelRange Rng = ws.Cells[$"A1:{GetColumnName(fundNameColumn)}{formulaRow}"])
            //{
            //    var tblcollection = ws.Tables;
            //    ExcelTable table = tblcollection.Add(Rng, $"tbl{sheetName}");
            //    table.TableStyle = TableStyles.Medium4;
            //}


            ws.Calculate();
            pck.Save();

            
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
                            var accountType = ConvertStrToAccountType(accountTypeStr);
                            returnData.AccountTypes.Add(accountType);

                            // All the Accounts are stored in columns 2 - 50
                            var accountStartColumn = 2;
                            var accountEndColumn = 50;
                            for (int i = accountStartColumn; i <= accountEndColumn; i++)
                            {
                                var columnValue = worksheet.Cells[row, i].Value;
                                var valueStr = columnValue?.ToString();

                                if (!string.IsNullOrWhiteSpace(valueStr))
                                {
                                    returnData.Accounts.Add(GetAccountData(valueStr, accountType));
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
                        var accountType = ConvertStrToAccountType(worksheet.Cells[row, 3].Value.ToString());                           
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
                            Description = worksheet.Cells[row, 5].GetValue<string>(),
                            Amount = decimal.Parse(worksheet.Cells[row, 6].Value.ToString())
                        });
                    }
                    row++;
                }
            }
            return returnList;
        }

        private static AccountType ConvertStrToAccountType(string accountTypeStr)
        {
            if (string.IsNullOrWhiteSpace(accountTypeStr))
            {
                throw new ArgumentNullException("Unable to convert empty string to AccountType enum value");
            }
            return (AccountType)Enum.Parse(typeof(AccountType), accountTypeStr);
        }

        private static Account GetAccountData(string accountStr, AccountType accountType)
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
            int row, AccountType accountType, ConfigData configData, AccountMonthlySummaryData monthlySummaryData,
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

        private static void AddCellValueCenterUnderline(ExcelWorksheet ws, int row, int column, object value, string numberFormat = null)
        {
            ws.Cells[row, column].Value = value;
            ws.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[row, column].Style.Font.UnderLine = true;
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
    }
}