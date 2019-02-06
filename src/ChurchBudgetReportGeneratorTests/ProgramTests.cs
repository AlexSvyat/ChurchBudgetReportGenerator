using Microsoft.VisualStudio.TestTools.UnitTesting;
using ChurchBudgetReportGenerator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ChurchBudgetReportGenerator.Models;

namespace ChurchBudgetReportGenerator.Tests
{
    [TestClass()]
    public class ProgramTests
    {
        [TestMethod()]
        public void GroupTransactionDataByMonth_2Month_Success_Test()
        {
            var testTrans = new List<TransactionData>
            {
                new TransactionData { TimeStamp = DateTime.Parse("1-1-19"), Account = new Account("1", "TestAccount"), Amount = 2 },
                new TransactionData { TimeStamp = DateTime.Parse("1-2-19"), Account = new Account("1", "TestAccount"), Amount = 3 },
                new TransactionData { TimeStamp = DateTime.Parse("2-1-19"), Account = new Account("1", "TestAccount"), Amount = 4 }
            };

            var grouppedByMonth = Program.GetTransactionsGrouppedByYearAndMonth(testTrans);
            Assert.AreEqual(2, grouppedByMonth.Count(), "Expected to get data for 2 months!");

            // Getting Total Amount for January
            Assert.AreEqual(5, grouppedByMonth.Where(i => i.MonthNumber == 1).Sum(i => i.Accounts.Sum(t => t.Amount)));

            // Getting Total Amount for February
            Assert.AreEqual(4, grouppedByMonth.Where(i => i.MonthNumber == 2).Sum(i => i.Accounts.Sum(t => t.Amount)));
        }

        [TestMethod()]
        public void GetAccountSummary_Success_Test()
        {
            var testTrans = new List<TransactionData>
            {
                new TransactionData { TimeStamp = DateTime.Parse("1-1-19"), Account = new Account("3001", "TestAccount"), Amount = 2 },
                new TransactionData { TimeStamp = DateTime.Parse("1-2-19"), Account = new Account("3002", "TestAccount"), Amount = 3 },
                new TransactionData { TimeStamp = DateTime.Parse("1-2-19"), Account = new Account("3001", "TestAccount"), Amount = 4 },
                new TransactionData { TimeStamp = DateTime.Parse("1-2-19"), Account = new Account("3003", "TestAccount"), Amount = 5.05m },
                new TransactionData { TimeStamp = DateTime.Parse("1-31-19"), Account = new Account("3002", "TestAccount"), Amount = 3.01m },

                new TransactionData { TimeStamp = DateTime.Parse("2-1-19"), Account = new Account("4001", "TestAccount"), Amount = 4 }
            };

            var grouppedByMonth = Program.GetTransactionsGrouppedByYearAndMonth(testTrans);
            var summaryDataForJanuary = grouppedByMonth.FirstOrDefault(i => i.MonthNumber == 1);
            var summaryDataForFebruary = grouppedByMonth.FirstOrDefault(i => i.MonthNumber == 2);

            Assert.AreEqual("January", summaryDataForJanuary.MonthName);
            Assert.AreEqual(3, summaryDataForJanuary.Accounts.Count);
            Assert.AreEqual(6, summaryDataForJanuary.Accounts.FirstOrDefault(d => d.Account.Number == 3001).Amount);
            Assert.AreEqual(6.01m, summaryDataForJanuary.Accounts.FirstOrDefault(d => d.Account.Number == 3002).Amount);
            Assert.AreEqual(5.05m, summaryDataForJanuary.Accounts.FirstOrDefault(d => d.Account.Number == 3003).Amount);

            Assert.AreEqual("February", summaryDataForFebruary.MonthName);
            Assert.AreEqual(1, summaryDataForFebruary.Accounts.Count);
            Assert.AreEqual(4, summaryDataForFebruary.Accounts.FirstOrDefault(d => d.Account.Number == 4001).Amount);
        }
    }
}