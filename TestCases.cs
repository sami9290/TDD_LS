using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop;
namespace TDD
{
    //fds
    [TestFixture]
    class TestCases
    {
        //gfd
        //test rrt
        Excel.Excel.Application xlApp;
        Excel.Excel.Workbook xlWorkBook;
        Excel.Excel.Worksheet xlWorkSheetFirst;
        Excel.Excel.Worksheet xlWorkSheetSecond;
        Excel.Excel.Worksheet xlWorkSheetThird;
        public void opendatasource()
        {
            CostCalculation cc = new CostCalculation();
            cc.opendatasource();
            Tuple<Excel.Excel.Application, Excel.Excel.Workbook, Excel.Excel.Worksheet,
                Excel.Excel.Worksheet, Excel.Excel.Worksheet> getData = cc.opendatasource();
            xlApp = getData.Item1;
            xlWorkBook = getData.Item2;
            xlWorkSheetFirst = getData.Item3;
            xlWorkSheetSecond = getData.Item4;
            xlWorkSheetThird = getData.Item5;
        }

        [TestCase]
        public void TC01()
        {
            CostCalculation cc = new CostCalculation();
            cc.opendatasource();
            Tuple<Excel.Excel.Application, Excel.Excel.Workbook, Excel.Excel.Worksheet,
                Excel.Excel.Worksheet, Excel.Excel.Worksheet> getData = cc.opendatasource();
            xlApp = getData.Item1;
            xlWorkBook = getData.Item2;
            xlWorkSheetFirst = getData.Item3;
            xlWorkSheetSecond = getData.Item4;
            xlWorkSheetThird = getData.Item5;
        }

        [TestCase]
        public void TC02()
        {
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string>(
                        Convert.ToDateTime("4/1/2015"), Convert.ToDateTime("4/9/2015"),
                        4, 90, "Partial Payment");

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.cal10euroeachday(Convert.ToDateTime(4 / 1 / 2015),
                Convert.ToDateTime( 4 / 9 / 2015),4));
        }
        [TestCase]
        public void TC03()
        {
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string>(
                      Convert.ToDateTime("4/1/2015"), Convert.ToDateTime("4/9/2015"),
                      4, 80, "Partial Payment");

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.cal10euroeachday(Convert.ToDateTime("4/1/2015"),
                Convert.ToDateTime("4/9/2015"), 4));
        }
        [TestCase]
        public void TC04()
        {
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string>(
               DateTime.Now, DateTime.Now.AddDays(2), 4, 20, "Partial Payment");

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.cal10euroeachday(DateTime.Now,
                DateTime.Now.AddDays(2), 4));
        }

        [TestCase]
        public void TC05()
        {
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string>(
               DateTime.Now, DateTime.Now.AddDays(35), 4, 350, "Partial Payment");

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.cal10euroeachday(DateTime.Now,
                DateTime.Now.AddDays(35), 4));
        }

        [TestCase]
        public void TC06()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string, bool>(
            Convert.ToDateTime("4 / 1 / 2015"), Convert.ToDateTime("4 / 9 / 2015"), 4, 80, "TK", true);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.calservicefees(Convert.ToDateTime("4 / 1 / 2015"),
                Convert.ToDateTime("4 / 9 / 2015"), 4, "Emergency Department", xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close(false);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }
  
        [TestCase]
        public void TC08()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string, bool>(
     Convert.ToDateTime("1/1/2015"), Convert.ToDateTime("8/1/2015"), 12, 6996, "Individually paid", false);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.calservicefees(Convert.ToDateTime("1/1/2015"),
                Convert.ToDateTime("8/1/2015"), 12, "Cataract surgery outcome", xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }
        [TestCase]
        public void TC09()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, int, string, bool>(
     Convert.ToDateTime("12/1/2015"), Convert.ToDateTime("12/12/2015"), 
     10, 440, "AXA", true);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.calservicefees(Convert.ToDateTime("12/1/2015"),
                Convert.ToDateTime("12/12/2015"), 10, "Preventive Care", xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }
        [TestCase]
        public void TC10()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime,int, bool>(
                         Convert.ToDateTime("4 / 1 / 2015"),
                         Convert.ToDateTime("4 / 9 / 2015"), 4, false);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.CalPaymentPerPerson( 4 ,
                xlWorkSheetFirst, xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }
        [TestCase]
        public void TC11()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, bool>(
                         Convert.ToDateTime("11/19/2015"),
                         Convert.ToDateTime("12/1/2015"), 4, true);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.CalPaymentPerPerson(4,
                xlWorkSheetFirst, xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }
        	
         [TestCase]
        public void TC12()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, bool>(
                         Convert.ToDateTime("1/1/2015"),
                         Convert.ToDateTime("8/1/2015"), 12, false);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.CalPaymentPerPerson(12,
                xlWorkSheetFirst, xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }


        [TestCase]
        public void TC13()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, bool>(
                         Convert.ToDateTime("1/1/2015"), 
                         Convert.ToDateTime("8/1/2015"), 12, false);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.CalPaymentPerPerson(12,
                xlWorkSheetFirst, xlWorkSheetSecond, xlWorkSheetThird));

            var TupleCompareServiceFees = new Tuple<DateTime, DateTime, int, int, string, bool>(
      Convert.ToDateTime("1/1/2015"), Convert.ToDateTime("8/1/2015"),
      12, 6996, "Individually paid", false);
            
            Assert.AreEqual(TupleCompareServiceFees,
                cc.calservicefees(Convert.ToDateTime("1/1/2015"), Convert.ToDateTime("8/1/2015"),
                12, "Cataract surgery outcome", xlWorkSheetSecond, xlWorkSheetThird));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }

        [TestCase]
        public void TC14()
        {
            opendatasource();
            var TupleCompare = new Tuple<DateTime, DateTime, int, bool>(
                        Convert.ToDateTime("4/1/2015"), Convert.ToDateTime("4/9/2015"),
                        4, true);

            CostCalculation cc = new CostCalculation();
            Assert.AreEqual(TupleCompare, cc.CalPaymentPerPerson(4, xlWorkSheetFirst,
                xlWorkSheetSecond, xlWorkSheetThird));

            var TupleCompareServiceFees = new Tuple<DateTime, DateTime,
                int, int, string, bool>(
      Convert.ToDateTime("4/1/2015"), Convert.ToDateTime("4/9/2015"),
      4, 80, "TK", true);

            Assert.AreEqual(TupleCompareServiceFees, 
                cc.calservicefees(Convert.ToDateTime("4/1/2015"), Convert.ToDateTime("4/9/2015"),
                4, "Emergency Department", xlWorkSheetSecond, xlWorkSheetThird));
            var TupleCompare10Europerday = new Tuple<DateTime, DateTime, int, int, string>(
                      Convert.ToDateTime("4/1/2015"), Convert.ToDateTime("4/9/2015"),
                      4, 80, "Partial Payment");

            Assert.AreEqual(TupleCompare10Europerday, cc.cal10euroeachday(Convert.ToDateTime("4/1/2015"),
                Convert.ToDateTime("4/9/2015"), 4));
            xlWorkBook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
        }
    }
}
