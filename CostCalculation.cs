using System;
using Excel= Microsoft.Office.Interop;
namespace TDD
{
    public class CostCalculation
    {
        Excel.Excel.Range currentFind ;
        Excel.Excel.Range firstFind ;
        Excel.Excel.Application xlApp;
        Excel.Excel.Workbook xlWorkBook;
        Excel.Excel.Worksheet xlWorkSheetFirst;
        Excel.Excel.Worksheet xlWorkSheetSecond;
        Excel.Excel.Worksheet xlWorkSheetThird;
        public string strExcelFilePath= @"C:\Users\Sami\Desktop\TDD\IntakeAndRelease.xlsx";
        public object misValue = System.Reflection.Missing.Value;
        public Tuple<Excel.Excel.Application, Excel.Excel.Workbook, Excel.Excel.Worksheet, Excel.Excel.Worksheet, Excel.Excel.Worksheet> opendatasource()
        {
                xlApp = new Excel.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(strExcelFilePath
                    , 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheetFirst = (Excel.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheetSecond = (Excel.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                xlWorkSheetThird = (Excel.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
          return  Tuple.Create(xlApp, xlWorkBook, xlWorkSheetFirst, xlWorkSheetSecond, xlWorkSheetThird);
        }
        public Tuple<DateTime, DateTime, int, int , string> cal10euroeachday(DateTime dtstartdate, DateTime dtenddate, int intPersonID )
        {
            TimeSpan t = dtenddate - dtstartdate;
            double NrOfDays = t.TotalDays;
            int intAmout = Convert.ToInt32(NrOfDays) * 10;
            if (intAmout > 228) { intAmout = 228; }
            var result = Tuple.Create(dtstartdate, dtenddate, intPersonID, intAmout, "Partial Payment");
            return result;
        }
        public Tuple<DateTime, DateTime, int, int, string, bool> calservicefees(DateTime dtstartdate, DateTime dtenddate, int intPersonID, string strServiceName, Excel.Excel.Worksheet xlWorkSheetSecond, Excel.Excel.Worksheet  xlWorkSheetThird)
        {
            TimeSpan t = dtenddate - dtstartdate;
            double NrOfDays = t.TotalDays;
            int cellRow = 0;
            string strInsurance=null;
            int x = 0;
            for (x = 1; x <= 9; x++)
            {
                if (xlWorkSheetSecond.Cells[x,1].text==(strServiceName))
                {
                    cellRow = x;
                    break;
                }
            }
           
            int intAmout = Convert.ToInt32(NrOfDays) * Convert.ToInt32(xlWorkSheetSecond.Cells[cellRow, 2].Value.ToString());
            for(int p=1; p < 27; p++)
            {
                if (xlWorkSheetThird.Cells[p , 1].text==(intPersonID.ToString()))
                {
                    strInsurance = xlWorkSheetThird.Cells[ p,7].text;
                   
                    break;
                }
            }
            if (strInsurance == "No Insurance")
            {
                intAmout += (Convert.ToInt32(NrOfDays) * 10);
                var result = Tuple.Create(dtstartdate, dtenddate, intPersonID, intAmout, "Individually paid", false);
                return result;
            }
            else {
                var result = Tuple.Create(dtstartdate, dtenddate, intPersonID, intAmout, strInsurance, true);
                return result;
            } 
        }
        public  Tuple<DateTime, DateTime, int, bool> CalPaymentPerPerson(int intPersonID , Excel.Excel.Worksheet xlWorkSheetFirst, Excel.Excel.Worksheet xlWorkSheetSecond, Excel.Excel.Worksheet xlWorkSheetThird)
        {
            DateTime dtStart ;
            DateTime dtEnd ;
            string strServiceName = null;
            currentFind = null;
            firstFind = null;
            bool blCheckInsurance = true;
            var result = Tuple.Create(DateTime.MinValue, DateTime.MinValue, intPersonID, blCheckInsurance);
            Excel.Excel.Range Fruits = xlWorkSheetFirst.get_Range("A1", "A42");
            // Specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = Fruits.Find(intPersonID, misValue,
                Excel.Excel.XlFindLookIn.xlValues, Excel.Excel.XlLookAt.xlPart,
                Excel.Excel.XlSearchOrder.xlByRows, Excel.Excel.XlSearchDirection.xlNext, false,
                misValue, misValue);
                
            while (currentFind != null)
            {
                string strRowNumber = currentFind.get_Address();
                strRowNumber = strRowNumber.Substring(strRowNumber.LastIndexOf('$') + 1);
                dtStart = Convert.ToDateTime(xlWorkSheetFirst.Cells[Convert.ToInt32(strRowNumber), 6].Value.ToString());
                dtEnd = Convert.ToDateTime(xlWorkSheetFirst.Cells[Convert.ToInt32(strRowNumber), 7].Value.ToString());
                strServiceName = xlWorkSheetFirst.Cells[Convert.ToInt32(strRowNumber), 3].Value.ToString();
                Tuple<DateTime, DateTime, int, int, string, bool> userData = calservicefees(dtStart, dtEnd, intPersonID, strServiceName ,xlWorkSheetSecond , xlWorkSheetThird );
                blCheckInsurance = userData.Item6;
                return  result = Tuple.Create(dtStart, dtEnd, intPersonID, blCheckInsurance);
                //Check if person has insurance
                if (blCheckInsurance == true)
                {
                    cal10euroeachday(dtStart, dtEnd, intPersonID);
                    return result = Tuple.Create(dtStart, dtEnd, intPersonID, blCheckInsurance);
                }
                // Keep track of the first range found. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }
                // If didn't move to a new range, done.
                else if (currentFind.get_Address(Excel.Excel.XlReferenceStyle.xlA1)
                        == firstFind.get_Address(Excel.Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }
                currentFind = Fruits.FindNext(currentFind);
            }
            return result;
        }
    }
}
