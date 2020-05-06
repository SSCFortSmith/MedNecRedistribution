using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace MedNecRedistribution
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 1 && Path.GetExtension(args[0]).Contains("xl") || 1==1)
            {
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                foreach (String file in args)
                {
                    Workbook wb = Excel.Workbooks.Open(file);
                    //Excel.DisplayAlerts = false;
                    Excel.Visible = true;
                    try
                    {
                        wb.Worksheets["BY COID"].Delete();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    working(wb.Sheets["working"]);
                    bypass(wb.Sheets["bypass"]);
                    DateTime localDate = DateTime.Now;
                    string lcdt = localDate.ToString("MMddyy");
                    //lcdt = lcdt.Replace("{}", Path.GetFileNameWithoutExtension(file).Substring(Path.GetFileNameWithoutExtension(file).IndexOf(".")+1, 2).Replace(".","").PadLeft(2,'0'));
                    Console.WriteLine("'" + file + "' has been successfully converted.");
                    wb.SaveAs(@"S:\ARSC Scripting Backup\Med Nec\" + "Med Nec Report " + lcdt + ".xlsx");
                    wb.Close();
                    Excel.DisplayAlerts = true;
                    wb = null;
                }
                Excel.Quit();
                Excel = null;
            }
            else
            {
                if (args.Count() == 1)
                    Console.WriteLine("Incorrect File Type: " + Path.GetExtension(args[0]));
                else
                    Console.WriteLine("Requires one file");
            }
            Console.ReadKey();
        }
        static void bypass(Worksheet res)
        {
            res.Columns["A:E"].Insert(XlInsertShiftDirection.xlShiftToRight);
            res.Cells[1, 1].Value2 = "COID";
            res.Cells[1, 2].Value2 = "ACCT";
            res.Cells[1, 3].Value2 = "986";
            res.Cells[1, 4].Value2 = "953";
            res.Cells[1, 5].Value2 = "NOTE";
            popValBypass(res);
        }
        static void working(Worksheet res)
        {    
            res.Columns["A:E"].Insert(XlInsertShiftDirection.xlShiftToRight);
            res.Cells[1, 1].Value2 = "COID";
            res.Cells[1, 2].Value2 = "ACCOUNT";
            res.Cells[1, 3].Value2 = "986";
            res.Cells[1, 4].Value2 = "953";
            res.Cells[1, 5].Value2 = "NOTE";
            popValWorking(res);
        }
        static void popValBypass(Worksheet res)
        {
            int records = 0;
            Range start = res.Range["H1"]; // first cell
            Range bottom = res.Range["H" + (res.UsedRange.Rows.Count + 1)]; // upper bound (maximum)
            Range end = bottom.End[XlDirection.xlUp]; // true last cell (maximum)
            Range column = res.Range[start, end]; //
            records = end.Row;
            String acct = "", note = ""; ;
            for (int i = 2; i <= records; i++)
            {
                    res.Cells[i, 1].NumberFormat = "@";
                    res.Cells[i, 2].NumberFormat = "@";
                    res.Cells[i, 3].NumberFormat = "#########0.00";
                    res.Cells[i, 4].NumberFormat = "#########0.00";
                    res.Cells[i, 1].Value2 = res.Cells[i, 8].Value2; // coid
                    acct = Convert.ToString(res.Cells[i, 12].Value2); // account
                    if (acct.Length > 7)
                        acct = acct.Substring(3);
                    if (acct.Length < 7)
                        acct = acct.PadLeft(7, '0');
                    res.Cells[i, 2].Value2 = acct;// acct
                    if (Convert.ToString(res.Cells[i, 6].Value2) == "N")
                        res.Cells[i, 6].Value2 = 0;
                    res.Cells[i, 3].Value2 = 0; // 986
                    res.Cells[i, 4].Value2 = Convert.ToDouble(res.Cells[i, 15].Value2); // 953
                    // end account processing
                    //begin note creation
                    note = "noncovered charges, " + res.Cells[i, 19].Value2 + ", " + res.Cells[i, 14].Value2 + ", requesting 953 for $" + res.Cells[i, 15].Value2;
                    res.Cells[i, 5].Value2 = note;

            }
            res.Range["F:V"].Delete();
        }
        static void popValWorking(Worksheet res)
        {
            int records = 0;
            Range start = res.Range["J1"]; // first cell
            Range bottom = res.Range["J" + (res.UsedRange.Rows.Count + 1)]; // upper bound (maximum)
            Range end = bottom.End[XlDirection.xlUp]; // true last cell (maximum)
            Range column = res.Range[start, end]; //
            records = end.Row;
            String acct = "", note = ""; ;
            for (int i = 2; i <= records; i++)
            {
                if (Convert.ToString(res.Cells[i,7].Value2) == "GZ")
                {
                    res.Cells[i, 1].NumberFormat = "@";
                    res.Cells[i, 2].NumberFormat = "@";
                    res.Cells[i, 3].NumberFormat = "#########0.00";
                    res.Cells[i, 4].NumberFormat = "#########0.00";
                    res.Cells[i, 1].Value2 = res.Cells[i, 10].Value2; // coid
                    acct = Convert.ToString(res.Cells[i, 14].Value2); // account
                    if (acct.Length > 7)
                        acct = acct.Substring(3);
                    if (acct.Length < 7)
                        acct = acct.PadLeft(7, '0');
                    res.Cells[i, 2].Value2 = acct;// coid

                    res.Cells[i, 3].Value2 = res.Cells[i, 8].Value2; // 986
                    res.Cells[i, 4].Value2 = res.Cells[i, 9].Value2; // 953
                    // end account processing
                    //begin note creation
                    note = "*ADJUSTED* DR2, RT1, " + res.Cells[i, 16].Value2 + ", $" + res.Cells[i, 8].Value2 + ", No appeal - Per " + res.Cells[i, 6].Value2 + 
                        " no dx and no valid ABN.  Moved to non covered and added GZ modifier in SSI. Requesting 986 for $"+Convert.ToString(res.Cells[i,8].Value2)+
                        " and 953 for $"+Convert.ToString(res.Cells[i,9].Value2);
                    res.Cells[i, 5].Value2 = note;
                }
                else
                {
                    res.Rows[i].Delete();
                    i--;
                    records--;
                }
            }
            res.Range["F:V"].Delete();
        }
    }
}
