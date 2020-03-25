using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace EComerceAnalysis {
    public class Program {
        static void Main (string[] args) {
            using (var package = new ExcelPackage (new FileInfo ("Online_Retail.xlsx"))) {

                // var list = package.Workbook.Worksheets["Online Retail"].ImportExcelToList<FundraiserStudentListModel> ();
                // var bb=list.Count();
                // package.Workbook.Worksheets["Online Retail"].Cells.Count();
                //  Console.WriteLine(package.Workbook.Worksheets["Online Retail"].);
                var firstSheet = package.Workbook.Worksheets["Online Retail"];
                Console.WriteLine ("Sheet 1 Data");
                Console.WriteLine ($"Cell A2 Value   : {firstSheet.Cells["A2"].Text}");
                Console.WriteLine ($"Cell A2 Color   : {firstSheet.Cells["B2"].Text}");
                Console.WriteLine ($"Cell B2 Formula : {firstSheet.Cells["C2"].Text}");
                Console.WriteLine ($"Cell B2 Value   : {firstSheet.Cells["D2"].Text}");
                Console.WriteLine ($"Cell B2 Border  : {firstSheet.Cells["E2"].Text}");
                Console.WriteLine ($"Cell B2 Border  : {firstSheet.Cells["F2"].Text}");
                Console.WriteLine ($"Cell B2 Border  : {firstSheet.Cells["G2"].Text}");
                Console.WriteLine ($"Cell B2 Border  : {firstSheet.Cells["H2"].Text}");
                Console.WriteLine (firstSheet.Dimension.Rows);
                Console.WriteLine ("");
                var list = new List<FundraiserStudentListModel> ();
                for (int i = 2; i <= firstSheet.Dimension.Rows; i++) {
                    // Console.WriteLine (firstSheet.Cells["A" + i].Text);
                    // Console.WriteLine (firstSheet.Cells["B" + i].Text);
                    // Console.WriteLine (firstSheet.Cells["C" + i].Text);
                    // Console.WriteLine (float.Parse (firstSheet.Cells["D" + i].Text));
                    // Console.WriteLine (DateTime.Parse (firstSheet.Cells["E" + i].Text));
                    // Console.WriteLine (float.Parse (firstSheet.Cells["F" + i].Text));
                    // Console.WriteLine (firstSheet.Cells["G" + i].Text);
                    // Console.WriteLine (firstSheet.Cells["H" + i].Text);
                    list.Add (new FundraiserStudentListModel () {
                        InvoiceNo = firstSheet.Cells["A" + i].Text,
                            StockCode = firstSheet.Cells["B" + i].Text,
                            Description = firstSheet.Cells["C" + i].Text,

                            Quantity = float.Parse (firstSheet.Cells["D" + i].Text),

                            InvoiceDate =DateTime.Parse (firstSheet.Cells["E" + i].Text) ,

                            UnitPrice = float.Parse (firstSheet.Cells["F" + i].Text),

                            CustomerID = firstSheet.Cells["G" + i].Text,
                            Country = firstSheet.Cells["H" + i].Text,
                    });
                }
                Console.WriteLine("dfdsfdsf=>",list.Count());

            }
        }
    }

    public class FundraiserStudentListModel {
        public string InvoiceNo { get; set; }
        public string StockCode { get; set; }
        public string Description { get; set; }

        public float Quantity { get; set; }

        public DateTime InvoiceDate { get; set; }

        public float UnitPrice { get; set; }

        public string CustomerID { get; set; }
        public string Country { get; set; }

    }
}