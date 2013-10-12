using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Web
{
    public partial class ExcelPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    using (System.IO.FileStream fs = new System.IO.FileStream(string.Format(@"{0}\a.xlsx", @"f:\example"), System.IO.FileMode.Create))
                    {
                        using (Excel.OpenXMLExcel exc = new Excel.OpenXMLExcel(fs))
                        {
                            string[][] arrString = new string[][] { 
                                new string[] { "Hello world", "2", "3", "4", "5" }, 
                                new string[] { "", "1", "2", "3", "4" }, 
                                new string[] { "1", "2", "3" }, 
                                new string[] { "Hello world" }
                            };

                            double?[][] arrDouble = new double?[][] { new double?[] { 1, 2, 3, 4, 5 }, new double?[] { null, 1, 2, 3.1 } };

                            uint styleIndex = exc.AddStyleSheet(
                                  new Excel.OpenXMLExcel.ExcelFont() { ColorHex = "FF00FF", IsBold = true, IsItalic = true },
                                  new Excel.OpenXMLExcel.ExcelBorder() { ColorHex = "000" },
                                  new Excel.OpenXMLExcel.ExcelFill() { ColorHex = "0FF" },
                                  new Excel.OpenXMLExcel.ExcelAlign() { Horizontal = Excel.OpenXMLExcel.ExcelAlign.ExcelAlignHorizontalValue.Center }
                            );

                            exc.WriteDataIntoWorkSheet(1, 1, "Hello world", styleIndex);
                            exc.WriteDataIntoWorkSheet(2, 1, new string[][] { new string[] { "Hello world" } });
                            exc.WriteDataIntoWorkSheet(2, 1, 1000);
                            exc.WriteDataIntoWorkSheet(2, 2, arrDouble);
                            exc.RenameCurrentWorksheet("Hello");
                            exc.AddNewWorksheet("Hello world");
                            exc.WriteDataIntoWorkSheet(3, 3, arrString, styleIndex);
                        }
                    }
                    WriteMessage("Success");
                }
                catch (Exception ex)
                {
                    WriteMessage(ex.ToString());
                }
            }
        }

        private void WriteMessage(string message)
        {
            //Response.Write(message);
            lbMessage.Text = message;
        }
    }
}