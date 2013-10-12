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
                        using (OpenXMLHelper.OpenXMLExcel exc = new OpenXMLHelper.OpenXMLExcel(fs))
                        {
                            string[][] arrString = new string[][] { 
                                new string[] { "Hello world", "2", "3", "4", "5" }, 
                                new string[] { "", "1", "2", "3", "4" }, 
                                new string[] { "1", "2", "3" }, 
                                new string[] { "Hello world" }
                            };

                            double?[][] arrDouble = new double?[][] { new double?[] { 1, 2, 3, 4, 5 }, new double?[] { null, 1, 2, 3.1 } };

                            uint styleIndex = exc.AddStyleSheet(
                                  new OpenXMLHelper.OpenXMLExcel.ExcelFont() { ColorHex = "FF00FF", IsBold = true, IsItalic = true },
                                  new OpenXMLHelper.OpenXMLExcel.ExcelBorder() { ColorHex = "000" },
                                  new OpenXMLHelper.OpenXMLExcel.ExcelFill() { ColorHex = "0FF" },
                                  new OpenXMLHelper.OpenXMLExcel.ExcelAlign() { Horizontal = OpenXMLHelper.OpenXMLExcel.ExcelAlign.ExcelAlignHorizontalValue.Center }
                            );

                            exc.InsertImage(@"F:\example\img.jpg", 0, 0, 10, 10);
                            exc.InsertImage(@"F:\example\img.png", 10, 10, 20, 20);
                            //exc.InsertImage(0, 70, @"F:\example\img.jpg");
                            //exc.InsertImage(0, 1000, @"F:\example\img.png");
                            exc.RenameCurrentWorksheet("Hello");

                            exc.AddNewWorksheet();
                            exc.WriteData(1, 1, "Hello world", styleIndex);
                            exc.WriteData(2, 1, new string[][] { new string[] { "Hello world" } });
                            exc.WriteData(2, 1, 1000);
                            exc.WriteData(2, 2, arrDouble);
                            exc.InsertImage(@"F:\example\img.jpg", 10, 0);

                            exc.AddNewWorksheet("Hello world");
                            //exc.InsertImage(10, 10, 20, 20, @"F:\example\img.jpg");
                            exc.InsertImage(@"F:\example\img.png", 10, 0);
                            exc.WriteData(3, 3, arrString, styleIndex);
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