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
                        using (OpenXMLHelper.Excel.Writer exc = new OpenXMLHelper.Excel.Writer(fs))
                        {
                            string[][] arrString = new string[][] { 
                                new string[] { "Hello world", "2", "3", "4", "5" }, 
                                new string[] { "", "1", "2", "3", "4" }, 
                                new string[] { "1", "2", "3" }, 
                                new string[] { null, "Hello world" }
                            };

                            double?[][] arrDouble = new double?[][] { new double?[] { 1, 2, 3, 4, 5 }, new double?[] { null, 1, 2, 3.1 } };

                            uint styleIndex = exc.AddStyleSheet(
                                  new OpenXMLHelper.Excel.Writer.ExcelFont() { Size = 20, FontName = "Arial Unicode MS", ColorHex = "FF00FF", IsBold = true, IsItalic = true },
                                  new OpenXMLHelper.Excel.Writer.ExcelBorder() { ColorHex = "000" },
                                  new OpenXMLHelper.Excel.Writer.ExcelFill() { ColorHex = "0FF" },
                                  new OpenXMLHelper.Excel.Writer.ExcelAlign() { Horizontal = OpenXMLHelper.Excel.Writer.ExcelAlign.ExcelAlignHorizontalValue.Center }
                            );

                            exc.WriteData(1, 1, new[] { "Hello world. This is very very long.", "Henry" });
                            //exc.WriteData(1, 1, "Hello world");
                            exc.MergeTwoCells("A1", "B1");

                            exc.WriteData(3, 3, new[] { "Merge A", "Merge B" });
                            exc.MergeTwoCells(3, 3, 3, 4);

                            exc.AddNewWorksheet();

                            exc.WriteData(1, 1, 1000, exc.NumberCellStyleIndex);

                            exc.AddNewWorksheet();
                            exc.InsertImage(@"F:\example\img.jpg", 0, 0, 10, 10);
                            exc.InsertImage(@"F:\example\img.png", 10, 10, 20, 20);
                            exc.RenameCurrentWorksheet("Hello");

                            exc.AddNewWorksheet();
                            exc.WriteData(1, 1, "Hello world", styleIndex);
                            exc.WriteData(2, 1, new string[][] { new string[] { "Hello world" } }, exc.StringCellStyleIndex);
                            exc.WriteData(2, 1, 1000);
                            exc.WriteData(2, 2, arrDouble);
                            exc.InsertImage(@"F:\example\img.jpg", 10, 0);

                            uint styleIndex2 = exc.AddStyleSheet(
                                  new OpenXMLHelper.Excel.Writer.ExcelFont() { Size = 20, FontName = "黑体", ColorHex = "FF00FF", IsBold = true, IsItalic = true },
                                  new OpenXMLHelper.Excel.Writer.ExcelBorder() { ColorHex = "000" },
                                  new OpenXMLHelper.Excel.Writer.ExcelFill() { ColorHex = "0FF" },
                                  new OpenXMLHelper.Excel.Writer.ExcelAlign() { Horizontal = OpenXMLHelper.Excel.Writer.ExcelAlign.ExcelAlignHorizontalValue.Center }
                            );

                            exc.AddNewWorksheet("Hello world");
                            exc.InsertImage(@"F:\example\img.png", 10, 0);
                            exc.WriteData(3, 3, arrString, styleIndex2);

                            exc.AddNewWorksheet("Output");
                            exc.WriteData(1, 1, new string[] { "Name", "Age" }, exc.TableHeaderCellStyleIndex);
                            exc.WriteData(2, 1, new string[][] { new[] { "Henry", "1" }, new[] { "Hello world", "2" } }, exc.TableBodyCellStyleIndex);

                            exc.WriteTable(5, 1, new[] { "Name", "Age" }, new string[][] { new[] { "Henry", "1" }, new[] { "Hello world", "2" } });
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