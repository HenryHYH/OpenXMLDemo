using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel;

namespace Web
{
    public partial class Excel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    using (System.IO.FileStream fs = new System.IO.FileStream(@"f:\example\a.xlsx", System.IO.FileMode.Create))
                    {
                        using (OpenXMLExcel exc = new OpenXMLExcel(fs))
                        {
                            string[][] arr = new string[][] { new string[] { "1", "2", "3", "4", "5" }, new string[] { "", "1", "2", "3", "4" }, new string[] { "1", "2", "3" } };
                            exc.WriteDataIntoWorkSheet(1, 1, arr);

                            exc.RenameCurrentWorksheet("Hello");

                            //exc.AddNewWorksheet("Hello world");

                            //exc.WriteDataIntoWorkSheet(2, 2, arr);
                        }
                    }
                    Response.Write("Success");
                }
                catch (Exception ex)
                {
                    Response.Write(ex.ToString());
                }
            }
        }
    }
}