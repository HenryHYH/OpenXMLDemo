using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Web
{
    public partial class ExcelReader : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string fileName = @"F:\example\B.xlsx";
            using (var reader = new OpenXMLHelper.Excel.Reader(fileName))
            {
                var dt = reader.Read();
            }
        }
    }
}