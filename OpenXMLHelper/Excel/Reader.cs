using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OpenXMLHelper.Excel
{
    public class Reader
    {
        public static string Read(string path)
        {
            StringBuilder sb = new StringBuilder();

            using (var document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                //Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "MasterRolePermissionMap");
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }
                //load sheet
                WorksheetPart sheetPart = wbPart.GetPartById(theSheet.Id) as WorksheetPart;

                //load shard string
                var shareStringPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                var shareStrings = new List<string>();
                foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                {
                    shareStrings.Add(item.InnerText);
                }

                SheetData sheetData = sheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string text = "null";
                        var cellValue = cell.CellValue;

                        if (cellValue != null)
                        {
                            var cellType = cell.DataType;
                            if (cellType == null)
                            {
                                text = cellValue.InnerText;
                            }
                            else
                            {
                                switch (cellType.InnerText)
                                {
                                    case "s":
                                        if (cellType.Value == CellValues.SharedString)
                                        {
                                            text = shareStrings[int.Parse(cellValue.InnerText)];
                                        }
                                        else
                                        {
                                            text = cell.CellValue.ToString();
                                        }
                                        break;
                                    default:
                                        text = cellValue.InnerText;
                                        break;
                                }
                            }
                        }
                        sb.AppendFormat("{0} ", text);
                    }
                    sb.AppendFormat("<br />");
                }
            }
            return sb.ToString();
        }
    }
}
