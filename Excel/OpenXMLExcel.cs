using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Data;

namespace Excel
{
    public class OpenXMLExcel : IDisposable
    {
        #region Constrution and dispose

        private readonly System.Type[] numberTypes = new[] { 
            typeof(int), typeof(long), typeof(uint), typeof(ulong), typeof(double), typeof(decimal), typeof(float), 
            typeof(int?), typeof(long?), typeof(uint?), typeof(ulong?), typeof(double?), typeof(decimal?), typeof(float?) 
        };

        SpreadsheetDocument spreadSheet;

        public WorksheetPart CurrentWorksheetPart { get; set; }

        SharedStringTablePart shareStringPart;

        Stylesheet stylesheet;

        public OpenXMLExcel(Stream stream)
        {
            spreadSheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = spreadSheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesheet = workbookPart.WorkbookStylesPart.Stylesheet = new Stylesheet();
            InitStyleSheet();

            WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            shareStringPart.SharedStringTable = new SharedStringTable();
            shareStringPart.SharedStringTable.Count = 1;
            shareStringPart.SharedStringTable.UniqueCount = 1;

            CurrentWorksheetPart = worksheetPart;
        }

        public void Dispose()
        {
            spreadSheet.Close();
            spreadSheet.Dispose();
        }

        private void InitStyleSheet()
        {
            stylesheet.Fonts = new Fonts() { Count = 1 };
            stylesheet.Fonts.Append(new Font(new FontSize() { Val = 11D }, new FontName() { Val = "Arial" }, new FontFamily() { Val = 2 }, new FontScheme() { Val = FontSchemeValues.Minor }, new Color() { Auto = true }));

            // 程序会自动占用 FillId = 0 和 FillId = 1 的 Fill，0 为 无背景，1 为灰色花纹，自定义只能从 2 开始
            stylesheet.Fills = new Fills() { Count = 2 };
            stylesheet.Fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None }));
            stylesheet.Fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }));

            stylesheet.Borders = new Borders() { Count = 1 };
            stylesheet.Borders.Append(new Border(new LeftBorder(), new RightBorder(), new TopBorder() { }, new BottomBorder(), new DiagonalBorder()));

            stylesheet.CellFormats = new CellFormats() { Count = 1 };
            stylesheet.CellFormats.Append(new CellFormat() { FontId = 0, ApplyFont = true, FillId = 0, ApplyFill = true, BorderId = 0, ApplyBorder = true });
        }

        #endregion

        #region Public interface

        public void WriteDataIntoWorkSheet<T>(int rowIndex, int columnIndex, T data, uint? styleIndex = null)
        {
            WriteDataIntoWorkSheet(rowIndex, columnIndex, new T[][] { new T[] { data } }, styleIndex);
        }

        public void WriteDataIntoWorkSheet(int rowIndex, int columnIndex, DataTable dt, uint? styleIndex = null)
        {
            if (rowIndex < 1) rowIndex = 1;
            if (columnIndex < 1) columnIndex = 1;

            WorksheetPart worksheetPart = CurrentWorksheetPart;
            columnIndex -= 1;
            int j = 0;
            foreach (DataRow dr in dt.Rows)
            {
                j++;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    string name = GetColumnName(i + columnIndex);
                    string text = Convert.IsDBNull(dr[i]) ? "" : dr[i].ToString();
                    int index = InsertSharedStringItem(text, shareStringPart);
                    Cell cell = InsertCellInWorksheet(name, Convert.ToUInt32(j + rowIndex), worksheetPart);
                    cell.StyleIndex = styleIndex;

                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    worksheetPart.Worksheet.Save();
                }
            }
        }

        public void WriteDataIntoWorkSheet<T>(int rowIndex, int columnIndex, T[][] data, uint? styleIndex = null)
        {
            if (rowIndex < 1) rowIndex = 1;
            if (columnIndex < 1) columnIndex = 1;

            WorksheetPart worksheetPart = CurrentWorksheetPart;
            columnIndex -= 1;
            int i = 0;
            foreach (T[] row in data)
            {
                int j = 0;
                foreach (T text in row)
                {
                    string name = GetColumnName(j + columnIndex);
                    Cell cell = InsertCellInWorksheet(name, Convert.ToUInt32(i + rowIndex), worksheetPart);
                    cell.StyleIndex = styleIndex;

                    if (numberTypes.Contains(typeof(T)))
                    {
                        cell.CellValue = new CellValue(text.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }
                    else
                    {
                        int index = InsertSharedStringItem(text.ToString(), shareStringPart);
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }

                    worksheetPart.Worksheet.Save();
                    j++;
                }
                i++;
            }
        }

        public void RenameCurrentWorksheet(string sheetName)
        {
            WorkbookPart workbookPart = spreadSheet.WorkbookPart;

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            Sheet sheet = sheets.Elements<Sheet>().Where(x => x.Id == workbookPart.GetIdOfPart(CurrentWorksheetPart)).FirstOrDefault();
            if (null != sheet)
            {
                sheet.Name = sheetName;
            }
        }

        public void AddNewWorksheet(string sheetName = "")
        {
            WorkbookPart workbookPart = spreadSheet.WorkbookPart;

            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();
            CurrentWorksheetPart = newWorksheetPart;
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheetName = "Sheet" + sheetId;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();
        }

        public uint AddStyleSheet(ExcelFont font = null, ExcelBorder border = null, ExcelFill fill = null, ExcelAlign align = null)
        {

            #region Font

            if (null == font)
            {
                font = new ExcelFont();
            }
            stylesheet.Fonts.Append(new Font(
                new FontSize() { Val = font.Size },
                new FontName() { Val = font.FontName },
                new Color() { Rgb = new HexBinaryValue() { Value = font.ColorHex } },
                new FontFamily() { Val = 2 },
                new FontScheme() { Val = FontSchemeValues.Minor },
                new Bold() { Val = font.IsBold },
                new Italic() { Val = font.IsItalic }
            ));
            stylesheet.Fonts.Count += 1;

            #endregion

            #region Border

            if (null == border)
            {
                stylesheet.Borders.Append(new Border()
                {
                    LeftBorder = new LeftBorder(),
                    RightBorder = new RightBorder(),
                    TopBorder = new TopBorder(),
                    BottomBorder = new BottomBorder(),
                    DiagonalBorder = new DiagonalBorder()
                });
            }
            else
            {
                stylesheet.Borders.Append(new Border()
                {
                    LeftBorder = new LeftBorder() { Color = new Color() { Rgb = new HexBinaryValue() { Value = border.ColorHex } }, Style = BorderStyleValues.Thin },
                    RightBorder = new RightBorder() { Color = new Color() { Rgb = new HexBinaryValue() { Value = border.ColorHex } }, Style = BorderStyleValues.Thin },
                    TopBorder = new TopBorder() { Color = new Color() { Rgb = new HexBinaryValue() { Value = border.ColorHex } }, Style = BorderStyleValues.Thin },
                    BottomBorder = new BottomBorder() { Color = new Color() { Rgb = new HexBinaryValue() { Value = border.ColorHex } }, Style = BorderStyleValues.Thin },
                    DiagonalBorder = new DiagonalBorder()
                });
            }
            stylesheet.Borders.Count += 1;

            #endregion

            #region Fill

            if (null == fill)
            {
                stylesheet.Fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None }));
            }
            else
            {
                stylesheet.Fills.Append(new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = fill.ColorHex } },
                        PatternType = PatternValues.Solid
                    }
                });
            }
            stylesheet.Fills.Count += 1;

            #endregion

            #region Align

            if (null == align)
            {
                align = new ExcelAlign();
            }
            Alignment alignment = new Alignment() { Horizontal = (HorizontalAlignmentValues)align.Horizontal, Vertical = (VerticalAlignmentValues)align.Vertical };

            #endregion

            #region CellFormat

            CellFormat cf = new CellFormat();

            cf.Alignment = alignment;
            cf.ApplyAlignment = true;

            cf.FontId = (stylesheet.Fonts.Count ?? 1) - 1;
            cf.ApplyFont = true;

            cf.BorderId = (stylesheet.Borders.Count ?? 1) - 1;
            cf.ApplyBorder = true;

            cf.FillId = (stylesheet.Fills.Count ?? 1) - 1;
            cf.ApplyFill = true;

            stylesheet.CellFormats.Append(cf);
            stylesheet.CellFormats.Count += 1;

            #endregion

            return (stylesheet.CellFormats.Count ?? 1) - 1;
        }

        #endregion

        #region private static OpenXml methods

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
                shareStringPart.SharedStringTable.Count = 1;
                shareStringPart.SharedStringTable.UniqueCount = 1;
            }
            int i = 0;
            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            workbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();
            return newWorksheetPart;
        }

        /// <summary>
        /// Given a Worksheet and a cell name, verifies that the specified cell exists.
        /// If it does not exist, creates a new cell. 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellName"></param>
        private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
            }

        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellName"></param>
        /// <returns></returns>
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the row index.
        /// </summary>
        /// <param name="cellName"></param>
        /// <returns></returns>
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        #endregion

        #region Utility methods

        private static string GetColumnName(int index)
        {
            string name = "";
            char[] columnNames = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            int num = index;
            do
            {
                int i = num % 26;
                name = columnNames[i] + name;
                num = num / 26 - 1;
            } while (num > -1);
            if (string.IsNullOrEmpty(name))
                name = "A";
            return name;
        }

        #endregion

        #region Class

        public class ExcelFont
        {
            public ExcelFont()
            {
                FontName = "Arial";
                ColorHex = "000";
                Size = 11;
                IsBold = false;
                IsItalic = false;
            }

            /// <summary>
            /// e.g Arial
            /// </summary>
            public string FontName { get; set; }

            /// <summary>
            /// e.g FF0000
            /// </summary>
            public string ColorHex { get; set; }

            /// <summary>
            /// e.g 11
            /// </summary>
            public double Size { get; set; }

            public bool IsBold { get; set; }

            public bool IsItalic { get; set; }
        }

        public class ExcelBorder
        {
            /// <summary>
            /// e.g FF0000
            /// </summary>
            public string ColorHex { get; set; }
        }

        public class ExcelFill
        {
            /// <summary>
            /// e.g FF0000
            /// </summary>
            public string ColorHex { get; set; }
        }

        public class ExcelAlign
        {
            public ExcelAlign()
            {
                Horizontal = ExcelAlignHorizontalValue.Justify;
                Vertical = ExcelAlignVerticalValue.Justify;
            }

            public enum ExcelAlignHorizontalValue
            {
                General = 0,
                Left = 1,
                Center = 2,
                Right = 3,
                Fill = 4,
                Justify = 5,
                CenterContinuous = 6,
                Distributed = 7
            }

            public enum ExcelAlignVerticalValue
            {
                Top = 0,
                Center = 1,
                Bottom = 2,
                Justify = 3,
                Distributed = 4
            }

            public ExcelAlignHorizontalValue Horizontal { get; set; }

            public ExcelAlignVerticalValue Vertical { get; set; }
        }

        #endregion
    }
}
