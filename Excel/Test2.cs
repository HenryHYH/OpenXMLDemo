using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using System;
using System.Collections.Generic;


namespace Excel
{
    class Test2
    {
        public static string sourceFile = @"..\..\Data\InputTemplate.xlsx";
        public static string targetFile = @"..\..\Data\Test.xlsx";
        public static string ImageFile = @"..\..\Data\Chrysanthemum.jpg";

        static void Main(string[] args)
        {
            File.Copy(sourceFile, targetFile, true);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(@"..\..\Data\Test.xlsx", true))
            {
                WorkbookPart workbookpart = document.WorkbookPart;
                //WorksheetPart sheet1 = workbookpart.WorksheetParts.First();
                WorksheetPart sheet1 = GetSheetByName(workbookpart, "sheet1");

                //insert Image by specifying two range
                InsertImage(sheet1, 1, 1, 3, 3, new FileStream(ImageFile, FileMode.Open));

                document.WorkbookPart.Workbook.Save();
                // Close the document handle.
                document.Close();
            }
            System.Diagnostics.Process.Start(@"..\..\Data\Test.xlsx");
        }

        /// <summary>
        /// Inserts the image at the specified location 
        /// </summary>
        /// <param name="sheet1">The WorksheetPart where image to be inserted</param>
        /// <param name="startRowIndex">The starting Row Index</param>
        /// <param name="startColumnIndex">The starting column index</param>
        /// <param name="endRowIndex">The ending row index</param>
        /// <param name="endColumnIndex">The ending column index</param>
        /// <param name="imageStream">Stream which contains the image data</param>
        private static void InsertImage(WorksheetPart sheet1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, Stream imageStream)
        {
            //Inserting a drawing element in worksheet
            //Make sure that the relationship id is same for drawing element in worksheet and its relationship part
            int drawingPartId = GetNextRelationShipID(sheet1);
            Drawing drawing1 = new Drawing() { Id = "rId" + drawingPartId.ToString() };

            //Check whether the WorksheetPart contains VmlDrawingParts (LegacyDrawing element)
            if (sheet1.VmlDrawingParts == null)
            {
                //if there is no VMLDrawing part (LegacyDrawing element) exists, just append the drawing part to the sheet
                sheet1.Worksheet.Append(drawing1);
            }
            else
            {
                //if VmlDrawingPart (LegacyDrawing element) exists, then find the index of legacy drawing in the sheet and inserts the new drawing element before VMLDrawing part
                int legacyDrawingIndex = GetIndexofLegacyDrawing(sheet1);
                if (legacyDrawingIndex != -1)
                    sheet1.Worksheet.InsertAt<OpenXmlElement>(drawing1, legacyDrawingIndex);
                else
                    sheet1.Worksheet.Append(drawing1);
            }
            //Adding the drawings.xml part
            DrawingsPart drawingsPart1 = sheet1.AddNewPart<DrawingsPart>("rId" + drawingPartId.ToString());
            GenerateDrawingsPart1Content(drawingsPart1, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
            //Adding the image
            ImagePart imagePart1 = drawingsPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
            imagePart1.FeedData(imageStream);
        }

        #region Helper methods

        /// <summary>
        /// Get the index of legacy drawing element in the specified WorksheetPart
        /// </summary>
        /// <param name="sheet1">The worksheetPart</param>
        /// <returns>Index of legacy drawing</returns>
        private static int GetIndexofLegacyDrawing(WorksheetPart sheet1)
        {
            for (int i = 0; i < sheet1.Worksheet.ChildElements.Count; i++)
            {
                OpenXmlElement element = sheet1.Worksheet.ChildElements[i];
                if (element is LegacyDrawing)
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Returns the WorksheetPart for the specified sheet name
        /// </summary>
        /// <param name="workbookpart">The WorkbookPart</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>Returns the WorksheetPart for the specified sheet name</returns>
        private static WorksheetPart GetSheetByName(WorkbookPart workbookpart, string sheetName)
        {
            foreach (WorksheetPart sheetPart in workbookpart.WorksheetParts)
            {
                string uri = sheetPart.Uri.ToString();
                if (uri.EndsWith(sheetName + ".xml"))
                    return sheetPart;
            }
            return null;
        }

        /// <summary>
        /// Returns the next relationship id for the specified WorksheetPart
        /// </summary>
        /// <param name="sheet1">The worksheetPart</param>
        /// <returns>Returns the next relationship id </returns>
        private static int GetNextRelationShipID(WorksheetPart sheet1)
        {
            int nextId = 0;
            List<int> ids = new List<int>();
            foreach (IdPartPair part in sheet1.Parts)
            {
                ids.Add(int.Parse(part.RelationshipId.Replace("rId", string.Empty)));
            }
            if (ids.Count > 0)
                nextId = ids.Max() + 1;
            else
                nextId = 1;
            return nextId;
        }

        // Generates content of drawingsPart1.
        private static void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = startColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "38100";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = startRowIndex.ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = endColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "542925";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = endRowIndex.ToString();
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "161925";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId1", CompressionState = A.BlipCompressionValues.Print };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 1257300L, Y = 762000L };
            A.Extents extents1 = new A.Extents() { Cx = 2943225L, Cy = 2257425L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        #endregion Helper methods
    }
}