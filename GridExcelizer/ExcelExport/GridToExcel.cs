using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;

namespace GridExcelizer.ExcelExport
{
    public class GridToExcel
    {
        private readonly HttpResponse _response;

        public GridToExcel(HttpResponse response)
        {
            _response = response ?? throw new ArgumentNullException(nameof(response));
        }

        public void ExportToExcel(GridView gridView, string fileName)
        {
            DataTable dt = ConvertGridViewToDataTable(gridView);
            using (MemoryStream memStream = new MemoryStream())
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(memStream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    // Setup styles
                    SetupStyles(workbookPart);

                    // Create worksheet
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                    sheets.Append(sheet);

                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    InsertDataToSheetData(dt, sheetData);

                    // This is to apply filtering on the column headers.
                    // Determine the address of the range to which the filter applies.
                    // This assumes that your data starts in cell A1 and extends
                    // to the last column and row of data in the worksheet.
                    string lastColumnLetter = GetColumnLetter(dt.Columns.Count);
                    uint lastRowNumber = (uint)(dt.Rows.Count + 1);  // Add 1 to account for the header row.
                    string filterRange = $"A1:{lastColumnLetter}{lastRowNumber}";

                    AutoFilter autoFilter = new AutoFilter() { Reference = filterRange };

                    worksheetPart.Worksheet.Append(autoFilter);

                    workbookPart.Workbook.Save();
                }

                _response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                _response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".xlsx");
                _response.BinaryWrite(memStream.ToArray());
                _response.End();
            }
        }

        private DataTable ConvertGridViewToDataTable(GridView gridView)
        {
            DataTable dt = new DataTable();

            // Headers
            foreach (TableCell cell in gridView.HeaderRow.Cells)
            {
                dt.Columns.Add(cell.Text);
            }
            // Rows
            foreach (GridViewRow row in gridView.Rows)
            {
                DataRow newRow = dt.NewRow();
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    // Extract text normally
                    string cellText = HttpUtility.HtmlDecode(row.Cells[i].Text).Trim();

                    // If the cell has controls and the first control is a LinkButton, extract its text
                    if (row.Cells[i].Controls.Count > 0 && row.Cells[i].Controls[0] is LinkButton)
                    {
                        LinkButton linkButton = (LinkButton)row.Cells[i].Controls[0];
                        cellText = HttpUtility.HtmlDecode(linkButton.Text).Trim();
                    }

                    // Convert ($X.XX) to -X.XX
                    if (cellText.StartsWith("($") && cellText.EndsWith(")"))
                    {
                        cellText = "-" + cellText.Trim('(', ')');
                    }

                    newRow[i] = cellText;
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }

        private void SetupStyles(WorkbookPart workbookPart)
        {
            Stylesheet stylesheet = new Stylesheet();

            // Add NumberingFormats to the stylesheet
            NumberingFormats numberingFormats = new NumberingFormats();

            NumberingFormat currencyFormat = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(164),
                FormatCode = StringValue.FromString("\"$\"#,##0.00")
            };
            numberingFormats.Append(currencyFormat);

            NumberingFormat percentFormat = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(165),
                FormatCode = StringValue.FromString("0.00%")
            };

            numberingFormats.Append(percentFormat);

            NumberingFormat dateFormat = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(166),
                FormatCode = StringValue.FromString("MM/DD/YYYY")
            };
            numberingFormats.Append(dateFormat);

            NumberingFormat numberFormat = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(167),
                FormatCode = StringValue.FromString("0")
            };
            numberingFormats.Append(numberFormat);

            stylesheet.Append(numberingFormats);

            // Mandatory elements, even if empty
            Fonts fonts = new Fonts();
            Font defaultFont = new Font();
            fonts.Append(defaultFont);
            stylesheet.Append(fonts);

            Fills fills = new Fills();
            Fill defaultFill = new Fill();
            fills.Append(defaultFill);
            stylesheet.Append(fills);

            Borders borders = new Borders();
            Border defaultBorder = new Border();
            borders.Append(defaultBorder);
            stylesheet.Append(borders);

            CellFormats cellFormats = new CellFormats();

            // Default format
            CellFormat defaultFormat = new CellFormat()
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellFormats.Append(defaultFormat);

            // Currency format
            CellFormat currencyCellStyle = new CellFormat()
            {
                NumberFormatId = 164,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellFormats.Append(currencyCellStyle);

            // Percentage format
            CellFormat percentCellStyle = new CellFormat()
            {
                NumberFormatId = 165,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellFormats.Append(percentCellStyle);

            // Date formet
            CellFormat dateCellStyle = new CellFormat()
            {
                NumberFormatId = 166,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellFormats.Append(dateCellStyle);

            CellFormat numberCellStyle = new CellFormat()
            {
                NumberFormatId = 167,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellFormats.Append(numberCellStyle);

            stylesheet.Append(cellFormats);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = stylesheet;
            workbookStylesPart.Stylesheet.Save();
        }

        private void InsertDataToSheetData(DataTable dt, SheetData sheetData)
        {
            // Initialize an array to keep track of the maximum width of each column.
            double[] maxColumnWidths = new double[dt.Columns.Count];

            // Insert header row to sheetData and update maxColumnWidths
            Row headerExcelRow = new Row();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                DataColumn column = dt.Columns[i];
                Cell headerCell = new Cell();
                headerCell.DataType = CellValues.String;
                headerCell.CellValue = new CellValue(column.ColumnName);
                headerExcelRow.AppendChild(headerCell);
                maxColumnWidths[i] = EstimateColumnWidth(column.ColumnName);
            }
            sheetData.AppendChild(headerExcelRow);

            // Insert data from DataTable to sheetData
            foreach (DataRow dr in dt.Rows)
            {
                Row newRow = new Row();
                for (int i = 0; i < dr.ItemArray.Length; i++)  // Updated loop to use index
                {
                    string value = dr.ItemArray[i].ToString();
                    Cell cell = new Cell();

                    if (IsNumeric(value))
                    {
                        cell.StyleIndex = 1;
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(value);
                        cell.StyleIndex = 4;
                    }
                    else if (IsPercentage(value))
                    {
                        decimal percentageValue = decimal.Parse(value.Replace("%", ""), NumberStyles.Number) / 100m;
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(percentageValue.ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = 2;
                    }
                    else if (IsDate(value))
                    {
                        cell.StyleIndex = 2;
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(DateTime.Parse(value).ToOADate().ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = 3;
                    }
                    else if (IsCurrency(value))
                    {
                        cell.DataType = CellValues.Number;
                        decimal amount = Decimal.Parse(value, NumberStyles.Currency, CultureInfo.CurrentCulture);
                        cell.CellValue = new CellValue(amount.ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = 1;
                    }
                    else
                    {
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(value);
                    }

                    newRow.AppendChild(cell);

                    // Update maxColumnWidths
                    double cellWidth = EstimateColumnWidth(value);
                    maxColumnWidths[i] = Math.Max(maxColumnWidths[i], cellWidth);
                }
                sheetData.AppendChild(newRow);
            }

            // After populating the SheetData, create a Columns element to set the column widths.
            Columns columns = new Columns();
            for (int i = 0; i < maxColumnWidths.Length; i++)
            {
                Column column = new Column()
                {
                    Min = (uint)(i + 1),
                    Max = (uint)(i + 1),
                    Width = maxColumnWidths[i],
                    CustomWidth = true
                };
                columns.Append(column);
            }
            // Insert the Columns element at the beginning of the Worksheet element.
            sheetData.Parent.InsertBefore(columns, sheetData);
        }

        // Method to convert a column number to a column letter (e.g., 1 to A, 2 to B, etc.)
        private string GetColumnLetter(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }

        private bool IsNumeric(string value)
        {
            return Decimal.TryParse(value, out _);
        }

        private bool IsPercentage(string value)
        {
            return value.EndsWith("%") && !value.Any(char.IsLetter);
        }

        private bool IsDate(string value)
        {
            return DateTime.TryParse(value, out _);
        }

        private bool IsCurrency(string value)
        {
            return Decimal.TryParse(value, NumberStyles.Currency, CultureInfo.CurrentCulture, out _);
        }

        //Column width units are not directly equivalent to pixels.
        //So this is used to estimate what the column width might be.
        private double EstimateColumnWidth(string text)
        {
            // Create a Graphics object to measure the text
            using (Graphics graphics = Graphics.FromImage(new Bitmap(1, 1)))
            {
                // Use the font and style that matches the Excel cell
                using (System.Drawing.Font font = new System.Drawing.Font("Calibri", 11))
                {
                    SizeF size = graphics.MeasureString(text, font);

                    // Convert pixels to Excel's column width units.
                    // The factor 7.4 is an estimation used to convert pixel measurements to Excel's column width units.
                    double width = size.Width / 7.4;

                    // Add padding for filter dropdown arrow
                    // This value is also an estimate
                    width += 3.25;

                    return width;
                }
            }
        }
    }
}