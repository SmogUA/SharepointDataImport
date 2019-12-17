using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
//using Microsoft.VisualBasic.CompilerServices;


namespace DataImport
{
    public class OpenXML
    {
        public const string CN_DATA_IMPORT_ITERATION_ERRORS = "Data Import Iteration Errors";
        private string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }
        private List<string> preableChars;
        private List<string> GetPreambles()
        {
            if (preableChars == null)
            {
                preableChars = new List<string>();
                preableChars.Add(System.Text.Encoding.Unicode.GetString(System.Text.Encoding.Unicode.GetPreamble()));
                preableChars.Add(System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble()));
                preableChars.Add(System.Text.Encoding.UTF7.GetString(System.Text.Encoding.UTF7.GetPreamble()));
                preableChars.Add(System.Text.Encoding.UTF32.GetString(System.Text.Encoding.UTF32.GetPreamble()));
                preableChars.Add(System.Text.Encoding.ASCII.GetString(System.Text.Encoding.ASCII.GetPreamble()));
                preableChars.Add(System.Text.Encoding.BigEndianUnicode.GetString(System.Text.Encoding.BigEndianUnicode.GetPreamble()));
            }
            return preableChars;
        }
        public List<string> GetSheetsFromFile(System.IO.Stream fileName)
        {
            var sheets = new List<string>();
            using (var document = SpreadsheetDocument.Open(fileName, false))
            {
                sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Select(sh => sh.Name.Value).ToList();
            }
            return sheets;
        }
        private int ConvertColumnNameToNumber(string columnName)
        {
            var alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName))
                throw new ArgumentException();
            var colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);
            int convertedValue = 0;
            for (int i = 0, loopTo = colLetters.Length - 1; i <= loopTo; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? (int)(letter) - 65 : (int)(letter) - 64;
                convertedValue += current * (int)Math.Pow(26, i);
            }
            return convertedValue;
        }
        public DataTable ImportToDataTable(System.IO.Stream fileName, string sheetName, string dateFormat, bool onlyHeaders = false)
        {
            var dt = new DataTable();
            using (var document = SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                var wbPart = document.WorkbookPart;
                // Find the sheet with the supplied name, and then use that Sheet object
                // to retrieve a reference to the appropriate worksheet.
                Sheet theSheet;
                theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
                // Throw an exception if there is no sheet.
                if (theSheet == null)
                    throw new ArgumentException(sheetName);
                // Retrieve a reference to the worksheet part.
                var wsPart = (WorksheetPart)wbPart.GetPartById(theSheet.Id);
                var sstPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable ssTable = null;
                if (sstPart != null)
                    ssTable = sstPart.SharedStringTable;
                // Get the CellFormats for cells without defined data types
                var workbookStylesPart = wbPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
                CellFormats cellFormats = null;
                if (workbookStylesPart != null)
                    cellFormats = workbookStylesPart.Stylesheet.CellFormats;
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                // Dim BOMMarkUTF8 As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
                if (sheetData.Elements<Row>().Count() > 0)
                {
                    if (onlyHeaders)
                    {
                        foreach (Cell cell in sheetData.Descendants<Row>().ElementAt(0))
                        {
                            string pcv = ProcessCellValue(cell, ssTable, cellFormats, dateFormat).ToString();
                            if (dt.Columns.Contains(pcv))
                                return null;
                            else if (!string.IsNullOrEmpty(pcv) && (pcv ?? "") != CN_DATA_IMPORT_ITERATION_ERRORS)
                                dt.Columns.Add(pcv);
                        }
                    }
                    else
                    {
                        foreach (Cell cell in sheetData.Descendants<Row>().ElementAt(0))
                        {
                            string cellRef = GetColumnName(cell.CellReference);
                            // If Not cellRef = errorsColName Then
                            dt.Columns.Add(new DataColumn(cellRef));
                        }
                        foreach (var row in sheetData.Descendants<Row>())
                        {
                            var tempRow = dt.NewRow();
                            int currentCount = 0;
                            foreach (Cell Cell in row.Descendants<Cell>())
                            {
                                string columnName = GetColumnName(Cell.CellReference);
                                // If Not columnName = errorsColName Then
                                int currentColumnIndex = ConvertColumnNameToNumber(columnName);
                                if (currentColumnIndex > dt.Columns.Count - 1)
                                    break;
                                while (currentCount < currentColumnIndex)
                                {
                                    tempRow[currentCount] = string.Empty;
                                    currentCount += 1;
                                }
                                tempRow[currentColumnIndex] = ProcessCellValue(Cell, ssTable, cellFormats, dateFormat);
                                currentCount += 1;
                            }
                            bool isFilled = false;
                            for (int index = 0, loopTo = currentCount - 1; index <= loopTo; index++)
                            {
                                if (!string.IsNullOrEmpty(tempRow[index].ToString()))
                                {
                                    isFilled = true;
                                    break;
                                }
                            }
                            if (!isFilled)
                                continue;
                            while (currentCount < dt.Columns.Count)
                            {
                                tempRow[currentCount] = string.Empty;
                                currentCount += 1;
                            }
                            dt.Rows.Add(tempRow);
                        }
                    }
                }
            }
            return dt;
        }
        public object ProcessCellValue(Cell c, SharedStringTable ssTable, CellFormats cellFormats, string dateFormat)
        {
            // If there is no data type, this must be a string that has been formatted as a number
            string preparedStringValue = string.Empty;
            if (c.DataType == null)
            {
                if (c.CellValue == null)
                    return string.Empty;
                if (c.StyleIndex != null)
                {
                    var cf = cellFormats.Descendants<CellFormat>().ElementAt(Convert.ToInt32(c.StyleIndex.Value));
                    if (cf.NumberFormatId.Value >= (long)0 && cf.NumberFormatId.Value <= (long)13)
                        return Convert.ToDecimal(c.CellValue.Text);
                    else if (cf.NumberFormatId.Value >= (long)14 && cf.NumberFormatId.Value <= (long)22 || cf.NumberFormatId.Value >= (long)45 && cf.NumberFormatId.Value <= (long)47
                                         || cf.NumberFormatId.Value >= (long)165 && cf.NumberFormatId.Value <= (long)180 || cf.NumberFormatId.Value == (long)278 || cf.NumberFormatId.Value == (long)185
                                         || cf.NumberFormatId.Value == (long)196 || cf.NumberFormatId.Value == (long)217 || cf.NumberFormatId.Value == (long)326)
                    {
                        DateTime checkDate;
                        try
                        {
                            preparedStringValue = PrepareStringValue(c.CellValue.Text);
                            checkDate = DateTime.FromOADate(Convert.ToDouble(preparedStringValue));
                        }
                        catch (Exception ex)
                        {
                            return preparedStringValue;
                        }

                        return checkDate.ToString(dateFormat);
                    }
                }

                return PrepareStringValue(c.CellValue.Text);
            }
            switch (c.DataType.Value)
            {
                case CellValues.SharedString:
                    {
                        return PrepareStringValue(ssTable.ChildElements[Convert.ToInt32(c.CellValue.Text)].InnerText);
                    }

                case CellValues.Boolean:
                    {
                        return (PrepareStringValue(c.CellValue.Text) ?? "") == "1" ? true : false;
                    }

                case CellValues.Date:
                    {
                        DateTime checkDate;
                        try
                        {
                            preparedStringValue = PrepareStringValue(c.CellValue.Text);
                            checkDate = DateTime.FromOADate(Convert.ToDouble(preparedStringValue));
                        }
                        catch (Exception ex)
                        {
                            return preparedStringValue;
                        }
                        return checkDate.ToString(dateFormat);
                    }

                case CellValues.Number:
                    {
                        preparedStringValue = PrepareStringValue(c.CellValue.Text);
                        return Convert.ToDecimal(preparedStringValue);
                    }

                case CellValues.InlineString:
                    {
                        return PrepareStringValue(c.InnerText);
                    }

                default:
                    {
                        if (c.CellValue != null)
                            return PrepareStringValue(c.CellValue.Text);
                        return string.Empty;
                    }
            }
        }
        public string PrepareStringValue(string val)
        {
            if (val.Length > 0)
            {
                string pr = GetPreambles().FirstOrDefault(p => Convert.ToString(val[0]) == p);
                if (pr != null)
                    val = val.Replace(pr, string.Empty);
            }
            return Regex.Replace(val, @"\s+", " ");
        }
        private Row CreateContentHeader(uint rowDataIndex, DataColumnCollection dataColumns, DataRow dr)
        {
            var resultRow = new Row() { RowIndex = rowDataIndex };
            for (int iterColIndex = 0, loopTo = dataColumns.Count - 1; iterColIndex <= loopTo; iterColIndex++)
            {
                var cell = CreateHeaderCell(dataColumns[iterColIndex].ColumnName, rowDataIndex, dr[iterColIndex]);
                resultRow.Append(cell);
            }
            return resultRow;
        }
        private Row CreateContentRow(uint rowDataIndex, DataRow dataRow, DataColumnCollection dataColumns)
        {
            var resultRow = new Row() { RowIndex = rowDataIndex };
            for (int iterColIndex = 0, loopTo = dataColumns.Count - 1; iterColIndex <= loopTo; iterColIndex++)
            {
                var cell = CreateContentCell(dataColumns[iterColIndex].ColumnName, rowDataIndex, dataRow[iterColIndex]);
                resultRow.Append(cell);
            }
            return resultRow;
        }
        private Cell CreateContentCell(string header, uint index, object inputValue)
        {
            Cell resultCell;
            var objectType = inputValue.GetType();
            var objectTypeCode = default(TypeCode);
            objectTypeCode = (TypeCode)(int)(Enum.Parse(objectTypeCode.GetType(), objectType.Name));

            switch (objectTypeCode)
            {
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    {
                        resultCell = CreateNumberCell(header, index, inputValue);
                        break;
                    }

                case TypeCode.DateTime:
                    {
                        resultCell = CreateDateCell(header, index, inputValue);
                        break;
                    }

                case TypeCode.Boolean:
                    {
                        resultCell = CreateBooleanCell(header, index, inputValue);
                        break;
                    }

                default:
                    {
                        resultCell = CreateTextCell(header, index, inputValue);
                        break;
                    }
            }

            return resultCell;
        }
        private Cell CreateHeaderCell(string header, uint index, object text)
        {
            var c = new Cell() { DataType = (EnumValue<CellValues>)CellValues.String, CellReference = header + index.ToString() };
            var cellValue = new CellValue() { Text = Convert.ToString(text) };
            c.Append(cellValue);
            return c;
        }
        private Cell CreateTextCell(string header, uint index, object text)
        {
            var c = new Cell() { DataType = (EnumValue<CellValues>)CellValues.InlineString, CellReference = header + index.ToString() };
            var istring = new InlineString();
            var t = new Text() { Text = Convert.ToString(text) };
            istring.Append(t);
            c.Append(istring);
            return c;
        }
        private Cell CreateNumberCell(string header, uint index, object text)
        {
            var c = new Cell() { DataType = (EnumValue<CellValues>)CellValues.Number, CellReference = header + index.ToString() };
            var cellValue = new CellValue() { Text = Convert.ToString(text) };
            c.Append(cellValue);
            return c;
        }
        private Cell CreateDateCell(string header, uint index, object text)
        {
            var c = new Cell() { DataType = (EnumValue<CellValues>)CellValues.Date, CellReference = header + index.ToString() };
            var cellValue = new CellValue() { Text = Convert.ToString(text) };
            c.Append(cellValue);
            return c;
        }
        private Cell CreateBooleanCell(string header, uint index, object text)
        {
            var c = new Cell() { DataType = (EnumValue<CellValues>)CellValues.Boolean, CellReference = header + index.ToString() };
            var cellValue = new CellValue() { Text = Convert.ToString(text) };
            c.Append(cellValue);
            return c;
        }
    }
}
