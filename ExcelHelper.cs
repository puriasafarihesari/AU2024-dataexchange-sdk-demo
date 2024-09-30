using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AU2024_smart_parameter_updater
{
    static class ExcelHelper
    {
        /// <summary>
        /// Use OpenXml to read data from an excel file
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        static internal DataTable ReadExcelToDataTable(string excelFilePath)
        {
            DataTable table = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                foreach (Cell cell in rows.ElementAt(0))
                {
                    table.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }
                //this will also include your header row...
                foreach (Row row in rows)
                {
                    DataRow tempRow = table.NewRow();
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }
                    table.Rows.Add(tempRow);
                }
            }
            table.Rows.RemoveAt(0);

            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns) { 
                    Console.WriteLine(column.ColumnName + " " + row[column.ColumnName].ToString());
                }
                Console.WriteLine(Environment.NewLine);
            }
            return table;
        }

        /// <summary>
        /// Retrieve the cell value for a given document and cell
        /// </summary>
        /// <param name="document"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }


        /// <summary>
        /// Use OpenXml to read data from an excel file
        /// </summary>
        /// <param name="table"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        public static Dictionary<String, String> GetDataFromExcel(DataTable table, string id)
        {
            var propToDic = new Dictionary<string, string>();
            string idColumnName = table.Columns[0].ColumnName.ToString();
            DataRow matchedRow = null;

            foreach (DataRow row in table.Rows)
            {
                if (row[idColumnName].ToString() == id)
                {
                    matchedRow = row;
                    break;
                }
            }

            if (matchedRow != null)
            {
                for (int i = 1; i < table.Columns.Count; i++)
                {
                    var column = table.Columns[i];
                    propToDic[column.ColumnName] = matchedRow[column].ToString();
                }
            }

            return propToDic;
        }

    }
 }
