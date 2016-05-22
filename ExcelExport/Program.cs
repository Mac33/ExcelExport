using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraSpreadsheet;
using System;
using System.Data;
using System.Linq;



namespace ExcelExport
{
    internal static class Program
    {
        private const int DateColumnIndex = 0;
        private const int ItemNameColumnIndex = 1;

        private const string DateColumnName = "Date";
        private const string ItemNameColumnName = "Name";


        private static string[] GetPriceColumNames()
        {
            var columNames = new string[] { "Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7","Col8", "Col9" };
            return columNames;
        }



        [STAThread]
        private static void Main()
        {
            var mainTable = CreateDataTableFromXLS("D:\\fin.xlsx");
            mainTable.WriteXml("D:\\file.xml");
            SaveDataToExcel(mainTable, "D:\\new.xlsx", "D:\\template.xlsx");
            Console.WriteLine("Done");
            Console.ReadKey();
        }



        private static DataTable CreateDataTableFromXLS(string fileName)
        {
            using (DataTable mainTable = CreateMainTableStructure())
            {
                using (DataSet dataset = LoadDataFromExcel(fileName))
                {
                    foreach (DataTable table in dataset.Tables)
                    {
                        DateTime date = DateTime.MinValue;
                        int emptyRows = 0;

                        Console.WriteLine("Exporting DatatTable: {0}", table.TableName);

                        foreach (DataRow row in table.Rows)
                        {
                            if (emptyRows > 5)
                            {
                                break;
                            }

                            if (string.IsNullOrEmpty(row[DateColumnIndex].ToString()))
                            {
                                if (date != DateTime.MinValue)
                                {
                                    row[DateColumnIndex] = date;
                                }
                            }
                            else
                            {
                                date = TryToSetCorrectDate(date, table.TableName, row);
                            }

                            if (!string.IsNullOrEmpty(row[ItemNameColumnIndex].ToString()))
                            {
                                AddNewRowToMainDataTable(mainTable, row);
                            }
                            else
                            {
                                emptyRows++;
                            }
                        }
                        
                    }
                }
               //GC.Collect();
                return mainTable;
            }
        }



        private static DateTime  TryToSetCorrectDate(DateTime checkDate, string tableName, DataRow row)
        {
            DateTime.TryParse(row[DateColumnIndex].ToString(), out checkDate);
            var year = System.Convert.ToInt32(tableName.Substring(3, 4));
            var month = System.Convert.ToInt32(tableName.Substring(0, 2));
            if (checkDate.Year != year && month < 12)
            {
                // Check year mistake 
                checkDate = new DateTime(year, checkDate.Month, checkDate.Day);
            }
            return checkDate;
        }



        private static DataTable CreateMainTableStructure()
        {
            var table = new DataTable("data");
            table.Columns.Add(DateColumnName, typeof(DateTime));
            table.Columns.Add(ItemNameColumnName, typeof(string));
            table.Columns.Add("...", typeof(string));

            foreach (string colName in GetPriceColumNames())
            {
                table.Columns.Add(colName, typeof(double));
            }
            return table;
        }



        private static DataSet LoadDataFromExcel(string fileName)
        {
            using (SpreadsheetControl control = new SpreadsheetControl())
            {
                Console.WriteLine("Loading Excel document...");
                control.Document.LoadDocument(fileName);

                DataSet dataset = new DataSet();

                foreach (Worksheet sheet in control.Document.Worksheets)
                {
                    DateTime date;
                    if (DateTime.TryParse(sheet.Name, out date))
                    {
                        string tableName = GetTableName(date, "/");
                        Console.WriteLine("Loading DatatTable: {0}", tableName);

                        DataTable table  = sheet.CreateDataTable(sheet.GetDataRange(), true);
                        table.TableName = tableName;
                        dataset.Tables.Add(table);

                        
                        var exporter = sheet.CreateDataTableExporter(sheet.GetDataRange(), table, true);
                        exporter.CellValueConversionError += exporter_CellValueConversionError;
                        exporter.Export();
                        exporter = null;

                        CorrectDecimalStrings(table);

                    }
                }

                return dataset;
            }
        }



        private static void CorrectDecimalStrings(DataTable table)
        {
            foreach (System.Data.DataColumn col in table.Columns)
            {
                if (GetPriceColumNames().Contains(col.ColumnName) && col.DataType != typeof(double))
                {
                    foreach (DataRow row in table.Rows)
                    {
                        row[col] = row[col].ToString().Replace(",", ".");
                    }
                }
            }
        }




        private static void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            Console.WriteLine("Error in cell " + e.Cell.GetReferenceA1());
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
        }





        private static void AddNewRowToMainDataTable(DataTable mainTable, DataRow row)
        {
            var newRow = mainTable.NewRow();

            newRow[DateColumnIndex] = row[DateColumnIndex];
            newRow[ItemNameColumnIndex] = row[ItemNameColumnIndex];

            CopyDecimalValues(row, newRow);

            mainTable.Rows.Add(newRow);
        }



        private static void CopyDecimalValues(DataRow row, DataRow newRow)
        {
            var valueWasSet = false;
            foreach (System.Data.DataColumn col in row.Table.Columns)
            {
                if (GetPriceColumNames().Contains(col.ColumnName.Trim()))
                {
                    if (!DBNull.Value.Equals(row[col.ColumnName]))
                    {
                        decimal decimalValue = 0;
                        if (decimal.TryParse(row[col.ColumnName].ToString(), out decimalValue))
                        {
                            newRow[col.ColumnName.Trim()] = decimalValue;
                            valueWasSet = true;
                        }
                        else
                        {
                            Console.WriteLine("Invalid conversion");
                        }
                    }
                }
            }
            if (!valueWasSet)
            {
                throw new Exception();
            }
        }
        


        private static void SaveDataToExcel(DataTable table, string saveToFile, string templateFileName)
        {
            InitIdColumn(table);

            DataView view = new DataView(table) { Sort = String.Format("{0} ASC, ID ASC", DateColumnName) };

            using (SpreadsheetControl control = new SpreadsheetControl())
            {
                control.Document.LoadDocument(templateFileName);
                var template = control.Document.Worksheets.ActiveWorksheet;

                var currentMonth = 0;
                var itemPosition = 1;

                foreach (DataRowView row in view)
                {
                    var itemDate = Convert.ToDateTime(row[DateColumnName]);
                    if (itemDate.Month != currentMonth)
                    {
                        currentMonth = itemDate.Month;
                        CreateNewSheet(control, template, itemDate);
                        itemPosition = 1;
                    }

                    SetValuesToActiveWorksheet(control, itemPosition, itemDate, row);


                    itemPosition++;
                }

                control.Document.SaveDocument(saveToFile);
            }
        }



        private static void InitIdColumn(DataTable table)
        {
            table.Columns.Add("ID", typeof(int));
            var id = 0;
            foreach (DataRow row in table.Rows)
            {
                row["ID"] = id;
                id++;
            }
        }



        private static void CreateNewSheet(SpreadsheetControl control, Worksheet template, DateTime date)
        {
            Console.WriteLine("Creating new excel sheet: {0}", control.Document.Worksheets.ActiveWorksheet.Name);
            control.Document.Worksheets.Add(GetTableName(date, "-"));
            control.Document.Worksheets.ActiveWorksheet.CopyFrom(template);
        }



        private static void SetValuesToActiveWorksheet(SpreadsheetControl control, int itemPosition, DateTime itemDate, DataRowView sourceRow)
        {
            control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 0].Value = itemDate;
            control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 1].Value = sourceRow[1].ToString();
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 3], (sourceRow[GetPriceColumNames()[0]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 4], (sourceRow[GetPriceColumNames()[1]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 5], (sourceRow[GetPriceColumNames()[2]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 6], (sourceRow[GetPriceColumNames()[3]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 7], (sourceRow[GetPriceColumNames()[4]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 8], (sourceRow[GetPriceColumNames()[5]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 9], (sourceRow[GetPriceColumNames()[6]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 10], (sourceRow[GetPriceColumNames()[7]]));
            SetDecimalValue(control.Document.Worksheets.ActiveWorksheet.Cells[itemPosition, 11], (sourceRow[GetPriceColumNames()[8]]));
        }



        private static void SetDecimalValue(Cell cell, object value)
        {
            if (!DBNull.Value.Equals(value))
            {
                cell.Value = Convert.ToDouble(value);
            }
        }



        private static string GetTableName(DateTime date, string separator)
        {
            if (date.Month < 10)
            {
                return string.Format("0{0}{1}{2}", date.Month, separator, date.Year);
            }
            else
            {
                return string.Format("{0}{1}{2}", date.Month, separator, date.Year);
            }
        }
    }
}
