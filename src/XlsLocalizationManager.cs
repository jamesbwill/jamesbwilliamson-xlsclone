using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using System.IO;
using System.Resources;
using System.Globalization;
using System.Data;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsLocalizationTool
{
    public class XlsLocalizationManager
    {
        object m_objOpt = System.Reflection.Missing.Value;

        private const int _fixedColumns = 4;
        private const int _firsDataRowIndex = 3;
        private const int _maxCultures = 10;

        private enum ExcelFixedColumnNames { FileSource,  FileDest, Key, Value};

        private static string GetFixedCellColumnName(ExcelFixedColumnNames fixedColumns)
        {
            string columnName = String.Empty;

            switch (fixedColumns)
            {
                case ExcelFixedColumnNames.FileSource:
                    columnName = "A";
                    break;
                case ExcelFixedColumnNames.FileDest:
                    columnName = "B";
                    break;
                case ExcelFixedColumnNames.Key:
                    columnName = "C";
                    break;
                case ExcelFixedColumnNames.Value:
                    columnName = "D";
                    break;
                default:

                    break;
            }

            return columnName;
        }

        private Cell FindFixedCell(ExcelFixedColumnNames fixedColumns, List<Cell> cells)
        {
            string columnName = GetFixedCellColumnName(fixedColumns);

            return cells.Where(c => GetColumnName(c.CellReference) == columnName.ToString()).SingleOrDefault();
        }

        private void DataSetToXls(ResourceData rd, string fileName)
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Localize" };
            sheets.Append(sheet);

            Cell c1 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.FileSource), 1, worksheetPart);
            Cell c2 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.FileDest), 1, worksheetPart);
            Cell c3 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.Key), 1, worksheetPart);
            Cell c4 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.Value), 1, worksheetPart);

            c1.Append(new CellValue("Resx source"));
            c2.Append(new CellValue("Resx Name"));
            c3.Append(new CellValue("Key"));
            c4.Append(new CellValue("Value"));

            string[] cultures = GetCulturesFromDataSet(rd);

            int index = _fixedColumns + 1;

            foreach (string cult in cultures)
            {
                CultureInfo ci = new CultureInfo(cult);

                string columnName = GetCultureColumnName(index);

                Cell cell1 = InsertCellInWorksheet(columnName, 1, worksheetPart);
                Cell cell2 = InsertCellInWorksheet(columnName, 2, worksheetPart);

                cell1.Append(new CellValue(ci.DisplayName));
                cell2.Append(new CellValue(ci.Name));

                index++;
            }

            DataView dw = rd.Resource.DefaultView;
            dw.Sort = "FileSource, Key";

            uint row = _firsDataRowIndex;

            foreach (DataRowView drw in dw)
            {
                ResourceData.ResourceRow r = (ResourceData.ResourceRow)drw.Row;

                Cell cell1 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.FileSource), row, worksheetPart);
                Cell cell2 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.FileDest), row, worksheetPart);
                Cell cell3 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.Key), row, worksheetPart);
                Cell cell4 = InsertCellInWorksheet(GetFixedCellColumnName(ExcelFixedColumnNames.Value), row, worksheetPart);

                cell1.Append(new CellValue(r.FileSource));
                cell1.Append(new CellValue(r.FileDestination));
                cell1.Append(new CellValue(r.Key));
                cell1.Append(new CellValue(r.Value));

                ResourceData.ResourceLocalizedRow[] rows = r.GetResourceLocalizedRows();

                foreach (ResourceData.ResourceLocalizedRow lr in rows)
                {
                    string culture = lr.Culture;

                    index = Array.IndexOf(cultures, culture);

                    if (index >= 0)
                    {
                        string columnName = GetCultureColumnName(_fixedColumns + index + 1);
                            
                        Cell cell = InsertCellInWorksheet(columnName, row, worksheetPart);
                        cell.Append(new CellValue(lr.Value));
                    }

                }

                row++;

            }
            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

        }

        private ResourceData XlsToDataSet(string xlsFile)
        {
            ResourceData rd = new ResourceData();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsFile, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                SharedStringTablePart shareStringPart = GetSharedStringTablePart(workbookPart);

                SheetData sheetData =
                worksheetPart.Worksheet.Elements<SheetData>().First();

                Dictionary<string, string> cultures = GetCulturesFromXls(sheetData, shareStringPart);

                IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex >= _firsDataRowIndex);
                
                foreach (Row row in rows)
                {
                    List<Cell> cells = row.Elements<Cell>().ToList();

                    string resourceSrc = GetCellValue(FindFixedCell(ExcelFixedColumnNames.FileSource, cells), shareStringPart);

                    if (!String.IsNullOrEmpty(resourceSrc))
                    {
                        ResourceData.ResourceRow r = rd.Resource.NewResourceRow();

                        r.FileSource = resourceSrc;
                        r.FileDestination = GetCellValue(FindFixedCell(ExcelFixedColumnNames.FileDest, cells), shareStringPart);
                        r.Key = GetCellValue(FindFixedCell(ExcelFixedColumnNames.Key, cells), shareStringPart);
                        r.Value = GetCellValue(FindFixedCell(ExcelFixedColumnNames.Value, cells), shareStringPart);

                        foreach (string culture in cultures.Keys)
                        {
                            string columnName = cultures[culture];

                            Cell cultureCell = cells.Where(c => GetColumnName(c.CellReference) == columnName).SingleOrDefault();

                            if (cultureCell != null)
                            {
                                ResourceData.ResourceLocalizedRow lr = rd.ResourceLocalized.NewResourceLocalizedRow();

                                string localizedValue = GetCellValue(cultureCell, shareStringPart);

                                lr.Culture = culture;
                                lr.Key = r.Key;
                                lr.Value = localizedValue;
                                lr.ParentId = r.Id;

                                lr.SetParentRow(r);

                                rd.ResourceLocalized.AddResourceLocalizedRow(lr);
                            }
                        }

                        rd.Resource.AddResourceRow(r);
                          
                    }
                }

                spreadsheetDocument.Close();
            } 
            
            rd.AcceptChanges();

            return rd;
        }

        private static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellReference)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellReference);

            return uint.Parse(match.Value);
        }

        private static string GetCellValue(Cell cell, SharedStringTablePart stringTablePart)
        {
            if (cell == null) return null;
            if (cell.ChildElements.Count == 0) return null;
            //Get the cell value. 
            string value = cell.CellValue.InnerText;
            //Look up the real value from shared string table. 
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;

            return value;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
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

        private SharedStringTablePart GetSharedStringTablePart(WorkbookPart workbookPart)
        {
            SharedStringTablePart shareStringPart;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

            return shareStringPart;
        }

        private string GetCellCulture(SheetData sheetData, WorksheetPart worksheetPart, SharedStringTablePart sharedStringTablePart, string cellReference)
        {
            string columnName = GetColumnName(cellReference);

            Cell cell = InsertCellInWorksheet(columnName, 2, worksheetPart);

            if (cell == null)
                return String.Empty;
            else
                return GetCellValue(cell, sharedStringTablePart);
        }

        private string GetCultureColumnName(int index)
        {
            string columnName = "E";

            switch (index)
            {
                case 5:
                    columnName = "E";
                    break;
                case 6:
                    columnName = "F";
                    break;
                case 7:
                    columnName = "G";
                    break;
                case 8:
                    columnName = "H";
                    break;
                case 9:
                    columnName = "I";
                    break;
                case 10:
                    columnName = "J";
                    break;
                case 11:
                    columnName = "K";
                    break;
                case 12:
                    columnName = "L";
                    break;
                case 13:
                    columnName = "M";
                    break;
                case 14:
                    columnName = "N";
                    break;
            }

            return columnName;
        }

        private ResourceData ResxToDataSet(string path, bool deepSearch, string[] cultureList, string[] excludeList, bool useFolderNamespacePrefix)
        {
            ResourceData rd = new ResourceData();

            string[] files;

            if (deepSearch)
                files = System.IO.Directory.GetFiles(path, "*.resx", SearchOption.AllDirectories);
            else
                files = System.IO.Directory.GetFiles(path, "*.resx", SearchOption.TopDirectoryOnly);


            foreach (string f in files)
            {
                if (!ResxIsCultureSpecific(f))
                {
                    this.ReadResx(f, path, rd, cultureList, excludeList, useFolderNamespacePrefix);
                }
            }

            return rd;
        }

        public void ResxToXls(string path, bool deepSearch, string xslFile, string[] cultures, string[] excludeList, bool useFolderNamespacePrefix)
        {
            if (!System.IO.Directory.Exists(path))
                return;

            ResourceData rd = ResxToDataSet(path, deepSearch, cultures, excludeList, useFolderNamespacePrefix);

            DataSetToXls(rd, xslFile);

            ShowXls(xslFile);
        }

        public void XlsToResx(string xlsFile)
        {
            XlsToResx(xlsFile, String.Empty);
        }

        public void XlsToResx(string xlsFile, string defaultLang)
        {
            if (!File.Exists(xlsFile))
                return;

            string path = new FileInfo(xlsFile).DirectoryName;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsFile, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                SharedStringTablePart shareStringPart = GetSharedStringTablePart(workbookPart);

                SheetData sheetData =
                worksheetPart.Worksheet.Elements<SheetData>().First();

                Dictionary<string, string> cultures = GetCulturesFromXls(sheetData, shareStringPart);

                Dictionary<string, ResXResourceWriter> cultureWriters = new Dictionary<string,ResXResourceWriter>();
                List<string> generatedfiles = new List<string>();

                IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex >= _firsDataRowIndex);

                foreach (Row row in rows)
                {
                    List<Cell> cells = row.Elements<Cell>().ToList();

                    string resourceSource = GetCellValue(FindFixedCell(ExcelFixedColumnNames.FileSource, cells), shareStringPart);
                    string resourceDest = GetCellValue(FindFixedCell(ExcelFixedColumnNames.FileDest, cells), shareStringPart);
                    string key = GetCellValue(FindFixedCell(ExcelFixedColumnNames.Key, cells), shareStringPart);

                    if ((key is String) && !String.IsNullOrEmpty(key))
                    {
                        foreach (string culture in cultures.Keys)
                        {
                            string columnName = cultures[culture];

                            Cell cultureCell = cells.Where(c => GetColumnName(c.CellReference) == columnName).SingleOrDefault();

                            if (cultureCell != null)
                            {
                                #region create a new writer for the current culture, and if culture file does not exists close the precedent writer and creates a new one
                                
                                string pathCulture = path + @"\" + culture;

                                if (!System.IO.Directory.Exists(pathCulture))
                                    System.IO.Directory.CreateDirectory(pathCulture);

                                string file = pathCulture + @"\" + JustStem(resourceDest) + "." + culture + ".resx";

                                if (culture.Equals(defaultLang, StringComparison.InvariantCultureIgnoreCase))
                                    file = pathCulture + @"\" + JustStem(resourceDest) + ".resx";

                                if (!generatedfiles.Contains(file))
                                {
                                    if (cultureWriters.ContainsKey(culture))
                                    {
                                        if (cultureWriters[culture] is ResXResourceWriter)
                                            cultureWriters[culture].Close();
                                    }
                                    cultureWriters[culture] = new ResXResourceWriter(file);
                                    generatedfiles.Add(file);
                                }
                                #endregion

                                Console.WriteLine(String.Format("[{0}] {1}", culture, key));

                                string localizedValue = GetCellValue(cultureCell, shareStringPart);

                                if (!String.IsNullOrEmpty(localizedValue))
                                {
                                    localizedValue = localizedValue.Replace("\\r", "\r");
                                    localizedValue = localizedValue.Replace("\\n", "\n");

                                    cultureWriters[culture].AddResource(new ResXDataNode(key, localizedValue));
                                }
                            }
                        }
                    }
                }

                spreadsheetDocument.Close();

                foreach (ResXResourceWriter rw in cultureWriters.Values)
                {
                    rw.Close();
                }
            }
        }

        public void XlsToUTF8Properties(string xlsFile, string defaultLang)
        {
            if (!File.Exists(xlsFile))
                return;

            string path = new FileInfo(xlsFile).DirectoryName;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsFile, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                SharedStringTablePart shareStringPart = GetSharedStringTablePart(workbookPart);

                SheetData sheetData =
                worksheetPart.Worksheet.Elements<SheetData>().First();

                Dictionary<string, string> cultures = GetCulturesFromXls(sheetData, shareStringPart);

                Dictionary<string, StreamWriter> cultureWriters = new Dictionary<string, StreamWriter>();
                List<string> generatedfiles = new List<string>();

                IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex >= _firsDataRowIndex);

                foreach (Row row in rows)
                {
                    List<Cell> cells = row.Elements<Cell>().ToList();

                    string resourceSource = GetCellValue(FindFixedCell(ExcelFixedColumnNames.FileSource, cells), shareStringPart);
                    string resourceDest = GetCellValue(FindFixedCell(ExcelFixedColumnNames.FileDest, cells), shareStringPart);
                    string key = GetCellValue(FindFixedCell(ExcelFixedColumnNames.Key, cells), shareStringPart);

                    if ((key is String) && !String.IsNullOrEmpty(key))
                    {
                        foreach (string culture in cultures.Keys)
                        {
                            string columnName = cultures[culture];

                            Cell cultureCell = cells.Where(c => GetColumnName(c.CellReference) == columnName).SingleOrDefault();

                            if (cultureCell != null)
                            {
                                #region create a new writer for the current culture, and if culture file does not exists close the precedent writer and creates a new one
                                
                                string pathCulture = path;

                                if (!System.IO.Directory.Exists(pathCulture))
                                    System.IO.Directory.CreateDirectory(pathCulture);

                                string file = pathCulture + @"\" + JustStem(resourceDest) + "_" + culture + ".properties";

                                if (culture.Equals(defaultLang, StringComparison.InvariantCultureIgnoreCase))
                                    file = pathCulture + @"\" + JustStem(resourceDest) + ".properties";

                                if (!generatedfiles.Contains(file))
                                {
                                    if (cultureWriters.ContainsKey(culture))
                                    {
                                        if (cultureWriters[culture] is StreamWriter)
                                            cultureWriters[culture].Close();
                                    }
                                    cultureWriters[culture] = new StreamWriter(file);
                                    generatedfiles.Add(file);
                                }

                                #endregion

                                Console.WriteLine(String.Format("[{0}] {1}", culture, key));

                                string localizedValue = GetCellValue(cultureCell, shareStringPart);

                                if (!String.IsNullOrEmpty(localizedValue))
                                {
                                    localizedValue = localizedValue.Replace("\\", "&#92;").Replace("'", "&#39;").Replace("\n", "<br/>");

                                    cultureWriters[culture].WriteLine("{0}={1}", key, localizedValue);
                                }
                            }
                        }
                    }
                }

                spreadsheetDocument.Close();

                foreach (TextWriter tw in cultureWriters.Values)
                {
                    tw.Close();
                }
            }
        }

        public void UpdateXls(string xlsFile, string projectRoot, bool deepSearch, string[] excludeList, bool useFolderNamespacePrefix)
        {
            if (!File.Exists(xlsFile))
                return;

            string path = new FileInfo(xlsFile).DirectoryName;

            string[] files;

            if (deepSearch)
                files = System.IO.Directory.GetFiles(projectRoot, "*.resx", SearchOption.AllDirectories);
            else
                files = System.IO.Directory.GetFiles(projectRoot, "*.resx", SearchOption.TopDirectoryOnly);


            ResourceData rd = XlsToDataSet(xlsFile);

            foreach (string f in files)
            {
                FileInfo fi = new FileInfo(f);

                string fileRelativePath = fi.FullName.Remove(0, AddBS(projectRoot).Length);

                string fileDestination;
                if (useFolderNamespacePrefix)
                    fileDestination = GetNamespacePrefix(AddBS(projectRoot), AddBS(fi.DirectoryName)) + fi.Name;
                else
                    fileDestination = fi.Name;

                ResXResourceReader reader = new ResXResourceReader(f);
                reader.BasePath = fi.DirectoryName;

                foreach (DictionaryEntry d in reader)
                {
                    if (d.Value is string)
                    {
                        bool exclude = false;
                        foreach (string e in excludeList)
                        {
                            if (d.Key.ToString().EndsWith(e))
                            {
                                exclude = true;
                                break;
                            }
                        }

                        if (!exclude)
                        {
                            string strWhere = String.Format("FileSource ='{0}' AND Key='{1}'", fileRelativePath, d.Key.ToString());
                            ResourceData.ResourceRow[] rows = (ResourceData.ResourceRow[])rd.Resource.Select(strWhere);

                            ResourceData.ResourceRow row = null;
                            if ((rows == null) | (rows.Length == 0))
                            {
                                // add row
                                row = rd.Resource.NewResourceRow();

                                row.FileSource = fileRelativePath;
                                row.FileDestination = fileDestination;
                                // I update the neutral value
                                row.Key = d.Key.ToString();

                                rd.Resource.AddResourceRow(row);

                            }
                            else
                                row = rows[0];

                            // update row
                            row.BeginEdit();

                            string value = d.Value.ToString();
                            value = value.Replace("\r", "\\r");
                            value = value.Replace("\n", "\\n");
                            row.Value = value;

                            row.EndEdit();
                        }
                    }
                }

            }

            //delete unchenged rows
            foreach (ResourceData.ResourceRow r in rd.Resource.Rows)
            {
                if (r.RowState == DataRowState.Unchanged)
                {
                    r.Delete();
                }
            }
            rd.AcceptChanges();

            DataSetToXls(rd, xlsFile);
        }

        private Dictionary<string, string> GetCulturesFromXls(SheetData sheetData, SharedStringTablePart sharedStringTablePart)
        {
            #region Read Cultures Row

            Dictionary<string, string> cultures = new Dictionary<string, string>();

            Row culturesRow = sheetData.Elements<Row>().Where(r => r.RowIndex == 2).SingleOrDefault();

            List<Cell> cells = culturesRow.Elements<Cell>().ToList();

            foreach (Cell cell in cells)
            {
                string culture = GetCellValue(cell, sharedStringTablePart);

                if (!String.IsNullOrEmpty(culture))
                { 
                    cultures.Add(culture, GetColumnName(cell.CellReference));
                }
            }

            #endregion

            return cultures;
        }

        private string[] GetCulturesFromDataSet(ResourceData rd)
        {
            if (rd.ResourceLocalized.Rows.Count > 0)
            {
                ArrayList list = new ArrayList();
                foreach (ResourceData.ResourceLocalizedRow r in rd.ResourceLocalized.Rows)
                {
                    if (r.Culture != String.Empty)
                    {
                        if (list.IndexOf(r.Culture) < 0)
                        {
                            list.Add(r.Culture);
                        }
                    }
                }

                string[] cultureList = new string[list.Count];

                int i = 0;
                foreach (string c in list)
                {
                    cultureList[i] = c;

                    i++;
                }

                return cultureList;
            }
            else
                return null;
        }

        private bool ResxIsCultureSpecific(string path)
        {
            FileInfo fi = new FileInfo(path);

            //Remove the extension and return the string	
            string fname = JustStem(fi.Name);

            string cult = String.Empty;
            if (fname.IndexOf(".") != -1)
                cult = fname.Substring(fname.LastIndexOf('.') + 1);

            if (cult == String.Empty)
                return false;

            try
            {
                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo(cult);

                return true;
            }
            catch
            {
                return false;
            }
        }

        private string GetNamespacePrefix(string projectRoot, string path)
        {
            path = path.Remove(0, projectRoot.Length);

            if (path.StartsWith(@"\"))
                path = path.Remove(0, 1);

            path = path.Replace(@"\", ".");

            return path;
        }

        private void ReadResx(string fileName, string projectRoot, ResourceData rd, string[] cultureList, string[] excludeList, bool useFolderNamespacePrefix)
        {
            FileInfo fi = new FileInfo(fileName);

            string fileRelativePath = fi.FullName.Remove(0, AddBS(projectRoot).Length);

            string fileDestination;
            if (useFolderNamespacePrefix)
                fileDestination = GetNamespacePrefix(AddBS(projectRoot), AddBS(fi.DirectoryName)) + fi.Name;
            else
                fileDestination = fi.Name;

            ResXResourceReader reader = new ResXResourceReader(fileName);
            reader.BasePath = fi.DirectoryName;

            try
            {
                IDictionaryEnumerator ide = reader.GetEnumerator();

                #region read
                foreach (DictionaryEntry de in reader)
                {
                    if (de.Value is string)
                    {
                        string key = (string)de.Key;

                        bool exclude = false;
                        foreach (string e in excludeList)
                        {
                            if (key.EndsWith(e))
                            {
                                exclude = true;
                                break;
                            }
                        }

                        if (!exclude)
                        {
                            string value = de.Value.ToString();

                            ResourceData.ResourceRow r = rd.Resource.NewResourceRow();

                            r.FileSource = fileRelativePath;
                            r.FileDestination = fileDestination;
                            r.Key = key;

                            value = value.Replace("\r", "\\r");
                            value = value.Replace("\n", "\\n");

                            r.Value = value;

                            rd.Resource.AddResourceRow(r);


                            foreach (string cult in cultureList)
                            {
                                ResourceData.ResourceLocalizedRow lr = rd.ResourceLocalized.NewResourceLocalizedRow();

                                lr.Key = r.Key;
                                lr.Value = String.Empty;
                                lr.Culture = cult;

                                lr.ParentId = r.Id;
                                lr.SetParentRow(r);

                                rd.ResourceLocalized.AddResourceLocalizedRow(lr);
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new Exception("A problem occured reading " + fileName, ex);
            }

            reader.Close();
        }

        public void ShowXls(string xslFilePath)
        {
            if (!System.IO.File.Exists(xslFilePath))
                return;

            System.Diagnostics.Process.Start(xslFilePath);
        }

        public static string AddBS(string cPath)
        {
            if (cPath.Trim().EndsWith("\\"))
            {
                return cPath.Trim();
            }
            else
            {
                return cPath.Trim() + "\\";
            }
        }

        public static string JustStem(string cPath)
        {
            //Get the name of the file
            string lcFileName = JustFName(cPath.Trim());

            //Remove the extension and return the string
            if (lcFileName.IndexOf(".") == -1)
                return lcFileName;
            else
                return lcFileName.Substring(0, lcFileName.LastIndexOf('.'));
        }

        public static string JustFName(string cFileName)
        {
            //Create the FileInfo object
            FileInfo fi = new FileInfo(cFileName);

            //Return the file name
            return fi.Name;
        }
    }
}
