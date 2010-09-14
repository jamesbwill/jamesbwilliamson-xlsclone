using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

using System.IO;
using System.Resources;
using System.Globalization;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlsLocalizationTool
{
    public class XlsLocalizationManager
    {
        object m_objOpt = System.Reflection.Missing.Value;

        private void DataSetToXls(ResourceData rd, string fileName)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

            Excel.Sheets sheets = wb.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.get_Item(1);
            sheet.Name = "Localize";

            sheet.Cells[1, 1] = "Resx source";
            sheet.Cells[1, 2] = "Resx Name";
            sheet.Cells[1, 3] = "Key";
            sheet.Cells[1, 4] = "Value";

            string[] cultures = GetCulturesFromDataSet(rd);

            int index = 5;
            foreach (string cult in cultures)
            {
                CultureInfo ci = new CultureInfo(cult);

                sheet.Cells[1, index] = ci.DisplayName;
                sheet.Cells[2, index] = ci.Name;
                index++;
            }

            DataView dw = rd.Resource.DefaultView;
            dw.Sort = "FileSource, Key";

            int row = 3;

            foreach (DataRowView drw in dw)
            {
                ResourceData.ResourceRow r = (ResourceData.ResourceRow)drw.Row;

                sheet.Cells[row, 1] = r.FileSource;
                sheet.Cells[row, 2] = r.FileDestination;
                sheet.Cells[row, 3] = r.Key;
                sheet.Cells[row, 4] = r.Value;

                ResourceData.ResourceLocalizedRow[] rows = r.GetResourceLocalizedRows();

                foreach (ResourceData.ResourceLocalizedRow lr in rows)
                {
                    string culture = lr.Culture;

                    int col = Array.IndexOf(cultures, culture);

                    if (col >= 0)
                        sheet.Cells[row, col + 5] = lr.Value;
                }

                row++;

            }

            sheet.Cells.get_Range("A1", "Z1").EntireColumn.AutoFit();

            // Save the Workbook and quit Excel.
            wb.SaveAs(fileName, m_objOpt, m_objOpt,
                m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
            wb.Close(false, m_objOpt, m_objOpt);

            app.Quit();
            ReleaseObj(app);
            app = null;
        }

        private ResourceData XlsToDataSet(string xlsFile)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xlsFile,
        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            Excel.Sheets sheets = wb.Worksheets;

            Excel.Worksheet sheet = (Excel.Worksheet)sheets.get_Item(1);

            ResourceData rd = new ResourceData();

            int row = 3;

            bool continueLoop = true;
            while (continueLoop)
            {
                string fileSrc = (sheet.Cells[row, 1] as Excel.Range).Text.ToString();

                if (String.IsNullOrEmpty(fileSrc))
                    break;

                ResourceData.ResourceRow r = rd.Resource.NewResourceRow();

                r.FileSource = (sheet.Cells[row, 1] as Excel.Range).Text.ToString();
                r.FileDestination = (sheet.Cells[row, 2] as Excel.Range).Text.ToString();
                r.Key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                r.Value = (sheet.Cells[row, 4] as Excel.Range).Text.ToString();

                rd.Resource.AddResourceRow(r);

                bool hasCulture = true;
                int col = 5;
                while (hasCulture)
                {
                    string cult = (sheet.Cells[2, col] as Excel.Range).Text.ToString();

                    if (String.IsNullOrEmpty(cult))
                        break;

                    ResourceData.ResourceLocalizedRow lr = rd.ResourceLocalized.NewResourceLocalizedRow();

                    lr.Culture = cult;
                    lr.Key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                    lr.Value = (sheet.Cells[row, col] as Excel.Range).Text.ToString();
                    lr.ParentId = r.Id;

                    lr.SetParentRow(r);

                    rd.ResourceLocalized.AddResourceLocalizedRow(lr);

                    col++;
                }

                row++;
            }

            rd.AcceptChanges();

            wb.Close(false, m_objOpt, m_objOpt);

            app.Quit();
            ReleaseObj(app);
            app = null;

            return rd;
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

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xlsFile,
        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            Excel.Sheets sheets = wb.Worksheets;

            Excel.Worksheet sheet = (Excel.Worksheet)sheets.get_Item(1);

            bool hasLanguage = true;
            int col = 5;

            while (hasLanguage)
            {

                object val = (sheet.Cells[2, col] as Excel.Range).Text;

                if (val is string)
                {
                    if (!String.IsNullOrEmpty((string)val))
                    {
                        string cult = (string)val;

                        string pathCulture = path + @"\" + cult;

                        if (!System.IO.Directory.Exists(pathCulture))
                            System.IO.Directory.CreateDirectory(pathCulture);


                        ResXResourceWriter rw = null;

                        int row = 3;

                        string fileSrc;
                        string fileDest;
                        bool readrow = true;

                        while (readrow)
                        {
                            fileSrc = (sheet.Cells[row, 1] as Excel.Range).Text.ToString();
                            fileDest = (sheet.Cells[row, 2] as Excel.Range).Text.ToString();

                            if (String.IsNullOrEmpty(fileDest))
                                break;

                            string f = pathCulture + @"\" + JustStem(fileDest) + "." + cult + ".resx";

                            if (cult.Equals(defaultLang, StringComparison.InvariantCultureIgnoreCase))
                                f = pathCulture + @"\" + JustStem(fileDest) + ".resx";




                            rw = new ResXResourceWriter(f);

                            while (readrow)
                            {

                                string key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                                object data = (sheet.Cells[row, col] as Excel.Range).Text.ToString();

                                Console.WriteLine(String.Format("[{0}] {1}", cult, key));

                                if ((key is String) && !String.IsNullOrEmpty(key))
                                {
                                    string text = data as string;

                                    if (!String.IsNullOrEmpty(text))
                                    {
                                        text = text.Replace("\\r", "\r");
                                        text = text.Replace("\\n", "\n");

                                        rw.AddResource(new ResXDataNode(key, text));
                                    }

                                    row++;

                                    string file = (sheet.Cells[row, 2] as Excel.Range).Text.ToString();

                                    if (file != fileDest)
                                        break;
                                }
                                else
                                {
                                    readrow = false;
                                }
                            }

                            rw.Close();

                        }
                    }
                    else
                        hasLanguage = false;
                }
                else
                    hasLanguage = false;

                col++;
            }

            app.Quit();
            ReleaseObj(app);
            app = null;

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

        public void XlsToUTF8Properties(string xlsFile, string defaultLang)
        {
            if (!File.Exists(xlsFile))
                return;

            string path = new FileInfo(xlsFile).DirectoryName;

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xlsFile,
        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            Excel.Sheets sheets = wb.Worksheets;

            Excel.Worksheet sheet = (Excel.Worksheet)sheets.get_Item(1);

            bool hasLanguage = true;
            int col = 5;

            while (hasLanguage)
            {

                object val = (sheet.Cells[2, col] as Excel.Range).Text;

                if (val is string)
                {
                    if (!String.IsNullOrEmpty((string)val))
                    {
                        string cult = (string)val;

                        string pathCulture = path;

                        if (!System.IO.Directory.Exists(pathCulture))
                            System.IO.Directory.CreateDirectory(pathCulture);


                        TextWriter tw = null;
                        //ResXResourceWriter rw = null;

                        int row = 3;

                        string fileSrc;
                        string fileDest;
                        bool readrow = true;

                        while (readrow)
                        {
                            fileSrc = (sheet.Cells[row, 1] as Excel.Range).Text.ToString();
                            fileDest = (sheet.Cells[row, 2] as Excel.Range).Text.ToString();

                            if (String.IsNullOrEmpty(fileDest))
                                break;

                            string f = pathCulture + @"\" + JustStem(fileDest) + "_" + cult + ".properties";

                            if (cult.Equals(defaultLang, StringComparison.InvariantCultureIgnoreCase))
                                f = pathCulture + @"\" + JustStem(fileDest) + ".properties";

                            tw = new StreamWriter(f);

                            while (readrow)
                            {

                                string key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                                object data = (sheet.Cells[row, col] as Excel.Range).Text.ToString();

                                Console.WriteLine(String.Format("[{0}] {1}", cult, key));

                                if ((key is String) && !String.IsNullOrEmpty(key))
                                {
                                    string text = data as string;

                                    if (!String.IsNullOrEmpty(text))
                                    {
                                        text = text.Replace("\\", "&#92;");
                                        text = text.Replace("'", "&#39;");
                                        text = text.Replace("\n", "<br/>");
                                        //text = HttpUtility.HtmlEncode(text);

                                        //rw.AddResource(new ResXDataNode(key, text));
                                        tw.WriteLine(key + "=" + text);
                                    }

                                    row++;

                                    string file = (sheet.Cells[row, 2] as Excel.Range).Text.ToString();

                                    if (file != fileDest)
                                        break;
                                }
                                else
                                {
                                    readrow = false;
                                }
                            }

                            //rw.Close();
                            tw.Close();
                        }
                    }
                    else
                        hasLanguage = false;
                }
                else
                    hasLanguage = false;

                col++;
            }

            app.Quit();
            ReleaseObj(app);
            app = null;

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

        /// <summary>
        /// Ensures that the Interop Application will be released correctly.
        /// </summary>
        /// <param name="obj">L'objet COM à tuer.</param>
        private void ReleaseObj(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void ShowXls(string xslFilePath)
        {
            if (!System.IO.File.Exists(xslFilePath))
                return;

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xslFilePath,
        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            app.Visible = true;
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
