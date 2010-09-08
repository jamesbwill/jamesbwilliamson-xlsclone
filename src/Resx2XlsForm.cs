using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Globalization;
using System.Collections;
using System.Resources;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace Resx2Xls
{
    public partial class Resx2XlsForm : Form
    {
        object m_objOpt = System.Reflection.Missing.Value;

        enum ResxToXlsOperation { Create, Build, Update };

        private ResxToXlsOperation _operation;

        string _summary1;
        string _summary2;
        string _summary3;

        public Resx2XlsForm()
        {
            //CultureInfo ci = new CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            //System.Threading.Thread.CurrentThread.CurrentUICulture = ci;

            InitializeComponent();

            this.textBoxFolder.Text = Properties.Settings.Default.FolderPath;
            this.textBoxExclude.Text = Properties.Settings.Default.ExcludeList;
            this.checkBoxFolderNaming.Checked = Properties.Settings.Default.FolderNamespaceNaming;

            FillCultures();

            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonUpdateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);

            _summary1 = "Operation:\r\nCreate a new Excel document ready for localization.";
            _summary2 = "Operation:\r\nBuild your localized Resource files from a Filled Excel Document.";
            _summary3 = "Operation:\r\nUpdate your Excel document with your Project Resource changes.";

            this.textBoxSummary.Text = _summary1;
        }

        void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            this.radioButtonCreateXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonUpdateXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);

            if (this.radioButtonCreateXls.Checked)
            {
                _operation = ResxToXlsOperation.Create;
                this.textBoxSummary.Text = _summary1;
            }
            if (this.radioButtonBuildXls.Checked)
            {
                _operation = ResxToXlsOperation.Build;
                this.textBoxSummary.Text = _summary2;
            }
            if (this.radioButtonUpdateXls.Checked)
            {
                _operation = ResxToXlsOperation.Update;
                this.textBoxSummary.Text = _summary3;
            }

            if (((RadioButton)sender).Checked)
            {
                if (((RadioButton)sender) == this.radioButtonCreateXls)
                {
                    this.radioButtonBuildXls.Checked = false;
                    this.radioButtonUpdateXls.Checked = false;
                }

                if (((RadioButton)sender) == this.radioButtonBuildXls)
                {
                    this.radioButtonCreateXls.Checked = false;
                    this.radioButtonUpdateXls.Checked = false;
                }

                if (((RadioButton)sender) == this.radioButtonUpdateXls)
                {
                    this.radioButtonCreateXls.Checked = false;
                    this.radioButtonBuildXls.Checked = false;
                }
            }
            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonUpdateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
        }

        public void ResxToXls(string path, bool deepSearch, string xslFile, string[] cultures, string[] excludeList, bool useFolderNamespacePrefix)
        {
            if (!System.IO.Directory.Exists(path))
                return;

            ResxData rd = ResxToDataSet(path, deepSearch, cultures, excludeList, useFolderNamespacePrefix);

            DataSetToXls(rd, xslFile);

            ShowXls(xslFile);
        }

        private void XlsToResx(string xlsFile)
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

                            rw = new ResXResourceWriter(f);

                            while (readrow)
                            {
                                string key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                                object data = (sheet.Cells[row, col] as Excel.Range).Text.ToString();

                                if ((key is String) & !String.IsNullOrEmpty(key))
                                {
                                    if (data is string)
                                    {
                                        string text = data as string;

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
        }

        private ResxData ResxToDataSet(string path, bool deepSearch, string[] cultureList, string[] excludeList, bool useFolderNamespacePrefix)
        {
            ResxData rd = new ResxData();

            string[] files;

            if (deepSearch)
                files = System.IO.Directory.GetFiles(path, "*.resx", SearchOption.AllDirectories);
            else
                files = System.IO.Directory.GetFiles(path, "*.resx", SearchOption.TopDirectoryOnly);


            foreach (string f in files)
            {
                if (!ResxIsCultureSpecific(f))
                {
                    ReadResx(f, path, rd, cultureList, excludeList, useFolderNamespacePrefix);
                }
            }

            return rd;
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

        private void ReadResx(string fileName, string projectRoot, ResxData rd, string[] cultureList, string[] excludeList, bool useFolderNamespacePrefix)
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

                            ResxData.ResxRow r = rd.Resx.NewResxRow();

                            r.FileSource = fileRelativePath;
                            r.FileDestination = fileDestination;
                            r.Key = key;

                            value = value.Replace("\r", "\\r");
                            value = value.Replace("\n", "\\n");

                            r.Value = value;

                            rd.Resx.AddResxRow(r);


                            foreach (string cult in cultureList)
                            {
                                ResxData.ResxLocalizedRow lr = rd.ResxLocalized.NewResxLocalizedRow();

                                lr.Key = r.Key;
                                lr.Value = String.Empty;
                                lr.Culture = cult;

                                lr.ParentId = r.Id;
                                lr.SetParentRow(r);

                                rd.ResxLocalized.AddResxLocalizedRow(lr);
                            }
                        }
                    }
                }
                #endregion
            }
            catch(Exception ex)
            {
                MessageBox.Show("A problem occured reading " + fileName + "\n" + ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            reader.Close();
        }

        private void FillCultures()
        {
            CultureInfo[] array = CultureInfo.GetCultures(CultureTypes.AllCultures);
            Array.Sort(array, new CultureInfoComparer());
            foreach (CultureInfo info in array)
            {
                if (info.Equals(CultureInfo.InvariantCulture))
                {
                    //this.listBoxCultures.Items.Add(info, "Default (Invariant Language)");
                }
                else
                {
                    this.listBoxCultures.Items.Add(info);
                }

            }

            string cList = Properties.Settings.Default.CultureList;

            string[] cultureList = cList.Split(';');

            foreach (string cult in cultureList)
            {
                CultureInfo info = new CultureInfo(cult);

                this.listBoxSelected.Items.Add(info);
            }
        }

        private void AddCultures()
        {
            for (int i = 0; i < this.listBoxCultures.SelectedItems.Count; i++)
            {
                CultureInfo ci = (CultureInfo)this.listBoxCultures.SelectedItems[i];

                if (this.listBoxSelected.Items.IndexOf(ci) == -1)
                    this.listBoxSelected.Items.Add(ci);
            }
        }

        private void SaveCultures()
        {
            string cultures = String.Empty;
            for (int i = 0; i < this.listBoxSelected.Items.Count; i++)
            {
                CultureInfo info = (CultureInfo)this.listBoxSelected.Items[i];

                if (cultures != String.Empty)
                    cultures = cultures + ";";

                cultures = cultures + info.Name;
            }

            Properties.Settings.Default.CultureList = cultures;
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFolder.Text = this.folderBrowserDialog.SelectedPath;
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            AddCultures();
        }

        private void listBoxCultures_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            AddCultures();
        }

        private void buttonBrowseXls_Click(object sender, EventArgs e)
        {
            if (this.openFileDialogXls.ShowDialog() == DialogResult.OK)
            {
                this.textBoxXls.Text = this.openFileDialogXls.FileName;
            }
        }


        private void listBoxSelected_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.listBoxSelected.SelectedItems.Count > 0)
            {
                this.listBoxSelected.Items.Remove(this.listBoxSelected.SelectedItems[0]);
            }
        }

        private void textBoxExclude_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ExcludeList = this.textBoxExclude.Text;

        }

        private void Resx2XlsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveCultures();

            Properties.Settings.Default.FolderNamespaceNaming = this.checkBoxFolderNaming.Checked;

            Properties.Settings.Default.Save();
        }

        private void textBoxFolder_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.FolderPath = this.textBoxFolder.Text;
        }

        private void UpdateXls(string xlsFile, string projectRoot, bool deepSearch, string[] excludeList, bool useFolderNamespacePrefix)
        {
            if (!File.Exists(xlsFile))
                return;

            string path = new FileInfo(xlsFile).DirectoryName;

            string[] files;

            if (deepSearch)
                files = System.IO.Directory.GetFiles(projectRoot, "*.resx", SearchOption.AllDirectories);
            else
                files = System.IO.Directory.GetFiles(projectRoot, "*.resx", SearchOption.TopDirectoryOnly);


            ResxData rd = XlsToDataSet(xlsFile);
            
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
                            ResxData.ResxRow[] rows = (ResxData.ResxRow[])rd.Resx.Select(strWhere);

                            ResxData.ResxRow row = null;
                            if ((rows == null) | (rows.Length == 0))
                            {
                                // add row
                                row = rd.Resx.NewResxRow();

                                row.FileSource = fileRelativePath;
                                row.FileDestination = fileDestination;
                                // I update the neutral value
                                row.Key = d.Key.ToString();

                                rd.Resx.AddResxRow(row);
                                
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
            foreach (ResxData.ResxRow r in rd.Resx.Rows)
            {
                if (r.RowState == DataRowState.Unchanged)
                {
                    r.Delete();
                }
            }
            rd.AcceptChanges();

            DataSetToXls(rd, xlsFile);
        }

        private ResxData XlsToDataSet(string xlsFile)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xlsFile,
        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            Excel.Sheets sheets = wb.Worksheets;

            Excel.Worksheet sheet = (Excel.Worksheet)sheets.get_Item(1);

            ResxData rd = new ResxData();

            int row = 3;

            bool continueLoop = true;
            while (continueLoop)
            {
                string fileSrc = (sheet.Cells[row, 1] as Excel.Range).Text.ToString();

                if (String.IsNullOrEmpty(fileSrc))
                    break;

                ResxData.ResxRow r = rd.Resx.NewResxRow();

                r.FileSource = (sheet.Cells[row, 1] as Excel.Range).Text.ToString();
                r.FileDestination = (sheet.Cells[row, 2] as Excel.Range).Text.ToString();
                r.Key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                r.Value = (sheet.Cells[row, 4] as Excel.Range).Text.ToString();

                rd.Resx.AddResxRow(r);

                bool hasCulture = true;
                int col = 5;
                while (hasCulture)
                {
                    string cult = (sheet.Cells[2, col] as Excel.Range).Text.ToString();

                    if (String.IsNullOrEmpty(cult))
                        break;

                    ResxData.ResxLocalizedRow lr = rd.ResxLocalized.NewResxLocalizedRow();

                    lr.Culture = cult;
                    lr.Key = (sheet.Cells[row, 3] as Excel.Range).Text.ToString();
                    lr.Value = (sheet.Cells[row, col] as Excel.Range).Text.ToString();
                    lr.ParentId = r.Id;

                    lr.SetParentRow(r);

                    rd.ResxLocalized.AddResxLocalizedRow(lr);

                    col++;
                }

                row++;
            }

            rd.AcceptChanges();

            wb.Close(false, m_objOpt, m_objOpt);
            app.Quit();

            return rd;
        }

        private void DataSetToXls(ResxData rd, string fileName)
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

            DataView dw = rd.Resx.DefaultView;
            dw.Sort = "FileSource, Key";

            int row = 3;

            foreach (DataRowView drw in dw )
            {
                ResxData.ResxRow r = (ResxData.ResxRow)drw.Row;

                sheet.Cells[row, 1] = r.FileSource;
                sheet.Cells[row, 2] = r.FileDestination;
                sheet.Cells[row, 3] = r.Key;
                sheet.Cells[row, 4] = r.Value;

                ResxData.ResxLocalizedRow[] rows = r.GetResxLocalizedRows();

                foreach (ResxData.ResxLocalizedRow lr in rows)
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
        }

        private string[] GetCulturesFromDataSet(ResxData rd)
        {
            if (rd.ResxLocalized.Rows.Count > 0)
            {
                ArrayList list = new ArrayList();
                foreach (ResxData.ResxLocalizedRow r in rd.ResxLocalized.Rows)
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

        private void FinishWizard()
        {
            Cursor = Cursors.WaitCursor;

            string[] excludeList = this.textBoxExclude.Text.Split(';');

            string[] cultures = null;

            cultures = new string[this.listBoxSelected.Items.Count];
            for (int i = 0; i < this.listBoxSelected.Items.Count; i++)
            {
                cultures[i] = ((CultureInfo)this.listBoxSelected.Items[i]).Name;
            }

            switch (_operation)
            {
                case ResxToXlsOperation.Create:

                    if (String.IsNullOrEmpty(this.textBoxFolder.Text))
                    {
                        MessageBox.Show("You must select a the .Net Project root wich contains your updated resx files", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.wizardControl1.CurrentStepIndex = this.intermediateStepProject.StepIndex;

                        return;
                    }

                    if (this.saveFileDialogXls.ShowDialog() == DialogResult.OK)
                    {
                        Application.DoEvents();

                        string path = this.saveFileDialogXls.FileName;

                        ResxToXls(this.textBoxFolder.Text, this.checkBoxSubFolders.Checked, path, cultures, excludeList, this.checkBoxFolderNaming.Checked);

                        MessageBox.Show("Excel Document created.", "Create", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    break;
                case ResxToXlsOperation.Build:
                    if (String.IsNullOrEmpty(this.textBoxXls.Text))
                    {
                        MessageBox.Show("You must select the Excel document to update", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.wizardControl1.CurrentStepIndex = this.intermediateStepXlsSelect.StepIndex;

                        return;
                    }

                    XlsToResx(this.textBoxXls.Text);

                    MessageBox.Show("Localized Resources created.", "Build", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    break;
                case ResxToXlsOperation.Update:
                    if (String.IsNullOrEmpty(this.textBoxFolder.Text))
                    {
                        MessageBox.Show("You must select a the .Net Project root wich contains your updated resx files", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.wizardControl1.CurrentStepIndex = this.intermediateStepProject.StepIndex;

                        return;
                    }

                    if (String.IsNullOrEmpty(this.textBoxXls.Text))
                    {
                        MessageBox.Show("You must select the Excel document to update", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.wizardControl1.CurrentStepIndex = this.intermediateStepXlsSelect.StepIndex;

                        return;
                    }


                    UpdateXls(this.textBoxXls.Text, this.textBoxFolder.Text, this.checkBoxSubFolders.Checked, excludeList, this.checkBoxFolderNaming.Checked);

                    MessageBox.Show("Excel Document Updated.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                default:
                    break;
            }

            Cursor = Cursors.Default;

            this.Close();
        }

        private void wizardControl1_CurrentStepIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void wizardControl1_NextButtonClick(WizardBase.WizardControl sender, WizardBase.WizardNextButtonClickEventArgs args)
        {
            int index = this.wizardControl1.CurrentStepIndex;

            int offset = 1; // è un bug? se non faccio così

            switch (index)
            {
                case 0:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Create:
                            this.wizardControl1.CurrentStepIndex = 1 - offset;
                            break;
                        case ResxToXlsOperation.Build:
                            this.wizardControl1.CurrentStepIndex = 4 - offset;
                            break;
                        case ResxToXlsOperation.Update:
                            this.wizardControl1.CurrentStepIndex = 1 - offset;
                            break;
                        default:
                            break;
                    }
                    break;

                case 1:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Update:
                            this.wizardControl1.CurrentStepIndex = 4 - offset;
                            break;
                        default:
                            break;
                    }
                    break;


                case 3:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Create:
                            this.wizardControl1.CurrentStepIndex = 5 - offset;
                            break;
                        default:
                            break;
                    }
                    break;
            }
        }

        private void wizardControl1_BackButtonClick(WizardBase.WizardControl sender, WizardBase.WizardClickEventArgs args)
        {
            int index = this.wizardControl1.CurrentStepIndex;

            int offset = 1; // è un bug? se non faccio così

            switch (index)
            {
                case 5:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Create:
                            this.wizardControl1.CurrentStepIndex = 3 + offset;
                            break;
                        default:
                            break;
                    }
                    break;
                case 4:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Build:
                            this.wizardControl1.CurrentStepIndex = 0 + offset;
                            break;
                        case ResxToXlsOperation.Update:
                            this.wizardControl1.CurrentStepIndex = 1 + offset;
                            break;
                        default:
                            break;

                    }
                    break;
            }
        }

        private void wizardControl1_FinishButtonClick(object sender, EventArgs e)
        {
            FinishWizard();
        }

        private void startStep1_Click(object sender, EventArgs e)
        {

        }

        
    }
}