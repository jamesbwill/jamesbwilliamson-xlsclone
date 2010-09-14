using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Windows.Forms;


namespace XlsLocalizationTool
{
    public partial class XlsLocalizationForm : Form
    {
        enum ResxToXlsOperation { Create, Build, Utf8Properties, Update };

        private ResxToXlsOperation _operation;

        string _summary1, _summary2,_summary3, _summary4;

        private XlsLocalizationManager _manager;

        public XlsLocalizationForm()
        {
            //CultureInfo ci = new CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            //System.Threading.Thread.CurrentThread.CurrentUICulture = ci;

            InitializeComponent();

            _manager = new XlsLocalizationManager();

            this.textBoxFolder.Text = Properties.Settings.Default.FolderPath;
            this.textBoxExclude.Text = Properties.Settings.Default.ExcludeList;
            this.checkBoxFolderNaming.Checked = Properties.Settings.Default.FolderNamespaceNaming;

            FillCultures();

            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonUpdateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildUtf8Properties.CheckedChanged += new EventHandler(radioButton_CheckedChanged);


            _summary1 = "Operation:\r\nCreate a new Excel document ready for localization.";
            _summary2 = "Operation:\r\nBuild your localized Resource files from a Filled Excel Document.";
            _summary3 = "Operation:\r\nUpdate your Excel document with your Project Resource changes.";
            _summary4 = "Operation:\r\nBuild a UTF8 encoded properties file from a Filled Excel Document.";

            this.textBoxSummary.Text = _summary1;
        }

        void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            this.radioButtonCreateXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonUpdateXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildUtf8Properties.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);


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

            if (this.radioButtonBuildUtf8Properties.Checked)
            {
                _operation = ResxToXlsOperation.Utf8Properties;
                this.textBoxSummary.Text = _summary4;
            }

            if (((RadioButton)sender).Checked)
            {
                if (((RadioButton)sender) == this.radioButtonCreateXls)
                {
                    this.radioButtonBuildXls.Checked = false;
                    this.radioButtonUpdateXls.Checked = false;
                    this.radioButtonBuildUtf8Properties.Checked = false;
                }

                if (((RadioButton)sender) == this.radioButtonBuildXls)
                {
                    this.radioButtonCreateXls.Checked = false;
                    this.radioButtonUpdateXls.Checked = false;
                    this.radioButtonBuildUtf8Properties.Checked = false;
                }

                if (((RadioButton)sender) == this.radioButtonUpdateXls)
                {
                    this.radioButtonCreateXls.Checked = false;
                    this.radioButtonBuildXls.Checked = false;
                    this.radioButtonBuildUtf8Properties.Checked = false;
                }

                if (((RadioButton)sender) == this.radioButtonUpdateXls)
                {
                    this.radioButtonCreateXls.Checked = false;
                    this.radioButtonBuildXls.Checked = false;
                    this.radioButtonUpdateXls.Checked = false;
                }

            }
            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonUpdateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildUtf8Properties.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
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

                        try
                        { 
                            _manager.ResxToXls(this.textBoxFolder.Text, this.checkBoxSubFolders.Checked, path, cultures, excludeList, this.checkBoxFolderNaming.Checked);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("A problem occured converting resource files" + "\n" + ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

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

                    _manager.XlsToResx(this.textBoxXls.Text);

                    MessageBox.Show("Localized Resources created.", "Build", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    break;

                case ResxToXlsOperation.Utf8Properties:
                    if (String.IsNullOrEmpty(this.textBoxXls.Text))
                    {
                        MessageBox.Show("You must select a the .Net Project root wich contains your updated resx files", "Utf8Properties", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.wizardControl1.CurrentStepIndex = this.intermediateStepXlsSelect.StepIndex;

                        return;
                    }

                    _manager.XlsToUTF8Properties(this.textBoxXls.Text, String.Empty);

                    MessageBox.Show("Localized UTF8 properties resources created.", "BuildUTF8", MessageBoxButtons.OK, MessageBoxIcon.Information);

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


                    _manager.UpdateXls(this.textBoxXls.Text, this.textBoxFolder.Text, this.checkBoxSubFolders.Checked, excludeList, this.checkBoxFolderNaming.Checked);

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

        private void radioButtonUpdateXls_CheckedChanged(object sender, EventArgs e)
        {

        }

        
    }
}
