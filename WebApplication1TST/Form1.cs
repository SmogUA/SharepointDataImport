using DataImport;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WebApplication1TST
{
    public partial class Form1 : Form
    {

        private List<string> headers;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = this.openFileDialog1.FileName;
                label1.Visible = true;
                label2.Visible = true;
                var fileStream = openFileDialog1.OpenFile();
                using (fileStream)
                {

                    SelectSpreadsheet.DataSource = blOpenXML.GetSheetsFromFile(fileStream);
                    SelectSpreadsheet.Visible = true;                    
                }
                selectWebs.DataSource = SP.GetAllWebs().Select(w => w.Url).ToList();
            }

        }

        private void SelectWebs_SelectedIndexChanged(object sender, EventArgs e)
        {
            webUrl = selectWebs.SelectedItem.ToString();
            web = SP.GetWeb(webUrl);
            var lists = SP.GetAllLists(webUrl).Select(l => l.Title).ToList();
            SelecList.DataSource = lists;
        }

        private void SelectSpreadsheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            SelecList.SelectedItem = null;
        }

        private void SelecList_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (SelecList.SelectedItem != null)
            {
                list = SP.GetListByDisplayName(selectWebs.SelectedItem.ToString(), SelecList.SelectedItem.ToString());
                fields = SP.GetListFields(list);
                var fileStream = openFileDialog1.OpenFile();
                using (fileStream)
                {
                    var dataTable = blOpenXML.ImportToDataTable(fileStream, SelectSpreadsheet.SelectedItem.ToString(), DateFormat.Text, true);
                    headers = new List<string>();
                    for (int i = 0, loopTo = dataTable.Columns.Count - 1; i <= loopTo; i++)
                        headers.Add(dataTable.Columns[i].ColumnName.Trim());
                    PossibleKeys.DataSource=headers;
                    foreach (var header in headers)
                    {
                        
                        var row = new DataGridViewRow();
                        row.Cells.Add(new DataGridViewTextBoxCell() { Value = header });
                        var combocell = new DataGridViewComboBoxCell();
                        combocell.DataSource = fields.Select(f => f.InternalName).ToList();
                        row.Cells.Add(combocell);
                        dataGridView1.Rows.Add(row);
                    }
                }
            }
        }

        private async void StartImport_Click(object sender, EventArgs e)
        {
            DateTime starttime = DateTime.Now;
            StartTime.Text = starttime.ToString("g");
            button1.Enabled = false;
            AddKey.Enabled = false;
            RemoveKey.Enabled = false;
            StartImport.Enabled = false;
            ValidateData.Enabled = false;
            bool CreateMissingLookupValues = checkBox1.Checked;

            richTextBox1.Clear();
            Cursor = Cursors.WaitCursor;
            Progress<int> progress = new Progress<int>(value => { progressBar1.Value = value; ProcessedItems.Text = value.ToString(); });
            List<DIMapping> mapping = new List<DIMapping>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var combocell = (DataGridViewComboBoxCell)row.Cells[1];
                if (combocell.Value != null)
                {
                    string internalFieldName = combocell.Value.ToString();
                    var txtCell = (DataGridViewTextBoxCell)row.Cells[0];
                    var fileFieldName = txtCell.Value;
                    if (!string.IsNullOrEmpty(internalFieldName))
                        mapping.Add(new DIMapping() { Name = internalFieldName, Value = fileFieldName.ToString() });
                }
            }
            if (mapping.Count == 0)
            {

                Cursor = Cursors.Arrow;
                button1.Enabled = true;
                AddKey.Enabled = true;
                RemoveKey.Enabled = true;
                StartImport.Enabled = true;               
                EndTime.Text = DateTime.Now.ToString("g");
                return;
            }
            List<string> SelectedKeys = new List<string>();
            if (ActiveKeys.Items.Count > 0)
            {
                SelectedKeys = ActiveKeys.Items.Cast<String>().ToList();
            }
            string TabName = SelectSpreadsheet.SelectedItem.ToString();
            int ItemsToImport = CountItemsToImport();
            progressBar1.Maximum = ItemsToImport;
            Task<string> result = Task.Run(() => RunTask(progress, web.Url, list.Title, mapping, TabName, SelectedKeys, DateFormat.Text, CreateMissingLookupValues));
            richTextBox1.Text = await result;
            EndTime.Text = DateTime.Now.ToString("g");
            Cursor = Cursors.Arrow;
            button1.Enabled = true;
            AddKey.Enabled = true;
            RemoveKey.Enabled = true;
            StartImport.Enabled = true;
        }

        private int CountItemsToImport()
        {
            System.IO.Stream fileStream = openFileDialog1.OpenFile();
            int rez = 0;
            if (!string.IsNullOrEmpty(SelectSpreadsheet.SelectedItem.ToString()))
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, false))
                {
                    Sheet theSheet = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == SelectSpreadsheet.SelectedItem.ToString()).FirstOrDefault();
                    var wsPart = (WorksheetPart)document.WorkbookPart.GetPartById(theSheet.Id);
                    var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                    if (sheetData.Elements<Row>().Count() > 0) rez = sheetData.Elements<Row>().Count() - 1 ;
                }
            }
            return rez;
        }

        private async Task<string> RunTask(IProgress<int> progress,  string WebURL, string ListName, List<DIMapping> mapping, string TabName, List<string> SelectedKeys, string str_DateFormat, bool CreateMissingLookupValues)
        {
            string rez=String.Empty;
            Dictionary<string, Hashtable> HashDictionary = null;
            Dictionary<string, int> SelectedKeysRows = new Dictionary<string, int>();
            Dictionary<string, SPList> LookupRelations = new Dictionary<string, SPList>();
            SPList listloc = null;

            SPSecurity.RunWithElevatedPrivileges(() =>
            {

             SharePoint SP = new SharePoint();
             SPWeb Web = SP.GetWeb(WebURL);
             listloc = SP.GetListByDisplayName(WebURL, ListName);

            //GenarateHash    

             if (SelectedKeys.Count > 0)
            {
               
                HashDictionary = GenerateHashTable.HashTableForListField(listloc, SelectedKeys, mapping, web, str_DateFormat, ref LookupRelations);

                for (var i = 0; i < SelectedKeys.Count; i++)
                {
                    int index = headers.IndexOf(SelectedKeys[i]);
                    if (index != -1)
                    {
                        SelectedKeysRows.Add(SelectedKeys[i], index);
                    }

                }
            }
            else
            {
                HashDictionary = GenerateHashTable.HashTableForListField(listloc, SelectedKeys, mapping, web, str_DateFormat, ref LookupRelations);
            }
                
            });
            //Hash is ready

            System.IO.Stream fileStream = openFileDialog1.OpenFile();
            using (fileStream)
            {
                var dataTable = blOpenXML.ImportToDataTable(fileStream, TabName, DateFormat.Text, false);
                DateTime starttime = DateTime.Now;
                DataImport.ProcessDataImport(CreateMissingLookupValues, progress, listloc, mapping, dataTable, str_DateFormat, web, LookupRelations, str_DateFormat, HashDictionary, SelectedKeysRows);
                rez = DataImport.GetErrors();

            }            

            return rez;
        }

        private void AddKey_Click(object sender, EventArgs e)
        {
            if ((PossibleKeys.SelectedIndex != -1) && (!ActiveKeys.Items.Contains(PossibleKeys.Items[PossibleKeys.SelectedIndex]))) {
                ActiveKeys.Items.Add(PossibleKeys.Items[PossibleKeys.SelectedIndex]);
                
            };

        }

        private void RemoveKey_Click(object sender, EventArgs e)
        {
            if (ActiveKeys.SelectedIndex != -1)
            {
                ActiveKeys.Items.Remove(ActiveKeys.Items[ActiveKeys.SelectedIndex]);

            };

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            var mapping = new List<DIMapping>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var combocell = (DataGridViewComboBoxCell)row.Cells[1];
                if (combocell.Value != null)
                {
                    string internalFieldName = combocell.Value.ToString();
                    var txtCell = (DataGridViewTextBoxCell)row.Cells[0];
                    var fileFieldName = txtCell.Value;
                    if (!string.IsNullOrEmpty(internalFieldName))
                        mapping.Add(new DIMapping() { Name = internalFieldName, Value = fileFieldName.ToString() });

                };
            }
            if (mapping.Count > 0)
            {
                this.StartImport.Enabled = true;

            }
            else
            {
                this.StartImport.Enabled = false;
            };
                
            
        }

    }
}
