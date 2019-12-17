using DataImport;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
                 List<string> sheets = new List<string>();

                    using (var document = SpreadsheetDocument.Open(fileStream, false))
                    {
                        sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Select(sh => sh.Name.Value).ToList();
                    }
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

        private void StartImport_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            Cursor = Cursors.WaitCursor;
                       
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
                }
            }
            if (mapping.Count == 0)
                return;

            //GenarateHash
            Dictionary<string, Hashtable> HashDictionary = null;
            string TimeFormat = DateFormat.Text;
            List<string> SelectedKeys = new List<string>();
            Dictionary<string, int> SelectedKeysRows = new Dictionary<string, int>();
            Dictionary<string, SPList> LookupRelations = new Dictionary<string, SPList>();

            if (ActiveKeys.Items.Count > 0)
            {
                SelectedKeys = ActiveKeys.Items.Cast<String>().ToList();
                HashDictionary = GenerateHashTable.HashTableForListField(list, SelectedKeys, mapping, web, TimeFormat, ref LookupRelations);

                for (var i = 0; i < SelectedKeys.Count; i++)
                {
                    int index = headers.IndexOf(SelectedKeys[i]);
                    if (index != -1)
                    {
                        SelectedKeysRows.Add(SelectedKeys[i], index);
                    }

                }


            }
            else {
                HashDictionary = GenerateHashTable.HashTableForListField(list, SelectedKeys, mapping, web, TimeFormat, ref LookupRelations);
            }
             //Hash is ready

            var fileStream = openFileDialog1.OpenFile();
            using (fileStream)
            {
                var dataTable = blOpenXML.ImportToDataTable(fileStream, SelectSpreadsheet.SelectedItem.ToString(), DateFormat.Text, false);
                DateTime starttime = DateTime.Now;
                DataImport.ProcessDataImport(list, mapping, dataTable, DateFormat.Text, web, LookupRelations, TimeFormat, HashDictionary, SelectedKeysRows);
                richTextBox1.Text = DataImport.GetErrors();
               

                StartTime.Text = starttime.ToString("g");
                EndTime.Text = DateTime.Now.ToString("g");
                //long diff = DateAndTime.DateDiff(DateInterval.Minute, starttime, DateTime.Now);
                ProcessedItems.Text = DataImport.GetNumberOfItems().ToString();
                //TimeTaken.Text = Conversions.ToString(diff);
            }
            Cursor = Cursors.Arrow;
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
