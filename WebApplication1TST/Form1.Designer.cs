using DataImport;
using Microsoft.SharePoint;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WebApplication1TST
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        public DataImport.SharePoint SP = new SharePoint();
        public DataImport.OpenXML blOpenXML = new OpenXML();
        public List<SPField> fields;
        public SPList list;
        public string webUrl;
        public SPWeb web;
        public DataImport.BLDataImport DataImport = new BLDataImport();

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SelectSpreadsheet = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ExFieldTitle = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ListColumn1 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.SelecList = new System.Windows.Forms.ComboBox();
            this.selectWebs = new System.Windows.Forms.ComboBox();
            this.DateFormat = new System.Windows.Forms.TextBox();
            this.StartImport = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.StartTime = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.EndTime = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.ProcessedItems = new System.Windows.Forms.Label();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.RemoveKey = new System.Windows.Forms.Button();
            this.AddKey = new System.Windows.Forms.Button();
            this.ActiveKeys = new System.Windows.Forms.ListBox();
            this.PossibleKeys = new System.Windows.Forms.ListBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.ValidateData = new System.Windows.Forms.Button();
            this.DateExample = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(31, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Select File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(196, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Filename";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "SelectSheet";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(30, 93);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Select Web";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(31, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Select List";
            // 
            // SelectSpreadsheet
            // 
            this.SelectSpreadsheet.FormattingEnabled = true;
            this.SelectSpreadsheet.Location = new System.Drawing.Point(199, 65);
            this.SelectSpreadsheet.Name = "SelectSpreadsheet";
            this.SelectSpreadsheet.Size = new System.Drawing.Size(121, 21);
            this.SelectSpreadsheet.TabIndex = 5;
            this.SelectSpreadsheet.SelectedValueChanged += new System.EventHandler(this.SelectSpreadsheet_SelectedIndexChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ExFieldTitle,
            this.ListColumn1});
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Location = new System.Drawing.Point(12, 234);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(515, 267);
            this.dataGridView1.TabIndex = 6;
            // 
            // ExFieldTitle
            // 
            this.ExFieldTitle.HeaderText = "Sheet Column Name";
            this.ExFieldTitle.Name = "ExFieldTitle";
            this.ExFieldTitle.ReadOnly = true;
            this.ExFieldTitle.Width = 200;
            // 
            // ListColumn1
            // 
            this.ListColumn1.HeaderText = "Sherepoint List Field";
            this.ListColumn1.Name = "ListColumn1";
            this.ListColumn1.Width = 200;
            // 
            // SelecList
            // 
            this.SelecList.FormattingEnabled = true;
            this.SelecList.Location = new System.Drawing.Point(199, 117);
            this.SelecList.Name = "SelecList";
            this.SelecList.Size = new System.Drawing.Size(121, 21);
            this.SelecList.TabIndex = 7;
            this.SelecList.SelectedValueChanged += new System.EventHandler(this.SelecList_SelectedIndexChanged);
            // 
            // selectWebs
            // 
            this.selectWebs.FormattingEnabled = true;
            this.selectWebs.Location = new System.Drawing.Point(199, 90);
            this.selectWebs.Name = "selectWebs";
            this.selectWebs.Size = new System.Drawing.Size(121, 21);
            this.selectWebs.TabIndex = 8;
            this.selectWebs.SelectedValueChanged += new System.EventHandler(this.SelectWebs_SelectedIndexChanged);
            // 
            // DateFormat
            // 
            this.DateFormat.Location = new System.Drawing.Point(199, 144);
            this.DateFormat.Name = "DateFormat";
            this.DateFormat.Size = new System.Drawing.Size(121, 20);
            this.DateFormat.TabIndex = 9;
            this.DateFormat.Text = "MM/dd/yyyy";
            // 
            // StartImport
            // 
            this.StartImport.Location = new System.Drawing.Point(557, 246);
            this.StartImport.Name = "StartImport";
            this.StartImport.Size = new System.Drawing.Size(75, 23);
            this.StartImport.TabIndex = 10;
            this.StartImport.Text = "Start Import";
            this.StartImport.UseVisualStyleBackColor = true;
            this.StartImport.Click += new System.EventHandler(this.StartImport_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(551, 298);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(55, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Start Time";
            // 
            // StartTime
            // 
            this.StartTime.AutoSize = true;
            this.StartTime.Location = new System.Drawing.Point(645, 298);
            this.StartTime.Name = "StartTime";
            this.StartTime.Size = new System.Drawing.Size(49, 13);
            this.StartTime.TabIndex = 12;
            this.StartTime.Text = "00:00:00";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(551, 324);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "End Time";
            // 
            // EndTime
            // 
            this.EndTime.AutoSize = true;
            this.EndTime.Location = new System.Drawing.Point(645, 324);
            this.EndTime.Name = "EndTime";
            this.EndTime.Size = new System.Drawing.Size(49, 13);
            this.EndTime.TabIndex = 14;
            this.EndTime.Text = "00:00:00";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(551, 350);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(81, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = "Records Added";
            // 
            // ProcessedItems
            // 
            this.ProcessedItems.AccessibleRole = System.Windows.Forms.AccessibleRole.WhiteSpace;
            this.ProcessedItems.AutoSize = true;
            this.ProcessedItems.Location = new System.Drawing.Point(665, 346);
            this.ProcessedItems.Name = "ProcessedItems";
            this.ProcessedItems.Size = new System.Drawing.Size(13, 13);
            this.ProcessedItems.TabIndex = 16;
            this.ProcessedItems.Text = "0";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 507);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(515, 157);
            this.richTextBox1.TabIndex = 17;
            this.richTextBox1.Text = "";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(31, 147);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(62, 13);
            this.label7.TabIndex = 18;
            this.label7.Text = "Date format";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.RemoveKey);
            this.panel1.Controls.Add(this.AddKey);
            this.panel1.Controls.Add(this.ActiveKeys);
            this.panel1.Controls.Add(this.PossibleKeys);
            this.panel1.Location = new System.Drawing.Point(402, 22);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(386, 159);
            this.panel1.TabIndex = 19;
            // 
            // RemoveKey
            // 
            this.RemoveKey.Location = new System.Drawing.Point(155, 71);
            this.RemoveKey.Name = "RemoveKey";
            this.RemoveKey.Size = new System.Drawing.Size(75, 23);
            this.RemoveKey.TabIndex = 3;
            this.RemoveKey.Text = "<- Remove";
            this.RemoveKey.UseVisualStyleBackColor = true;
            this.RemoveKey.Click += new System.EventHandler(this.RemoveKey_Click);
            // 
            // AddKey
            // 
            this.AddKey.Location = new System.Drawing.Point(155, 33);
            this.AddKey.Name = "AddKey";
            this.AddKey.Size = new System.Drawing.Size(75, 23);
            this.AddKey.TabIndex = 2;
            this.AddKey.Text = "Add ->";
            this.AddKey.UseVisualStyleBackColor = true;
            this.AddKey.Click += new System.EventHandler(this.AddKey_Click);
            // 
            // ActiveKeys
            // 
            this.ActiveKeys.FormattingEnabled = true;
            this.ActiveKeys.Location = new System.Drawing.Point(246, 10);
            this.ActiveKeys.Name = "ActiveKeys";
            this.ActiveKeys.Size = new System.Drawing.Size(120, 134);
            this.ActiveKeys.TabIndex = 1;
            // 
            // PossibleKeys
            // 
            this.PossibleKeys.FormattingEnabled = true;
            this.PossibleKeys.Location = new System.Drawing.Point(14, 10);
            this.PossibleKeys.Name = "PossibleKeys";
            this.PossibleKeys.Size = new System.Drawing.Size(120, 134);
            this.PossibleKeys.TabIndex = 0;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 200);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(776, 23);
            this.progressBar1.TabIndex = 20;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(554, 404);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(163, 17);
            this.checkBox1.TabIndex = 21;
            this.checkBox1.Text = "Create missing lookup values";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // ValidateData
            // 
            this.ValidateData.Location = new System.Drawing.Point(668, 246);
            this.ValidateData.Name = "ValidateData";
            this.ValidateData.Size = new System.Drawing.Size(75, 23);
            this.ValidateData.TabIndex = 22;
            this.ValidateData.Text = "Validate";
            this.ValidateData.UseVisualStyleBackColor = true;
            // 
            // DateExample
            // 
            this.DateExample.Location = new System.Drawing.Point(199, 167);
            this.DateExample.Name = "DateExample";
            this.DateExample.Size = new System.Drawing.Size(121, 13);
            this.DateExample.TabIndex = 23;
            this.DateExample.Text = "02/14/2019";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 676);
            this.Controls.Add(this.DateExample);
            this.Controls.Add(this.ValidateData);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.ProcessedItems);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.EndTime);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.StartTime);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.StartImport);
            this.Controls.Add(this.DateFormat);
            this.Controls.Add(this.selectWebs);
            this.Controls.Add(this.SelecList);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.SelectSpreadsheet);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Sharepoint Data Import Tool";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox SelectSpreadsheet;
        private System.Windows.Forms.DataGridView dataGridView1;
        //private DataGridViewTextBoxColumn ExFieldTitle;
        //private DataGridViewComboBoxColumn ListColumn1;
        private ComboBox SelecList;
        private ComboBox selectWebs;
        private TextBox DateFormat;
        private Button StartImport;
        private Label label5;
        private Label StartTime;
        private Label label6;
        private Label EndTime;
        private Label label8;
        private Label ProcessedItems;
        private RichTextBox richTextBox1;
        private DataGridViewTextBoxColumn ExFieldTitle;
        private DataGridViewComboBoxColumn ListColumn1;
        private Label label7;
        private Panel panel1;
        private ListBox PossibleKeys;
        private ListBox ActiveKeys;
        private Button RemoveKey;
        private Button AddKey;
        private ProgressBar progressBar1;
        private CheckBox checkBox1;
        private Button ValidateData;
        private Label DateExample;
    }
}