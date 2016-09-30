namespace Tester
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.btnNewExcel = new System.Windows.Forms.Button();
            this.btnExistingExcel = new System.Windows.Forms.Button();
            this.btnExcelByProcess = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnAutoFilter = new System.Windows.Forms.Button();
            this.btnExcelListOpen = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnWordByProc = new System.Windows.Forms.Button();
            this.btnExistingWord = new System.Windows.Forms.Button();
            this.btnNewWord = new System.Windows.Forms.Button();
            this.grpADOB = new System.Windows.Forms.GroupBox();
            this.btnFromDataTable = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.cbxControlBars = new System.Windows.Forms.ComboBox();
            this.btnExistingAccess = new System.Windows.Forms.Button();
            this.btnNewAccess = new System.Windows.Forms.Button();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnOutlookProcess = new System.Windows.Forms.Button();
            this.btnOutlookExisting = new System.Windows.Forms.Button();
            this.btnNewOutlook = new System.Windows.Forms.Button();
            this.ExcelToPdfButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpADOB.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnNewExcel
            // 
            this.btnNewExcel.Location = new System.Drawing.Point(15, 19);
            this.btnNewExcel.Name = "btnNewExcel";
            this.btnNewExcel.Size = new System.Drawing.Size(98, 23);
            this.btnNewExcel.TabIndex = 0;
            this.btnNewExcel.Text = "New workbook";
            this.btnNewExcel.UseVisualStyleBackColor = true;
            this.btnNewExcel.Click += new System.EventHandler(this.btnNewExcel_Click);
            // 
            // btnExistingExcel
            // 
            this.btnExistingExcel.Location = new System.Drawing.Point(119, 19);
            this.btnExistingExcel.Name = "btnExistingExcel";
            this.btnExistingExcel.Size = new System.Drawing.Size(106, 23);
            this.btnExistingExcel.TabIndex = 1;
            this.btnExistingExcel.Text = "Existing workbook";
            this.btnExistingExcel.UseVisualStyleBackColor = true;
            this.btnExistingExcel.Click += new System.EventHandler(this.btnNewExcel_Click);
            // 
            // btnExcelByProcess
            // 
            this.btnExcelByProcess.Location = new System.Drawing.Point(15, 48);
            this.btnExcelByProcess.Name = "btnExcelByProcess";
            this.btnExcelByProcess.Size = new System.Drawing.Size(98, 23);
            this.btnExcelByProcess.TabIndex = 2;
            this.btnExcelByProcess.Text = "Existing process";
            this.btnExcelByProcess.UseVisualStyleBackColor = true;
            this.btnExcelByProcess.Click += new System.EventHandler(this.btnExcelByProcess_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ExcelToPdfButton);
            this.groupBox1.Controls.Add(this.btnAutoFilter);
            this.groupBox1.Controls.Add(this.btnNewExcel);
            this.groupBox1.Controls.Add(this.btnExcelListOpen);
            this.groupBox1.Controls.Add(this.btnExcelByProcess);
            this.groupBox1.Controls.Add(this.btnExistingExcel);
            this.groupBox1.Location = new System.Drawing.Point(12, 38);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(253, 131);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel";
            // 
            // btnAutoFilter
            // 
            this.btnAutoFilter.Location = new System.Drawing.Point(15, 89);
            this.btnAutoFilter.Name = "btnAutoFilter";
            this.btnAutoFilter.Size = new System.Drawing.Size(98, 23);
            this.btnAutoFilter.TabIndex = 0;
            this.btnAutoFilter.Text = "AutoFilter";
            this.btnAutoFilter.UseVisualStyleBackColor = true;
            this.btnAutoFilter.Click += new System.EventHandler(this.btnAutoFilter_Click);
            // 
            // btnExcelListOpen
            // 
            this.btnExcelListOpen.Location = new System.Drawing.Point(119, 89);
            this.btnExcelListOpen.Name = "btnExcelListOpen";
            this.btnExcelListOpen.Size = new System.Drawing.Size(113, 23);
            this.btnExcelListOpen.TabIndex = 2;
            this.btnExcelListOpen.Text = "List open workbooks";
            this.btnExcelListOpen.UseVisualStyleBackColor = true;
            this.btnExcelListOpen.Click += new System.EventHandler(this.btnExcelListOpen_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.btnWordByProc);
            this.groupBox2.Controls.Add(this.btnExistingWord);
            this.groupBox2.Controls.Add(this.btnNewWord);
            this.groupBox2.Location = new System.Drawing.Point(271, 38);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(255, 100);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Word";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(126, 49);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "SavePDF";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnWordByProc
            // 
            this.btnWordByProc.Location = new System.Drawing.Point(6, 48);
            this.btnWordByProc.Name = "btnWordByProc";
            this.btnWordByProc.Size = new System.Drawing.Size(113, 23);
            this.btnWordByProc.TabIndex = 3;
            this.btnWordByProc.Text = "Existing process";
            this.btnWordByProc.UseVisualStyleBackColor = true;
            this.btnWordByProc.Click += new System.EventHandler(this.btnWordByProc_Click);
            // 
            // btnExistingWord
            // 
            this.btnExistingWord.Location = new System.Drawing.Point(120, 19);
            this.btnExistingWord.Name = "btnExistingWord";
            this.btnExistingWord.Size = new System.Drawing.Size(108, 23);
            this.btnExistingWord.TabIndex = 1;
            this.btnExistingWord.Text = "Existing document";
            this.btnExistingWord.UseVisualStyleBackColor = true;
            this.btnExistingWord.Click += new System.EventHandler(this.btnExistingWord_Click);
            // 
            // btnNewWord
            // 
            this.btnNewWord.Location = new System.Drawing.Point(6, 19);
            this.btnNewWord.Name = "btnNewWord";
            this.btnNewWord.Size = new System.Drawing.Size(108, 23);
            this.btnNewWord.TabIndex = 0;
            this.btnNewWord.Text = "New document";
            this.btnNewWord.UseVisualStyleBackColor = true;
            this.btnNewWord.Click += new System.EventHandler(this.btnExistingWord_Click);
            // 
            // grpADOB
            // 
            this.grpADOB.Controls.Add(this.btnFromDataTable);
            this.grpADOB.Location = new System.Drawing.Point(12, 175);
            this.grpADOB.Name = "grpADOB";
            this.grpADOB.Size = new System.Drawing.Size(253, 96);
            this.grpADOB.TabIndex = 5;
            this.grpADOB.TabStop = false;
            this.grpADOB.Text = "ADODB";
            // 
            // btnFromDataTable
            // 
            this.btnFromDataTable.Location = new System.Drawing.Point(15, 19);
            this.btnFromDataTable.Name = "btnFromDataTable";
            this.btnFromDataTable.Size = new System.Drawing.Size(98, 23);
            this.btnFromDataTable.TabIndex = 0;
            this.btnFromDataTable.Text = "From Datatable";
            this.btnFromDataTable.UseVisualStyleBackColor = true;
            this.btnFromDataTable.Click += new System.EventHandler(this.btnFromDataTable_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.cbxControlBars);
            this.groupBox3.Controls.Add(this.btnExistingAccess);
            this.groupBox3.Controls.Add(this.btnNewAccess);
            this.groupBox3.Location = new System.Drawing.Point(271, 175);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(255, 96);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Access";
            // 
            // cbxControlBars
            // 
            this.cbxControlBars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxControlBars.FormattingEnabled = true;
            this.cbxControlBars.Location = new System.Drawing.Point(27, 52);
            this.cbxControlBars.Name = "cbxControlBars";
            this.cbxControlBars.Size = new System.Drawing.Size(201, 21);
            this.cbxControlBars.TabIndex = 3;
            this.cbxControlBars.SelectedIndexChanged += new System.EventHandler(this.cbxControlBars_SelectedIndexChanged);
            // 
            // btnExistingAccess
            // 
            this.btnExistingAccess.Location = new System.Drawing.Point(120, 19);
            this.btnExistingAccess.Name = "btnExistingAccess";
            this.btnExistingAccess.Size = new System.Drawing.Size(108, 23);
            this.btnExistingAccess.TabIndex = 2;
            this.btnExistingAccess.Text = "Existing instance";
            this.btnExistingAccess.UseVisualStyleBackColor = true;
            this.btnExistingAccess.Click += new System.EventHandler(this.btnExistingAccess_Click);
            // 
            // btnNewAccess
            // 
            this.btnNewAccess.Location = new System.Drawing.Point(6, 19);
            this.btnNewAccess.Name = "btnNewAccess";
            this.btnNewAccess.Size = new System.Drawing.Size(108, 23);
            this.btnNewAccess.TabIndex = 0;
            this.btnNewAccess.Text = "New instance";
            this.btnNewAccess.UseVisualStyleBackColor = true;
            this.btnNewAccess.Click += new System.EventHandler(this.btnNewAccess_Click);
            // 
            // toolStrip
            // 
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(802, 25);
            this.toolStrip.TabIndex = 6;
            this.toolStrip.Text = "toolStrip1";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnOutlookProcess);
            this.groupBox4.Controls.Add(this.btnOutlookExisting);
            this.groupBox4.Controls.Add(this.btnNewOutlook);
            this.groupBox4.Location = new System.Drawing.Point(532, 42);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(253, 96);
            this.groupBox4.TabIndex = 5;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Outlook";
            // 
            // btnOutlookProcess
            // 
            this.btnOutlookProcess.Location = new System.Drawing.Point(119, 48);
            this.btnOutlookProcess.Name = "btnOutlookProcess";
            this.btnOutlookProcess.Size = new System.Drawing.Size(98, 23);
            this.btnOutlookProcess.TabIndex = 0;
            this.btnOutlookProcess.Text = "Existing process";
            this.btnOutlookProcess.UseVisualStyleBackColor = true;
            this.btnOutlookProcess.Click += new System.EventHandler(this.btnOutlookProcess_Click);
            // 
            // btnOutlookExisting
            // 
            this.btnOutlookExisting.Location = new System.Drawing.Point(15, 48);
            this.btnOutlookExisting.Name = "btnOutlookExisting";
            this.btnOutlookExisting.Size = new System.Drawing.Size(98, 23);
            this.btnOutlookExisting.TabIndex = 0;
            this.btnOutlookExisting.Text = "Existing ROT";
            this.btnOutlookExisting.UseVisualStyleBackColor = true;
            this.btnOutlookExisting.Click += new System.EventHandler(this.btnOutlookExisting_Click);
            // 
            // btnNewOutlook
            // 
            this.btnNewOutlook.Location = new System.Drawing.Point(15, 19);
            this.btnNewOutlook.Name = "btnNewOutlook";
            this.btnNewOutlook.Size = new System.Drawing.Size(98, 23);
            this.btnNewOutlook.TabIndex = 0;
            this.btnNewOutlook.Text = "New Instance";
            this.btnNewOutlook.UseVisualStyleBackColor = true;
            this.btnNewOutlook.Click += new System.EventHandler(this.btnNewOutlook_Click);
            // 
            // ExcelToPdfButton
            // 
            this.ExcelToPdfButton.Location = new System.Drawing.Point(119, 48);
            this.ExcelToPdfButton.Name = "ExcelToPdfButton";
            this.ExcelToPdfButton.Size = new System.Drawing.Size(106, 23);
            this.ExcelToPdfButton.TabIndex = 3;
            this.ExcelToPdfButton.Text = "To pdf";
            this.ExcelToPdfButton.UseVisualStyleBackColor = true;
            this.ExcelToPdfButton.Click += new System.EventHandler(this.ExcelToPdfButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(802, 466);
            this.Controls.Add(this.toolStrip);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.grpADOB);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.grpADOB.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnNewExcel;
        private System.Windows.Forms.Button btnExistingExcel;
        private System.Windows.Forms.Button btnExcelByProcess;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnNewWord;
        private System.Windows.Forms.Button btnExistingWord;
        private System.Windows.Forms.Button btnWordByProc;
        private System.Windows.Forms.GroupBox grpADOB;
        private System.Windows.Forms.Button btnFromDataTable;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnNewAccess;
        private System.Windows.Forms.Button btnExistingAccess;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.ComboBox cbxControlBars;
        private System.Windows.Forms.Button btnAutoFilter;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnExcelListOpen;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btnNewOutlook;
        private System.Windows.Forms.Button btnOutlookExisting;
        private System.Windows.Forms.Button btnOutlookProcess;
        private System.Windows.Forms.Button ExcelToPdfButton;
    }
}

