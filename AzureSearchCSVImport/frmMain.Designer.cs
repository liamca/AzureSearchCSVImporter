namespace AzureSearchCSVImport
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSearchService = new System.Windows.Forms.TextBox();
            this.txtApiKey = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnChooseFile = new System.Windows.Forms.Button();
            this.csvDataGrid = new System.Windows.Forms.DataGridView();
            this.searchDataGrid = new System.Windows.Forms.DataGridView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnCreateIndex = new System.Windows.Forms.Button();
            this.btnImportData = new System.Windows.Forms.Button();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.txtIndexName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.txtBatchSize = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.csvDataGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.searchDataGrid)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 40);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Search Service";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(344, 41);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "API Key";
            // 
            // txtSearchService
            // 
            this.txtSearchService.Location = new System.Drawing.Point(92, 38);
            this.txtSearchService.Margin = new System.Windows.Forms.Padding(2);
            this.txtSearchService.Name = "txtSearchService";
            this.txtSearchService.Size = new System.Drawing.Size(131, 20);
            this.txtSearchService.TabIndex = 4;
            this.txtSearchService.Leave += new System.EventHandler(this.txtSearchService_Leave);
            // 
            // txtApiKey
            // 
            this.txtApiKey.Location = new System.Drawing.Point(388, 37);
            this.txtApiKey.Margin = new System.Windows.Forms.Padding(2);
            this.txtApiKey.Name = "txtApiKey";
            this.txtApiKey.Size = new System.Drawing.Size(213, 20);
            this.txtApiKey.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(227, 41);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(104, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = ".search.windows.net";
            // 
            // btnChooseFile
            // 
            this.btnChooseFile.Location = new System.Drawing.Point(10, 8);
            this.btnChooseFile.Margin = new System.Windows.Forms.Padding(2);
            this.btnChooseFile.Name = "btnChooseFile";
            this.btnChooseFile.Size = new System.Drawing.Size(116, 25);
            this.btnChooseFile.TabIndex = 7;
            this.btnChooseFile.Text = "Choose CSV File";
            this.btnChooseFile.UseVisualStyleBackColor = true;
            this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
            // 
            // csvDataGrid
            // 
            this.csvDataGrid.AllowUserToAddRows = false;
            this.csvDataGrid.AllowUserToDeleteRows = false;
            this.csvDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.csvDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.csvDataGrid.Location = new System.Drawing.Point(8, 62);
            this.csvDataGrid.Margin = new System.Windows.Forms.Padding(2);
            this.csvDataGrid.Name = "csvDataGrid";
            this.csvDataGrid.RowTemplate.Height = 28;
            this.csvDataGrid.Size = new System.Drawing.Size(819, 198);
            this.csvDataGrid.TabIndex = 8;
            // 
            // searchDataGrid
            // 
            this.searchDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.searchDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.searchDataGrid.Location = new System.Drawing.Point(8, 282);
            this.searchDataGrid.Margin = new System.Windows.Forms.Padding(2);
            this.searchDataGrid.Name = "searchDataGrid";
            this.searchDataGrid.RowTemplate.Height = 28;
            this.searchDataGrid.Size = new System.Drawing.Size(819, 271);
            this.searchDataGrid.TabIndex = 10;
            this.searchDataGrid.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.searchDataGrid_EditingControlShowing);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnCreateIndex
            // 
            this.btnCreateIndex.Location = new System.Drawing.Point(130, 7);
            this.btnCreateIndex.Margin = new System.Windows.Forms.Padding(2);
            this.btnCreateIndex.Name = "btnCreateIndex";
            this.btnCreateIndex.Size = new System.Drawing.Size(93, 26);
            this.btnCreateIndex.TabIndex = 11;
            this.btnCreateIndex.Text = "Create Index";
            this.btnCreateIndex.UseVisualStyleBackColor = true;
            this.btnCreateIndex.Click += new System.EventHandler(this.btnCreateIndex_Click);
            // 
            // btnImportData
            // 
            this.btnImportData.Location = new System.Drawing.Point(226, 7);
            this.btnImportData.Margin = new System.Windows.Forms.Padding(2);
            this.btnImportData.Name = "btnImportData";
            this.btnImportData.Size = new System.Drawing.Size(149, 26);
            this.btnImportData.TabIndex = 12;
            this.btnImportData.Text = "Upload Data To Search";
            this.btnImportData.UseVisualStyleBackColor = true;
            this.btnImportData.Click += new System.EventHandler(this.btnImportData_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(638, 8);
            this.txtFileName.Margin = new System.Windows.Forms.Padding(2);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.ReadOnly = true;
            this.txtFileName.Size = new System.Drawing.Size(189, 20);
            this.txtFileName.TabIndex = 3;
            this.txtFileName.Visible = false;
            // 
            // txtIndexName
            // 
            this.txtIndexName.Location = new System.Drawing.Point(683, 38);
            this.txtIndexName.Margin = new System.Windows.Forms.Padding(2);
            this.txtIndexName.Name = "txtIndexName";
            this.txtIndexName.Size = new System.Drawing.Size(144, 20);
            this.txtIndexName.TabIndex = 14;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(615, 41);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Index Name";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 267);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(142, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Azure Search Index Schema";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 565);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(843, 22);
            this.statusStrip1.TabIndex = 16;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // txtBatchSize
            // 
            this.txtBatchSize.Location = new System.Drawing.Point(469, 9);
            this.txtBatchSize.Margin = new System.Windows.Forms.Padding(2);
            this.txtBatchSize.Name = "txtBatchSize";
            this.txtBatchSize.Size = new System.Drawing.Size(131, 20);
            this.txtBatchSize.TabIndex = 18;
            this.txtBatchSize.Text = "1000";
            this.txtBatchSize.TextChanged += new System.EventHandler(this.txtBatchSize_TextChanged);
            this.txtBatchSize.Leave += new System.EventHandler(this.txtBatchSize_Leave);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(385, 11);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "Batch Size";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(843, 587);
            this.Controls.Add(this.txtBatchSize);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtIndexName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnImportData);
            this.Controls.Add(this.btnCreateIndex);
            this.Controls.Add(this.searchDataGrid);
            this.Controls.Add(this.csvDataGrid);
            this.Controls.Add(this.btnChooseFile);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtApiKey);
            this.Controls.Add(this.txtSearchService);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmMain";
            this.Text = "Azure Search CSV Importer";
            ((System.ComponentModel.ISupportInitialize)(this.csvDataGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.searchDataGrid)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSearchService;
        private System.Windows.Forms.TextBox txtApiKey;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnChooseFile;
        private System.Windows.Forms.DataGridView csvDataGrid;
        private System.Windows.Forms.DataGridView searchDataGrid;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnCreateIndex;
        private System.Windows.Forms.Button btnImportData;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.TextBox txtIndexName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.TextBox txtBatchSize;
        private System.Windows.Forms.Label label6;
    }
}

