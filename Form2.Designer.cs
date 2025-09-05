namespace ExcelToDB
{
    partial class Form2
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
            this.lstTables = new System.Windows.Forms.ListBox();
            this.dgvTableData = new System.Windows.Forms.DataGridView();
            this.cmbDatabases = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnDownloadTemplate = new System.Windows.Forms.Button();
            this.btnExportToExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTableData)).BeginInit();
            this.SuspendLayout();
            // 
            // lstTables
            // 
            this.lstTables.FormattingEnabled = true;
            this.lstTables.ItemHeight = 16;
            this.lstTables.Location = new System.Drawing.Point(55, 32);
            this.lstTables.Name = "lstTables";
            this.lstTables.Size = new System.Drawing.Size(136, 372);
            this.lstTables.TabIndex = 0;
            this.lstTables.SelectedIndexChanged += new System.EventHandler(this.lstTables_SelectedIndexChanged);
            // 
            // dgvTableData
            // 
            this.dgvTableData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTableData.Location = new System.Drawing.Point(216, 32);
            this.dgvTableData.Name = "dgvTableData";
            this.dgvTableData.RowHeadersWidth = 51;
            this.dgvTableData.RowTemplate.Height = 24;
            this.dgvTableData.Size = new System.Drawing.Size(739, 376);
            this.dgvTableData.TabIndex = 1;
            // 
            // cmbDatabases
            // 
            this.cmbDatabases.FormattingEnabled = true;
            this.cmbDatabases.Location = new System.Drawing.Point(55, 2);
            this.cmbDatabases.Name = "cmbDatabases";
            this.cmbDatabases.Size = new System.Drawing.Size(121, 24);
            this.cmbDatabases.TabIndex = 2;
            this.cmbDatabases.SelectedIndexChanged += new System.EventHandler(this.cmbDatabases_SelectedIndexChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(689, 411);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(99, 27);
            this.button1.TabIndex = 3;
            this.button1.Text = "Excel Ekle";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnImportExcel_Click);
            // 
            // btnDownloadTemplate
            // 
            this.btnDownloadTemplate.Location = new System.Drawing.Point(582, 414);
            this.btnDownloadTemplate.Name = "btnDownloadTemplate";
            this.btnDownloadTemplate.Size = new System.Drawing.Size(101, 24);
            this.btnDownloadTemplate.TabIndex = 4;
            this.btnDownloadTemplate.Text = "Örnek Şablon";
            this.btnDownloadTemplate.UseVisualStyleBackColor = true;
            this.btnDownloadTemplate.Click += new System.EventHandler(this.btnDownloadTemplate_Click);
            // 
            // btnExportToExcel
            // 
            this.btnExportToExcel.Location = new System.Drawing.Point(414, 414);
            this.btnExportToExcel.Name = "btnExportToExcel";
            this.btnExportToExcel.Size = new System.Drawing.Size(162, 24);
            this.btnExportToExcel.TabIndex = 5;
            this.btnExportToExcel.Text = "Excel olarak dışa aktar";
            this.btnExportToExcel.UseVisualStyleBackColor = true;
            this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1003, 504);
            this.Controls.Add(this.btnExportToExcel);
            this.Controls.Add(this.btnDownloadTemplate);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cmbDatabases);
            this.Controls.Add(this.dgvTableData);
            this.Controls.Add(this.lstTables);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTableData)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lstTables;
        private System.Windows.Forms.DataGridView dgvTableData;
        private System.Windows.Forms.ComboBox cmbDatabases;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnDownloadTemplate;
        private System.Windows.Forms.Button btnExportToExcel;
    }
}