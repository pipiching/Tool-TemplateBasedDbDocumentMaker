namespace DbDocumentMaker
{
    partial class QueryForm
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
            this.Label5 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.txtObjectName = new System.Windows.Forms.TextBox();
            this.dgvSchemaColumns = new System.Windows.Forms.DataGridView();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.txtQuery = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSchemaColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(8, 17);
            this.Label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(61, 12);
            this.Label5.TabIndex = 25;
            this.Label5.Text = "Query Text:";
            // 
            // Label3
            // 
            this.Label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(6, 290);
            this.Label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(190, 12);
            this.Label3.TabIndex = 24;
            this.Label3.Text = "Generated Results Schema POCO Class:";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(8, 129);
            this.Label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(70, 12);
            this.Label2.TabIndex = 23;
            this.Label2.Text = "SchemaTable:";
            // 
            // Label4
            // 
            this.Label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(6, 495);
            this.Label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(62, 12);
            this.Label4.TabIndex = 22;
            this.Label4.Text = "Class Name:";
            // 
            // txtObjectName
            // 
            this.txtObjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtObjectName.Location = new System.Drawing.Point(75, 493);
            this.txtObjectName.Margin = new System.Windows.Forms.Padding(2);
            this.txtObjectName.Name = "txtObjectName";
            this.txtObjectName.Size = new System.Drawing.Size(400, 22);
            this.txtObjectName.TabIndex = 21;
            this.txtObjectName.Text = "POCO";
            // 
            // dgvSchemaColumns
            // 
            this.dgvSchemaColumns.AllowUserToAddRows = false;
            this.dgvSchemaColumns.AllowUserToDeleteRows = false;
            this.dgvSchemaColumns.AllowUserToOrderColumns = true;
            this.dgvSchemaColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSchemaColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSchemaColumns.Location = new System.Drawing.Point(8, 144);
            this.dgvSchemaColumns.Margin = new System.Windows.Forms.Padding(2);
            this.dgvSchemaColumns.Name = "dgvSchemaColumns";
            this.dgvSchemaColumns.ReadOnly = true;
            this.dgvSchemaColumns.RowTemplate.Height = 24;
            this.dgvSchemaColumns.Size = new System.Drawing.Size(711, 134);
            this.dgvSchemaColumns.TabIndex = 18;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGenerate.Location = new System.Drawing.Point(635, 490);
            this.btnGenerate.Margin = new System.Windows.Forms.Padding(2);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(83, 22);
            this.btnGenerate.TabIndex = 20;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtCode
            // 
            this.txtCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCode.Location = new System.Drawing.Point(8, 304);
            this.txtCode.Margin = new System.Windows.Forms.Padding(2);
            this.txtCode.Multiline = true;
            this.txtCode.Name = "txtCode";
            this.txtCode.ReadOnly = true;
            this.txtCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtCode.Size = new System.Drawing.Size(712, 169);
            this.txtCode.TabIndex = 19;
            // 
            // txtQuery
            // 
            this.txtQuery.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtQuery.Location = new System.Drawing.Point(8, 31);
            this.txtQuery.Margin = new System.Windows.Forms.Padding(2);
            this.txtQuery.Multiline = true;
            this.txtQuery.Name = "txtQuery";
            this.txtQuery.Size = new System.Drawing.Size(708, 89);
            this.txtQuery.TabIndex = 17;
            // 
            // QueryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(727, 532);
            this.Controls.Add(this.Label5);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label4);
            this.Controls.Add(this.txtObjectName);
            this.Controls.Add(this.dgvSchemaColumns);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.txtQuery);
            this.Name = "QueryForm";
            this.Text = "QueryForm";
            ((System.ComponentModel.ISupportInitialize)(this.dgvSchemaColumns)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.TextBox txtObjectName;
        internal System.Windows.Forms.DataGridView dgvSchemaColumns;
        internal System.Windows.Forms.Button btnGenerate;
        internal System.Windows.Forms.TextBox txtCode;
        internal System.Windows.Forms.TextBox txtQuery;
    }
}