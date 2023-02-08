namespace DbDocumentMaker
{
    partial class FrmNewConnection
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
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtConnStr = new System.Windows.Forms.TextBox();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnConnTest = new System.Windows.Forms.Button();
            this.lbConnName = new System.Windows.Forms.Label();
            this.lbConnStr = new System.Windows.Forms.Label();
            this.txtSystemDescription = new System.Windows.Forms.TextBox();
            this.txtSystemName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(12, 29);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(346, 22);
            this.txtName.TabIndex = 0;
            // 
            // txtConnStr
            // 
            this.txtConnStr.Location = new System.Drawing.Point(12, 75);
            this.txtConnStr.Multiline = true;
            this.txtConnStr.Name = "txtConnStr";
            this.txtConnStr.Size = new System.Drawing.Size(346, 55);
            this.txtConnStr.TabIndex = 1;
            // 
            // btnCreate
            // 
            this.btnCreate.BackColor = System.Drawing.Color.LightGreen;
            this.btnCreate.Location = new System.Drawing.Point(230, 200);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(108, 39);
            this.btnCreate.TabIndex = 2;
            this.btnCreate.Text = "Create";
            this.btnCreate.UseVisualStyleBackColor = false;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnConnTest
            // 
            this.btnConnTest.Location = new System.Drawing.Point(13, 207);
            this.btnConnTest.Name = "btnConnTest";
            this.btnConnTest.Size = new System.Drawing.Size(103, 25);
            this.btnConnTest.TabIndex = 3;
            this.btnConnTest.Text = "Connection Test ";
            this.btnConnTest.UseVisualStyleBackColor = true;
            this.btnConnTest.Click += new System.EventHandler(this.btnConnTest_Click);
            // 
            // lbConnName
            // 
            this.lbConnName.AutoSize = true;
            this.lbConnName.Location = new System.Drawing.Point(10, 14);
            this.lbConnName.Name = "lbConnName";
            this.lbConnName.Size = new System.Drawing.Size(32, 12);
            this.lbConnName.TabIndex = 4;
            this.lbConnName.Text = "Name";
            // 
            // lbConnStr
            // 
            this.lbConnStr.AutoSize = true;
            this.lbConnStr.Location = new System.Drawing.Point(10, 60);
            this.lbConnStr.Name = "lbConnStr";
            this.lbConnStr.Size = new System.Drawing.Size(80, 12);
            this.lbConnStr.TabIndex = 5;
            this.lbConnStr.Text = "Conntion String";
            // 
            // txtSystemDescription
            // 
            this.txtSystemDescription.Location = new System.Drawing.Point(109, 164);
            this.txtSystemDescription.Name = "txtSystemDescription";
            this.txtSystemDescription.Size = new System.Drawing.Size(186, 22);
            this.txtSystemDescription.TabIndex = 20;
            // 
            // txtSystemName
            // 
            this.txtSystemName.Location = new System.Drawing.Point(109, 136);
            this.txtSystemName.Name = "txtSystemName";
            this.txtSystemName.Size = new System.Drawing.Size(186, 22);
            this.txtSystemName.TabIndex = 19;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 169);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 12);
            this.label2.TabIndex = 18;
            this.label2.Text = "System Description:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 139);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 12);
            this.label1.TabIndex = 17;
            this.label1.Text = "System Name:";
            // 
            // FrmNewConnection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(370, 252);
            this.Controls.Add(this.txtSystemDescription);
            this.Controls.Add(this.txtSystemName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbConnStr);
            this.Controls.Add(this.lbConnName);
            this.Controls.Add(this.btnConnTest);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.txtConnStr);
            this.Controls.Add(this.txtName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "FrmNewConnection";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Create New Connection";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtConnStr;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnConnTest;
        private System.Windows.Forms.Label lbConnName;
        private System.Windows.Forms.Label lbConnStr;
        private System.Windows.Forms.TextBox txtSystemDescription;
        private System.Windows.Forms.TextBox txtSystemName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}