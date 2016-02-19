namespace ExcelToAccess
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
            this.btnReadPerson = new System.Windows.Forms.Button();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.dsInfo = new System.Data.DataSet();
            this.dgShow = new System.Windows.Forms.DataGridView();
            this.btnReadMedication = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dsInfo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgShow)).BeginInit();
            this.SuspendLayout();
            // 
            // btnReadPerson
            // 
            this.btnReadPerson.Location = new System.Drawing.Point(357, 33);
            this.btnReadPerson.Name = "btnReadPerson";
            this.btnReadPerson.Size = new System.Drawing.Size(112, 21);
            this.btnReadPerson.TabIndex = 0;
            this.btnReadPerson.Text = "Read Person File";
            this.btnReadPerson.UseVisualStyleBackColor = true;
            this.btnReadPerson.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(39, 33);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtFileName.Size = new System.Drawing.Size(301, 21);
            this.txtFileName.TabIndex = 1;
            this.txtFileName.Text = "F:\\HW\\cs492\\patientSamll.xls";
            this.txtFileName.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // dsInfo
            // 
            this.dsInfo.DataSetName = "NewDataSet";
            // 
            // dgShow
            // 
            this.dgShow.AllowUserToAddRows = false;
            this.dgShow.AllowUserToDeleteRows = false;
            this.dgShow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgShow.Location = new System.Drawing.Point(12, 60);
            this.dgShow.Name = "dgShow";
            this.dgShow.ReadOnly = true;
            this.dgShow.Size = new System.Drawing.Size(547, 204);
            this.dgShow.TabIndex = 2;
            // 
            // btnReadMedication
            // 
            this.btnReadMedication.Location = new System.Drawing.Point(357, 12);
            this.btnReadMedication.Name = "btnReadMedication";
            this.btnReadMedication.Size = new System.Drawing.Size(112, 21);
            this.btnReadMedication.TabIndex = 3;
            this.btnReadMedication.Text = "Read Med File";
            this.btnReadMedication.UseVisualStyleBackColor = true;
            this.btnReadMedication.Click += new System.EventHandler(this.btnReadMedication_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(571, 275);
            this.Controls.Add(this.btnReadMedication);
            this.Controls.Add(this.dgShow);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.btnReadPerson);
            this.Name = "Form1";
            this.Text = "Excel To Access Converter";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dsInfo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgShow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnReadPerson;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Data.DataSet dsInfo;
        private System.Windows.Forms.DataGridView dgShow;
        private System.Windows.Forms.Button btnReadMedication;
    }
}

