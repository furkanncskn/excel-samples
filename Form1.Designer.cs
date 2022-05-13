namespace Excel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtCellValue = new System.Windows.Forms.TextBox();
            this.BtnCreateExcel = new System.Windows.Forms.Button();
            this.BtnWriteToExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtCellValue
            // 
            this.txtCellValue.Font = new System.Drawing.Font("Segoe UI", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtCellValue.Location = new System.Drawing.Point(12, 122);
            this.txtCellValue.Name = "txtCellValue";
            this.txtCellValue.Size = new System.Drawing.Size(194, 38);
            this.txtCellValue.TabIndex = 0;
            this.txtCellValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnCreateExcel
            // 
            this.BtnCreateExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BtnCreateExcel.Location = new System.Drawing.Point(12, 12);
            this.BtnCreateExcel.Name = "BtnCreateExcel";
            this.BtnCreateExcel.Size = new System.Drawing.Size(193, 49);
            this.BtnCreateExcel.TabIndex = 1;
            this.BtnCreateExcel.Text = "CREATE EXCEL FILE";
            this.BtnCreateExcel.UseVisualStyleBackColor = true;
            this.BtnCreateExcel.Click += new System.EventHandler(this.BtnCreateExcel_Click);
            // 
            // BtnWriteToExcel
            // 
            this.BtnWriteToExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BtnWriteToExcel.Location = new System.Drawing.Point(12, 67);
            this.BtnWriteToExcel.Name = "BtnWriteToExcel";
            this.BtnWriteToExcel.Size = new System.Drawing.Size(193, 49);
            this.BtnWriteToExcel.TabIndex = 2;
            this.BtnWriteToExcel.Text = "WRITE TO EXCEL";
            this.BtnWriteToExcel.UseVisualStyleBackColor = true;
            this.BtnWriteToExcel.Click += new System.EventHandler(this.BtnWriteToExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(215, 170);
            this.Controls.Add(this.BtnWriteToExcel);
            this.Controls.Add(this.BtnCreateExcel);
            this.Controls.Add(this.txtCellValue);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtCellValue;
        private System.Windows.Forms.Button BtnCreateExcel;
        private System.Windows.Forms.Button BtnWriteToExcel;
    }
}
