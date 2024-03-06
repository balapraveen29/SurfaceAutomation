namespace SurfaceAutomation
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
            this.btnActivateWindow = new System.Windows.Forms.Button();
            this.btnExcelCreated = new System.Windows.Forms.Button();
            this.btnTxt2Csv = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnActivateWindow
            // 
            this.btnActivateWindow.Location = new System.Drawing.Point(38, 31);
            this.btnActivateWindow.Name = "btnActivateWindow";
            this.btnActivateWindow.Size = new System.Drawing.Size(75, 46);
            this.btnActivateWindow.TabIndex = 0;
            this.btnActivateWindow.Text = "Activate Window";
            this.btnActivateWindow.UseVisualStyleBackColor = true;
            this.btnActivateWindow.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnExcelCreated
            // 
            this.btnExcelCreated.Location = new System.Drawing.Point(38, 98);
            this.btnExcelCreated.Name = "btnExcelCreated";
            this.btnExcelCreated.Size = new System.Drawing.Size(75, 46);
            this.btnExcelCreated.TabIndex = 1;
            this.btnExcelCreated.Text = "Is Excel Created";
            this.btnExcelCreated.UseVisualStyleBackColor = true;
            this.btnExcelCreated.Click += new System.EventHandler(this.btnExcelCreated_Click);
            // 
            // btnTxt2Csv
            // 
            this.btnTxt2Csv.Location = new System.Drawing.Point(38, 166);
            this.btnTxt2Csv.Name = "btnTxt2Csv";
            this.btnTxt2Csv.Size = new System.Drawing.Size(75, 46);
            this.btnTxt2Csv.TabIndex = 2;
            this.btnTxt2Csv.Text = "Convert Text to CSV";
            this.btnTxt2Csv.UseVisualStyleBackColor = true;
            this.btnTxt2Csv.Click += new System.EventHandler(this.btnTxt2Csv_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnTxt2Csv);
            this.Controls.Add(this.btnExcelCreated);
            this.Controls.Add(this.btnActivateWindow);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnActivateWindow;
        private System.Windows.Forms.Button btnExcelCreated;
        private System.Windows.Forms.Button btnTxt2Csv;
    }
}

