namespace JiraInteraction
{
    partial class FrmPrmtrInpt
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
            this.label1 = new System.Windows.Forms.Label();
            this.lstbxPrjcts = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dtUpdtDt = new System.Windows.Forms.DateTimePicker();
            this.chbxUpdtMsp = new System.Windows.Forms.CheckBox();
            this.chbxScrpTrllo = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbxXlsFlNm = new System.Windows.Forms.TextBox();
            this.btnStrt = new System.Windows.Forms.Button();
            this.btnCncl = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.875F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(451, 42);
            this.label1.TabIndex = 0;
            this.label1.Text = "MSP Update From Trello";
            // 
            // lstbxPrjcts
            // 
            this.lstbxPrjcts.FormattingEnabled = true;
            this.lstbxPrjcts.ItemHeight = 25;
            this.lstbxPrjcts.Location = new System.Drawing.Point(206, 76);
            this.lstbxPrjcts.Name = "lstbxPrjcts";
            this.lstbxPrjcts.Size = new System.Drawing.Size(433, 54);
            this.lstbxPrjcts.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(25, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 33);
            this.label2.TabIndex = 2;
            this.label2.Text = "Project";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(25, 189);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(177, 33);
            this.label3.TabIndex = 3;
            this.label3.Text = "Update Date";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // dtUpdtDt
            // 
            this.dtUpdtDt.Location = new System.Drawing.Point(206, 189);
            this.dtUpdtDt.Name = "dtUpdtDt";
            this.dtUpdtDt.Size = new System.Drawing.Size(380, 31);
            this.dtUpdtDt.TabIndex = 4;
            // 
            // chbxUpdtMsp
            // 
            this.chbxUpdtMsp.AutoSize = true;
            this.chbxUpdtMsp.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chbxUpdtMsp.Location = new System.Drawing.Point(434, 295);
            this.chbxUpdtMsp.Name = "chbxUpdtMsp";
            this.chbxUpdtMsp.Size = new System.Drawing.Size(214, 35);
            this.chbxUpdtMsp.TabIndex = 6;
            this.chbxUpdtMsp.Text = "Update MSP?";
            this.chbxUpdtMsp.UseVisualStyleBackColor = true;
            // 
            // chbxScrpTrllo
            // 
            this.chbxScrpTrllo.AutoSize = true;
            this.chbxScrpTrllo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chbxScrpTrllo.Location = new System.Drawing.Point(31, 295);
            this.chbxScrpTrllo.Name = "chbxScrpTrllo";
            this.chbxScrpTrllo.Size = new System.Drawing.Size(268, 35);
            this.chbxScrpTrllo.TabIndex = 7;
            this.chbxScrpTrllo.Text = "Scrape from Trello";
            this.chbxScrpTrllo.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(25, 401);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(175, 33);
            this.label4.TabIndex = 8;
            this.label4.Text = "Excel output";
            this.label4.Click += new System.EventHandler(this.label4_Click_1);
            // 
            // tbxXlsFlNm
            // 
            this.tbxXlsFlNm.Location = new System.Drawing.Point(206, 405);
            this.tbxXlsFlNm.Name = "tbxXlsFlNm";
            this.tbxXlsFlNm.Size = new System.Drawing.Size(563, 31);
            this.tbxXlsFlNm.TabIndex = 9;
            // 
            // btnStrt
            // 
            this.btnStrt.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStrt.Location = new System.Drawing.Point(31, 505);
            this.btnStrt.Name = "btnStrt";
            this.btnStrt.Size = new System.Drawing.Size(163, 62);
            this.btnStrt.TabIndex = 10;
            this.btnStrt.Text = "Start";
            this.btnStrt.UseVisualStyleBackColor = true;
            this.btnStrt.Click += new System.EventHandler(this.btnStrt_Click);
            // 
            // btnCncl
            // 
            this.btnCncl.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCncl.Location = new System.Drawing.Point(476, 505);
            this.btnCncl.Name = "btnCncl";
            this.btnCncl.Size = new System.Drawing.Size(163, 62);
            this.btnCncl.TabIndex = 11;
            this.btnCncl.Text = "Cancel";
            this.btnCncl.UseVisualStyleBackColor = true;
            // 
            // FrmPrmtrInpt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(838, 621);
            this.Controls.Add(this.btnCncl);
            this.Controls.Add(this.btnStrt);
            this.Controls.Add(this.tbxXlsFlNm);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.chbxScrpTrllo);
            this.Controls.Add(this.chbxUpdtMsp);
            this.Controls.Add(this.dtUpdtDt);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lstbxPrjcts);
            this.Controls.Add(this.label1);
            this.Name = "FrmPrmtrInpt";
            this.Text = "FrmPrmtrInpt";
            this.Load += new System.EventHandler(this.FrmPrmtrInpt_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstbxPrjcts;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtUpdtDt;
        private System.Windows.Forms.CheckBox chbxUpdtMsp;
        private System.Windows.Forms.CheckBox chbxScrpTrllo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbxXlsFlNm;
        private System.Windows.Forms.Button btnStrt;
        private System.Windows.Forms.Button btnCncl;
    }
}