namespace AVDApplication
{
    partial class GEWConfirmExport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GEWConfirmExport));
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnGEWCancel = new System.Windows.Forms.Button();
            this.btnGEWOK = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdFre = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.rdTran = new System.Windows.Forms.RadioButton();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnGEWCancel);
            this.groupBox2.Controls.Add(this.btnGEWOK);
            this.groupBox2.Location = new System.Drawing.Point(209, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(111, 133);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            // 
            // btnGEWCancel
            // 
            this.btnGEWCancel.Location = new System.Drawing.Point(16, 80);
            this.btnGEWCancel.Name = "btnGEWCancel";
            this.btnGEWCancel.Size = new System.Drawing.Size(75, 23);
            this.btnGEWCancel.TabIndex = 1;
            this.btnGEWCancel.Text = "Cancel";
            this.btnGEWCancel.UseVisualStyleBackColor = true;
            this.btnGEWCancel.Click += new System.EventHandler(this.btnGEWCancel_Click);
            // 
            // btnGEWOK
            // 
            this.btnGEWOK.Location = new System.Drawing.Point(16, 44);
            this.btnGEWOK.Name = "btnGEWOK";
            this.btnGEWOK.Size = new System.Drawing.Size(75, 23);
            this.btnGEWOK.TabIndex = 0;
            this.btnGEWOK.Text = "Export";
            this.btnGEWOK.UseVisualStyleBackColor = true;
            this.btnGEWOK.Click += new System.EventHandler(this.btnGEWOK_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdTran);
            this.groupBox1.Controls.Add(this.rdFre);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(191, 133);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // rdFre
            // 
            this.rdFre.AutoSize = true;
            this.rdFre.Location = new System.Drawing.Point(14, 80);
            this.rdFre.Name = "rdFre";
            this.rdFre.Size = new System.Drawing.Size(108, 17);
            this.rdFre.TabIndex = 2;
            this.rdFre.TabStop = true;
            this.rdFre.Text = "Export Frequency";
            this.rdFre.UseVisualStyleBackColor = true;
            this.rdFre.CheckedChanged += new System.EventHandler(this.rdFre_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select type of Export";
            // 
            // rdTran
            // 
            this.rdTran.AutoSize = true;
            this.rdTran.Location = new System.Drawing.Point(14, 50);
            this.rdTran.Name = "rdTran";
            this.rdTran.Size = new System.Drawing.Size(110, 17);
            this.rdTran.TabIndex = 3;
            this.rdTran.TabStop = true;
            this.rdTran.Text = "Export Transmitter";
            this.rdTran.UseVisualStyleBackColor = true;
            this.rdTran.CheckedChanged += new System.EventHandler(this.rdTran_CheckedChanged);
            // 
            // GEWConfirmExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(337, 165);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "GEWConfirmExport";
            this.Text = "GEWConfirmExport";
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnGEWCancel;
        private System.Windows.Forms.Button btnGEWOK;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdFre;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rdTran;
    }
}