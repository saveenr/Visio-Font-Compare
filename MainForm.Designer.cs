namespace VisioFontCompare
{
    partial class MainForm
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
            this.buttonRun = new System.Windows.Forms.Button();
            this.comboBoxFont1 = new System.Windows.Forms.ComboBox();
            this.comboBoxFont2 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxFont1Bold = new System.Windows.Forms.CheckBox();
            this.checkBoxFont1Italic = new System.Windows.Forms.CheckBox();
            this.checkBoxFont2Italic = new System.Windows.Forms.CheckBox();
            this.checkBoxFont2Bold = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // buttonRun
            // 
            this.buttonRun.Location = new System.Drawing.Point(303, 226);
            this.buttonRun.Name = "buttonRun";
            this.buttonRun.Size = new System.Drawing.Size(75, 23);
            this.buttonRun.TabIndex = 0;
            this.buttonRun.Text = "Run";
            this.buttonRun.UseVisualStyleBackColor = true;
            this.buttonRun.Click += new System.EventHandler(this.buttonRun_Click);
            // 
            // comboBoxFont1
            // 
            this.comboBoxFont1.FormattingEnabled = true;
            this.comboBoxFont1.Location = new System.Drawing.Point(12, 29);
            this.comboBoxFont1.Name = "comboBoxFont1";
            this.comboBoxFont1.Size = new System.Drawing.Size(366, 21);
            this.comboBoxFont1.TabIndex = 1;
            // 
            // comboBoxFont2
            // 
            this.comboBoxFont2.FormattingEnabled = true;
            this.comboBoxFont2.Location = new System.Drawing.Point(12, 166);
            this.comboBoxFont2.Name = "comboBoxFont2";
            this.comboBoxFont2.Size = new System.Drawing.Size(366, 21);
            this.comboBoxFont2.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "First Font";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 150);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Second Font";
            // 
            // checkBoxFont1Bold
            // 
            this.checkBoxFont1Bold.AutoSize = true;
            this.checkBoxFont1Bold.Location = new System.Drawing.Point(16, 57);
            this.checkBoxFont1Bold.Name = "checkBoxFont1Bold";
            this.checkBoxFont1Bold.Size = new System.Drawing.Size(47, 17);
            this.checkBoxFont1Bold.TabIndex = 5;
            this.checkBoxFont1Bold.Text = "Bold";
            this.checkBoxFont1Bold.UseVisualStyleBackColor = true;
            // 
            // checkBoxFont1Italic
            // 
            this.checkBoxFont1Italic.AutoSize = true;
            this.checkBoxFont1Italic.Location = new System.Drawing.Point(16, 80);
            this.checkBoxFont1Italic.Name = "checkBoxFont1Italic";
            this.checkBoxFont1Italic.Size = new System.Drawing.Size(48, 17);
            this.checkBoxFont1Italic.TabIndex = 6;
            this.checkBoxFont1Italic.Text = "Italic";
            this.checkBoxFont1Italic.UseVisualStyleBackColor = true;
            // 
            // checkBoxFont2Italic
            // 
            this.checkBoxFont2Italic.AutoSize = true;
            this.checkBoxFont2Italic.Location = new System.Drawing.Point(16, 216);
            this.checkBoxFont2Italic.Name = "checkBoxFont2Italic";
            this.checkBoxFont2Italic.Size = new System.Drawing.Size(48, 17);
            this.checkBoxFont2Italic.TabIndex = 8;
            this.checkBoxFont2Italic.Text = "Italic";
            this.checkBoxFont2Italic.UseVisualStyleBackColor = true;
            // 
            // checkBoxFont2Bold
            // 
            this.checkBoxFont2Bold.AutoSize = true;
            this.checkBoxFont2Bold.Location = new System.Drawing.Point(16, 193);
            this.checkBoxFont2Bold.Name = "checkBoxFont2Bold";
            this.checkBoxFont2Bold.Size = new System.Drawing.Size(47, 17);
            this.checkBoxFont2Bold.TabIndex = 7;
            this.checkBoxFont2Bold.Text = "Bold";
            this.checkBoxFont2Bold.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(389, 307);
            this.Controls.Add(this.checkBoxFont2Italic);
            this.Controls.Add(this.checkBoxFont2Bold);
            this.Controls.Add(this.checkBoxFont1Italic);
            this.Controls.Add(this.checkBoxFont1Bold);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxFont2);
            this.Controls.Add(this.comboBoxFont1);
            this.Controls.Add(this.buttonRun);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "MainForm";
            this.Text = "VisioFontCompare";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonRun;
        private System.Windows.Forms.ComboBox comboBoxFont1;
        private System.Windows.Forms.ComboBox comboBoxFont2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBoxFont1Bold;
        private System.Windows.Forms.CheckBox checkBoxFont1Italic;
        private System.Windows.Forms.CheckBox checkBoxFont2Italic;
        private System.Windows.Forms.CheckBox checkBoxFont2Bold;
    }
}

