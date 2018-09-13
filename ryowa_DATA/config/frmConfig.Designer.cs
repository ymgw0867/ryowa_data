namespace ryowa_DATA.config
{
    partial class frmConfig
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmConfig));
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.txtPath2 = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            this.txtDataSpan = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtJyosetsu = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtTokkinmu1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTokkinmu2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtYakan = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtShokumu = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(468, 321);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(94, 34);
            this.button2.TabIndex = 5;
            this.button2.Text = "登録";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button3.Location = new System.Drawing.Point(568, 321);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(94, 34);
            this.button3.TabIndex = 6;
            this.button3.Text = "終了";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtPath2
            // 
            this.txtPath2.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtPath2.Location = new System.Drawing.Point(36, 262);
            this.txtPath2.Name = "txtPath2";
            this.txtPath2.Size = new System.Drawing.Size(551, 27);
            this.txtPath2.TabIndex = 3;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(593, 262);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(69, 28);
            this.button4.TabIndex = 1;
            this.button4.Text = "参照...";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // txtDataSpan
            // 
            this.txtDataSpan.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtDataSpan.Location = new System.Drawing.Point(227, 25);
            this.txtDataSpan.Name = "txtDataSpan";
            this.txtDataSpan.Size = new System.Drawing.Size(142, 27);
            this.txtDataSpan.TabIndex = 3;
            this.txtDataSpan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtDataSpan.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 234);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(241, 19);
            this.label1.TabIndex = 7;
            this.label1.Text = "給与大臣用CSV出力先フォルダパス：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(43, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(148, 19);
            this.label2.TabIndex = 8;
            this.label2.Text = "勤怠データ保存月数：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtJyosetsu
            // 
            this.txtJyosetsu.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtJyosetsu.Location = new System.Drawing.Point(227, 58);
            this.txtJyosetsu.Name = "txtJyosetsu";
            this.txtJyosetsu.Size = new System.Drawing.Size(142, 27);
            this.txtJyosetsu.TabIndex = 9;
            this.txtJyosetsu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtJyosetsu.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(77, 61);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(114, 19);
            this.label3.TabIndex = 10;
            this.label3.Text = "除雪当番単価：";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtTokkinmu1
            // 
            this.txtTokkinmu1.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTokkinmu1.Location = new System.Drawing.Point(227, 91);
            this.txtTokkinmu1.Name = "txtTokkinmu1";
            this.txtTokkinmu1.Size = new System.Drawing.Size(142, 27);
            this.txtTokkinmu1.TabIndex = 11;
            this.txtTokkinmu1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtTokkinmu1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(159, 19);
            this.label4.TabIndex = 12;
            this.label4.Text = "特殊勤務手当単価１：";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtTokkinmu2
            // 
            this.txtTokkinmu2.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTokkinmu2.Location = new System.Drawing.Point(227, 124);
            this.txtTokkinmu2.Name = "txtTokkinmu2";
            this.txtTokkinmu2.Size = new System.Drawing.Size(142, 27);
            this.txtTokkinmu2.TabIndex = 13;
            this.txtTokkinmu2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtTokkinmu2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(32, 127);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(159, 19);
            this.label5.TabIndex = 14;
            this.label5.Text = "特殊勤務手当単価２：";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtYakan
            // 
            this.txtYakan.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtYakan.Location = new System.Drawing.Point(227, 157);
            this.txtYakan.Name = "txtYakan";
            this.txtYakan.Size = new System.Drawing.Size(142, 27);
            this.txtYakan.TabIndex = 15;
            this.txtYakan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtYakan.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(77, 160);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(114, 19);
            this.label6.TabIndex = 16;
            this.label6.Text = "夜間手当単価：";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtShokumu
            // 
            this.txtShokumu.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtShokumu.Location = new System.Drawing.Point(227, 190);
            this.txtShokumu.Name = "txtShokumu";
            this.txtShokumu.Size = new System.Drawing.Size(142, 27);
            this.txtShokumu.TabIndex = 17;
            this.txtShokumu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtShokumu.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(77, 193);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(114, 19);
            this.label7.TabIndex = 18;
            this.label7.Text = "職務手当単価：";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(413, 94);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(218, 19);
            this.label8.TabIndex = 19;
            this.label8.Text = "※特殊勤務が4.0Ｈ/日未満のとき";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(413, 127);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(218, 19);
            this.label9.TabIndex = 20;
            this.label9.Text = "※特殊勤務が4.0Ｈ/日以上のとき";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(683, 371);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtShokumu);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtYakan);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtTokkinmu2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtTokkinmu1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtJyosetsu);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtDataSpan);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.txtPath2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Font = new System.Drawing.Font("Meiryo UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "frmConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "環境設定";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmConfig_FormClosing);
            this.Load += new System.EventHandler(this.frmConfig_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox txtPath2;
        private System.Windows.Forms.TextBox txtDataSpan;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtJyosetsu;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtTokkinmu1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtTokkinmu2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtYakan;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtShokumu;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
    }
}