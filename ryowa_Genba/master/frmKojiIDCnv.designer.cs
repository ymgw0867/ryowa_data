namespace ryowa_Genba.master
{
    partial class frmKojiIDCnv
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmKojiIDCnv));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFrmID = new System.Windows.Forms.TextBox();
            this.txtToID = new System.Windows.Forms.TextBox();
            this.lblFrmName = new System.Windows.Forms.Label();
            this.btnok = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.lblToName = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(22, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "変換元ID";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(22, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 19);
            this.label2.TabIndex = 1;
            this.label2.Text = "変換後ID";
            // 
            // txtFrmID
            // 
            this.txtFrmID.Font = new System.Drawing.Font("Meiryo UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFrmID.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFrmID.Location = new System.Drawing.Point(99, 33);
            this.txtFrmID.MaxLength = 6;
            this.txtFrmID.Name = "txtFrmID";
            this.txtFrmID.Size = new System.Drawing.Size(65, 26);
            this.txtFrmID.TabIndex = 0;
            this.txtFrmID.TextChanged += new System.EventHandler(this.txtFrmID_TextChanged);
            this.txtFrmID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // txtToID
            // 
            this.txtToID.Font = new System.Drawing.Font("Meiryo UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtToID.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtToID.Location = new System.Drawing.Point(99, 82);
            this.txtToID.MaxLength = 6;
            this.txtToID.Name = "txtToID";
            this.txtToID.Size = new System.Drawing.Size(65, 26);
            this.txtToID.TabIndex = 1;
            this.txtToID.TextChanged += new System.EventHandler(this.txtToID_TextChanged);
            // 
            // lblFrmName
            // 
            this.lblFrmName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFrmName.Font = new System.Drawing.Font("Meiryo UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblFrmName.Location = new System.Drawing.Point(170, 33);
            this.lblFrmName.Name = "lblFrmName";
            this.lblFrmName.Size = new System.Drawing.Size(372, 26);
            this.lblFrmName.TabIndex = 4;
            this.lblFrmName.Text = "label3";
            // 
            // btnok
            // 
            this.btnok.Location = new System.Drawing.Point(360, 135);
            this.btnok.Name = "btnok";
            this.btnok.Size = new System.Drawing.Size(88, 32);
            this.btnok.TabIndex = 2;
            this.btnok.Text = "実行(&D)";
            this.btnok.UseVisualStyleBackColor = true;
            this.btnok.Click += new System.EventHandler(this.btnok_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Location = new System.Drawing.Point(454, 135);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(88, 32);
            this.btnRtn.TabIndex = 3;
            this.btnRtn.Text = "戻る(&E)";
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // lblToName
            // 
            this.lblToName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblToName.Font = new System.Drawing.Font("Meiryo UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblToName.Location = new System.Drawing.Point(170, 82);
            this.lblToName.Name = "lblToName";
            this.lblToName.Size = new System.Drawing.Size(372, 26);
            this.lblToName.TabIndex = 5;
            this.lblToName.Text = "label4";
            // 
            // frmKojiIDCnv
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(561, 174);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnok);
            this.Controls.Add(this.lblToName);
            this.Controls.Add(this.lblFrmName);
            this.Controls.Add(this.txtToID);
            this.Controls.Add(this.txtFrmID);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmKojiIDCnv";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "工事マスターID一括変換";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmKojiIDCnv_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFrmID;
        private System.Windows.Forms.TextBox txtToID;
        private System.Windows.Forms.Label lblFrmName;
        private System.Windows.Forms.Button btnok;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.Label lblToName;
    }
}