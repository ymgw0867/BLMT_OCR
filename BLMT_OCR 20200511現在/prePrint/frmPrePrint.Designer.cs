namespace BLMT_OCR.prePrint
{
    partial class frmPrePrint
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPrePrint));
            this.cmbArea = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbShimebi = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnPrn = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.linkAllOn = new System.Windows.Forms.LinkLabel();
            this.linkAllOff = new System.Windows.Forms.LinkLabel();
            this.lblCnt = new System.Windows.Forms.Label();
            this.txtEyear = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtSyear = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtEmonth = new System.Windows.Forms.TextBox();
            this.txtSmonth = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbArea
            // 
            this.cmbArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbArea.FormattingEnabled = true;
            this.cmbArea.Location = new System.Drawing.Point(105, 18);
            this.cmbArea.Name = "cmbArea";
            this.cmbArea.Size = new System.Drawing.Size(194, 21);
            this.cmbArea.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.Lavender;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label5.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(13, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(91, 21);
            this.label5.TabIndex = 9;
            this.label5.Text = "エリア指定";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.Lavender;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(303, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(49, 21);
            this.label7.TabIndex = 12;
            this.label7.Text = "締日";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbShimebi
            // 
            this.cmbShimebi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShimebi.FormattingEnabled = true;
            this.cmbShimebi.Location = new System.Drawing.Point(353, 18);
            this.cmbShimebi.Name = "cmbShimebi";
            this.cmbShimebi.Size = new System.Drawing.Size(76, 21);
            this.cmbShimebi.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(13, 47);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(1066, 547);
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.TabIndex = 11;
            // 
            // btnPrn
            // 
            this.btnPrn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrn.BackColor = System.Drawing.Color.Lavender;
            this.btnPrn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrn.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPrn.Location = new System.Drawing.Point(923, 18);
            this.btnPrn.Name = "btnPrn";
            this.btnPrn.Size = new System.Drawing.Size(76, 21);
            this.btnPrn.TabIndex = 7;
            this.btnPrn.Text = "印刷(&P)";
            this.btnPrn.UseVisualStyleBackColor = false;
            this.btnPrn.Click += new System.EventHandler(this.btnPrn_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRtn.BackColor = System.Drawing.Color.Lavender;
            this.btnRtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRtn.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(1002, 18);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(76, 21);
            this.btnRtn.TabIndex = 8;
            this.btnRtn.Text = "戻る(&E)";
            this.btnRtn.UseVisualStyleBackColor = false;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Lavender;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(436, 18);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(74, 21);
            this.button1.TabIndex = 2;
            this.button1.Text = "表示(&C)";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // linkAllOn
            // 
            this.linkAllOn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linkAllOn.AutoSize = true;
            this.linkAllOn.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkAllOn.Location = new System.Drawing.Point(17, 612);
            this.linkAllOn.Name = "linkAllOn";
            this.linkAllOn.Size = new System.Drawing.Size(161, 12);
            this.linkAllOn.TabIndex = 9;
            this.linkAllOn.TabStop = true;
            this.linkAllOn.Text = "全てのチェックをオンにする";
            this.linkAllOn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkAllOn_LinkClicked);
            // 
            // linkAllOff
            // 
            this.linkAllOff.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linkAllOff.AutoSize = true;
            this.linkAllOff.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkAllOff.Location = new System.Drawing.Point(204, 612);
            this.linkAllOff.Name = "linkAllOff";
            this.linkAllOff.Size = new System.Drawing.Size(161, 12);
            this.linkAllOff.TabIndex = 10;
            this.linkAllOff.TabStop = true;
            this.linkAllOff.Text = "全てのチェックをオフにする";
            this.linkAllOff.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lblCnt
            // 
            this.lblCnt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblCnt.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblCnt.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblCnt.Location = new System.Drawing.Point(920, 612);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(159, 12);
            this.lblCnt.TabIndex = 22;
            this.lblCnt.Text = "label9";
            // 
            // txtEyear
            // 
            this.txtEyear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtEyear.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtEyear.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtEyear.Location = new System.Drawing.Point(796, 18);
            this.txtEyear.MaxLength = 4;
            this.txtEyear.Name = "txtEyear";
            this.txtEyear.Size = new System.Drawing.Size(45, 21);
            this.txtEyear.TabIndex = 5;
            this.txtEyear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(776, 22);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(21, 13);
            this.label8.TabIndex = 32;
            this.label8.Text = "～";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtSyear
            // 
            this.txtSyear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSyear.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSyear.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtSyear.Location = new System.Drawing.Point(652, 18);
            this.txtSyear.MaxLength = 4;
            this.txtSyear.Name = "txtSyear";
            this.txtSyear.Size = new System.Drawing.Size(45, 21);
            this.txtSyear.TabIndex = 3;
            this.txtSyear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.BackColor = System.Drawing.Color.Lavender;
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label6.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(576, 18);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 21);
            this.label6.TabIndex = 31;
            this.label6.Text = "印刷期間";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtEmonth
            // 
            this.txtEmonth.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtEmonth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtEmonth.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtEmonth.Location = new System.Drawing.Point(863, 18);
            this.txtEmonth.MaxLength = 2;
            this.txtEmonth.Name = "txtEmonth";
            this.txtEmonth.Size = new System.Drawing.Size(32, 21);
            this.txtEmonth.TabIndex = 6;
            this.txtEmonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtSmonth
            // 
            this.txtSmonth.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSmonth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSmonth.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtSmonth.Location = new System.Drawing.Point(719, 18);
            this.txtSmonth.MaxLength = 2;
            this.txtSmonth.Name = "txtSmonth";
            this.txtSmonth.Size = new System.Drawing.Size(32, 21);
            this.txtSmonth.TabIndex = 4;
            this.txtSmonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.BackColor = System.Drawing.Color.Lavender;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(894, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(24, 21);
            this.label3.TabIndex = 30;
            this.label3.Text = "月";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.BackColor = System.Drawing.Color.Lavender;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(840, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(24, 21);
            this.label4.TabIndex = 27;
            this.label4.Text = "年";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.BackColor = System.Drawing.Color.Lavender;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(750, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(24, 21);
            this.label2.TabIndex = 26;
            this.label2.Text = "月";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.BackColor = System.Drawing.Color.Lavender;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(696, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(24, 21);
            this.label1.TabIndex = 24;
            this.label1.Text = "年";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmPrePrint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1090, 633);
            this.Controls.Add(this.txtEyear);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtSyear);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtEmonth);
            this.Controls.Add(this.txtSmonth);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.linkAllOff);
            this.Controls.Add(this.linkAllOn);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cmbArea);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnPrn);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.cmbShimebi);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Font = new System.Drawing.Font("ＭＳ ゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmPrePrint";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "出勤簿 兼 就業時間管理表印刷";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmPrePrint_FormClosing);
            this.Load += new System.EventHandler(this.frmPrePrint_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbArea;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbShimebi;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnPrn;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.LinkLabel linkAllOn;
        private System.Windows.Forms.LinkLabel linkAllOff;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.TextBox txtEyear;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtSyear;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtEmonth;
        private System.Windows.Forms.TextBox txtSmonth;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}