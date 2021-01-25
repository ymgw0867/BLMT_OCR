namespace BLMT_OCR.OCR
{
    partial class frmPastCorrect
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPastCorrect));
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnMinus = new System.Windows.Forms.Button();
            this.btnPlus = new System.Windows.Forms.Button();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblErrMsg = new System.Windows.Forms.Label();
            this.btnRtn = new System.Windows.Forms.Button();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.chkEdit = new System.Windows.Forms.CheckBox();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.lblEdit = new System.Windows.Forms.Label();
            this.gcMultiRow3 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.template31 = new BLMT_OCR.OCR.Template3();
            this.gcMultiRow2 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.template21 = new BLMT_OCR.OCR.Template2();
            this.gcMultiRow1 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.template12 = new BLMT_OCR.OCR.Template1();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolTip1
            // 
            this.toolTip1.BackColor = System.Drawing.Color.LemonChiffon;
            // 
            // btnMinus
            // 
            this.btnMinus.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMinus.Image = ((System.Drawing.Image)(resources.GetObject("btnMinus.Image")));
            this.btnMinus.Location = new System.Drawing.Point(42, 668);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(37, 34);
            this.btnMinus.TabIndex = 9;
            this.btnMinus.TabStop = false;
            this.btnMinus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnMinus, "画像を縮小表示します");
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // btnPlus
            // 
            this.btnPlus.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPlus.Image = ((System.Drawing.Image)(resources.GetObject("btnPlus.Image")));
            this.btnPlus.Location = new System.Drawing.Point(5, 668);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(37, 34);
            this.btnPlus.TabIndex = 8;
            this.btnPlus.TabStop = false;
            this.btnPlus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnPlus, "画像を拡大表示します");
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // lblNoImage
            // 
            this.lblNoImage.Font = new System.Drawing.Font("メイリオ", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNoImage.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.lblNoImage.Location = new System.Drawing.Point(145, 329);
            this.lblNoImage.Name = "lblNoImage";
            this.lblNoImage.Size = new System.Drawing.Size(322, 42);
            this.lblNoImage.TabIndex = 119;
            this.lblNoImage.Text = "画像はありません";
            this.lblNoImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(5, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(595, 655);
            this.pictureBox1.TabIndex = 120;
            this.pictureBox1.TabStop = false;
            // 
            // leadImg
            // 
            this.leadImg.Location = new System.Drawing.Point(5, 4);
            this.leadImg.Name = "leadImg";
            this.leadImg.Size = new System.Drawing.Size(595, 658);
            this.leadImg.TabIndex = 121;
            this.leadImg.TabStop = false;
            this.leadImg.MouseLeave += new System.EventHandler(this.leadImg_MouseLeave);
            this.leadImg.MouseMove += new System.Windows.Forms.MouseEventHandler(this.leadImg_MouseMove);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblErrMsg);
            this.panel1.Location = new System.Drawing.Point(606, 669);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(630, 35);
            this.panel1.TabIndex = 162;
            // 
            // lblErrMsg
            // 
            this.lblErrMsg.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblErrMsg.Font = new System.Drawing.Font("ＭＳ ゴシック", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblErrMsg.ForeColor = System.Drawing.Color.Red;
            this.lblErrMsg.Location = new System.Drawing.Point(0, 0);
            this.lblErrMsg.Name = "lblErrMsg";
            this.lblErrMsg.Size = new System.Drawing.Size(626, 31);
            this.lblErrMsg.TabIndex = 0;
            this.lblErrMsg.Text = "label33";
            this.lblErrMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnRtn
            // 
            this.btnRtn.BackColor = System.Drawing.Color.Lavender;
            this.btnRtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRtn.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Location = new System.Drawing.Point(1009, 627);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(228, 33);
            this.btnRtn.TabIndex = 0;
            this.btnRtn.Text = "終了(&E)";
            this.btnRtn.UseVisualStyleBackColor = false;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click_1);
            // 
            // txtMemo
            // 
            this.txtMemo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtMemo.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo.Location = new System.Drawing.Point(1009, 372);
            this.txtMemo.Multiline = true;
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(228, 78);
            this.txtMemo.TabIndex = 4;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Lavender;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(1009, 545);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(228, 35);
            this.button2.TabIndex = 6;
            this.button2.Text = "データ印刷";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.Lavender;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button3.Location = new System.Drawing.Point(1008, 586);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(228, 35);
            this.button3.TabIndex = 7;
            this.button3.Text = "画像印刷";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // chkEdit
            // 
            this.chkEdit.AutoSize = true;
            this.chkEdit.Font = new System.Drawing.Font("ＭＳ ゴシック", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkEdit.Location = new System.Drawing.Point(1131, 452);
            this.chkEdit.Name = "chkEdit";
            this.chkEdit.Size = new System.Drawing.Size(106, 19);
            this.chkEdit.TabIndex = 5;
            this.chkEdit.Text = "編集モード";
            this.chkEdit.UseVisualStyleBackColor = true;
            this.chkEdit.CheckedChanged += new System.EventHandler(this.chkEdit_CheckedChanged);
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // lblEdit
            // 
            this.lblEdit.Font = new System.Drawing.Font("Meiryo UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblEdit.ForeColor = System.Drawing.Color.Red;
            this.lblEdit.Location = new System.Drawing.Point(1009, 497);
            this.lblEdit.Name = "lblEdit";
            this.lblEdit.Size = new System.Drawing.Size(228, 31);
            this.lblEdit.TabIndex = 309;
            this.lblEdit.Text = "現在、編集モードです";
            this.lblEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // gcMultiRow3
            // 
            this.gcMultiRow3.AllowUserToAddRows = false;
            this.gcMultiRow3.AllowUserToDeleteRows = false;
            this.gcMultiRow3.AllowUserToResize = false;
            this.gcMultiRow3.AllowUserToZoom = false;
            this.gcMultiRow3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow3.EditMode = GrapeCity.Win.MultiRow.EditMode.EditProgrammatically;
            this.gcMultiRow3.Location = new System.Drawing.Point(1009, 50);
            this.gcMultiRow3.Name = "gcMultiRow3";
            this.gcMultiRow3.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.gcMultiRow3.Size = new System.Drawing.Size(228, 307);
            this.gcMultiRow3.TabIndex = 3;
            this.gcMultiRow3.Template = this.template31;
            this.gcMultiRow3.Text = "gcMultiRow3";
            this.gcMultiRow3.CellValidating += new System.EventHandler<GrapeCity.Win.MultiRow.CellValidatingEventArgs>(this.gcMultiRow3_CellValidating);
            this.gcMultiRow3.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow3_CellValueChanged);
            this.gcMultiRow3.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow3_CellEnter);
            this.gcMultiRow3.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow3_EditingControlShowing);
            // 
            // gcMultiRow2
            // 
            this.gcMultiRow2.AllowUserToAddRows = false;
            this.gcMultiRow2.AllowUserToDeleteRows = false;
            this.gcMultiRow2.AllowUserToResize = false;
            this.gcMultiRow2.AllowUserToZoom = false;
            this.gcMultiRow2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow2.EditMode = GrapeCity.Win.MultiRow.EditMode.EditProgrammatically;
            this.gcMultiRow2.Location = new System.Drawing.Point(606, 3);
            this.gcMultiRow2.Name = "gcMultiRow2";
            this.gcMultiRow2.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.gcMultiRow2.Size = new System.Drawing.Size(631, 44);
            this.gcMultiRow2.TabIndex = 1;
            this.gcMultiRow2.Template = this.template21;
            this.gcMultiRow2.Text = "gcMultiRow2";
            this.gcMultiRow2.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow2_CellValueChanged);
            this.gcMultiRow2.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow2_CellEnter);
            this.gcMultiRow2.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow2_EditingControlShowing);
            // 
            // gcMultiRow1
            // 
            this.gcMultiRow1.AllowUserToAddRows = false;
            this.gcMultiRow1.AllowUserToDeleteRows = false;
            this.gcMultiRow1.AllowUserToResize = false;
            this.gcMultiRow1.AllowUserToZoom = false;
            this.gcMultiRow1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow1.Location = new System.Drawing.Point(606, 50);
            this.gcMultiRow1.Name = "gcMultiRow1";
            this.gcMultiRow1.ScrollBarMode = GrapeCity.Win.MultiRow.ScrollBarMode.Automatic;
            this.gcMultiRow1.Size = new System.Drawing.Size(397, 611);
            this.gcMultiRow1.TabIndex = 2;
            this.gcMultiRow1.Template = this.template12;
            this.gcMultiRow1.Text = "gcMultiRow1";
            this.gcMultiRow1.CellValidating += new System.EventHandler<GrapeCity.Win.MultiRow.CellValidatingEventArgs>(this.gcMultiRow1_CellValidating);
            this.gcMultiRow1.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellValueChanged);
            this.gcMultiRow1.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellEnter);
            this.gcMultiRow1.CellLeave += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellLeave);
            this.gcMultiRow1.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow1_EditingControlShowing);
            this.gcMultiRow1.CellContentClick += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellContentClick);
            this.gcMultiRow1.CurrentCellDirtyStateChanged += new System.EventHandler(this.gcMultiRow1_CurrentCellDirtyStateChanged);
            // 
            // frmPastCorrect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1244, 711);
            this.Controls.Add(this.lblEdit);
            this.Controls.Add(this.chkEdit);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.gcMultiRow3);
            this.Controls.Add(this.gcMultiRow2);
            this.Controls.Add(this.gcMultiRow1);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.lblNoImage);
            this.Controls.Add(this.leadImg);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmPastCorrect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "過去出勤簿表示";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCorrect_FormClosing);
            this.Load += new System.EventHandler(this.frmCorrect_Load);
            this.Shown += new System.EventHandler(this.frmCorrect_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblNoImage;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Leadtools.WinForms.RasterImageViewer leadImg;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblErrMsg;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow1;
        private Template1 template11;
        //private Template2 template21;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow2;
        private Template1 template12;
        private Template2 template21;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow3;
        private Template3 template31;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.CheckBox chkEdit;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.Label lblEdit;
    }
}