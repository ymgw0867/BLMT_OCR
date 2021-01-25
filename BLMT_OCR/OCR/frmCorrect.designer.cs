namespace BLMT_OCR.OCR
{
    partial class frmCorrect
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCorrect));
            this.hScrollBar1 = new System.Windows.Forms.HScrollBar();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnMinus = new System.Windows.Forms.Button();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnEnd = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBefore = new System.Windows.Forms.Button();
            this.btnFirst = new System.Windows.Forms.Button();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblErrMsg = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.btnRtn = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnDataMake = new System.Windows.Forms.Button();
            this.btnErrCheck = new System.Windows.Forms.Button();
            this.lblCnt = new System.Windows.Forms.Label();
            this.btnHold = new System.Windows.Forms.Button();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
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
            // hScrollBar1
            // 
            this.hScrollBar1.Location = new System.Drawing.Point(228, 668);
            this.hScrollBar1.Name = "hScrollBar1";
            this.hScrollBar1.Size = new System.Drawing.Size(332, 34);
            this.hScrollBar1.TabIndex = 13;
            this.toolTip1.SetToolTip(this.hScrollBar1, "出勤簿を移動します");
            this.hScrollBar1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.hScrollBar1_Scroll);
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
            this.btnMinus.TabIndex = 8;
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
            this.btnPlus.TabIndex = 7;
            this.btnPlus.TabStop = false;
            this.btnPlus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnPlus, "画像を拡大表示します");
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnEnd
            // 
            this.btnEnd.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnEnd.Image = ((System.Drawing.Image)(resources.GetObject("btnEnd.Image")));
            this.btnEnd.Location = new System.Drawing.Point(190, 668);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(37, 34);
            this.btnEnd.TabIndex = 12;
            this.btnEnd.TabStop = false;
            this.toolTip1.SetToolTip(this.btnEnd, "最後尾の出勤簿データへ移動します");
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.Location = new System.Drawing.Point(153, 668);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(37, 34);
            this.btnNext.TabIndex = 11;
            this.btnNext.TabStop = false;
            this.toolTip1.SetToolTip(this.btnNext, "次の出勤簿データへ移動します");
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnBefore
            // 
            this.btnBefore.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnBefore.Image = ((System.Drawing.Image)(resources.GetObject("btnBefore.Image")));
            this.btnBefore.Location = new System.Drawing.Point(116, 668);
            this.btnBefore.Name = "btnBefore";
            this.btnBefore.Size = new System.Drawing.Size(37, 34);
            this.btnBefore.TabIndex = 10;
            this.btnBefore.TabStop = false;
            this.toolTip1.SetToolTip(this.btnBefore, "前の出勤簿データへ移動します");
            this.btnBefore.UseVisualStyleBackColor = true;
            this.btnBefore.Click += new System.EventHandler(this.btnBefore_Click);
            // 
            // btnFirst
            // 
            this.btnFirst.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnFirst.Image")));
            this.btnFirst.Location = new System.Drawing.Point(79, 668);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(37, 34);
            this.btnFirst.TabIndex = 9;
            this.btnFirst.TabStop = false;
            this.toolTip1.SetToolTip(this.btnFirst, "先頭の出勤簿データへ移動します");
            this.btnFirst.UseVisualStyleBackColor = true;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
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
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Location = new System.Drawing.Point(563, 667);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(37, 37);
            this.button1.TabIndex = 14;
            this.button1.TabStop = false;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.checkBox1.Location = new System.Drawing.Point(1016, 487);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(75, 20);
            this.checkBox1.TabIndex = 5;
            this.checkBox1.Text = "確認済";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
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
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.Lavender;
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDelete.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDelete.Location = new System.Drawing.Point(1125, 588);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(112, 33);
            this.btnDelete.TabIndex = 9;
            this.btnDelete.Text = "削除(&D)";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnDataMake
            // 
            this.btnDataMake.BackColor = System.Drawing.Color.Lavender;
            this.btnDataMake.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDataMake.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDataMake.Location = new System.Drawing.Point(1125, 549);
            this.btnDataMake.Name = "btnDataMake";
            this.btnDataMake.Size = new System.Drawing.Size(112, 33);
            this.btnDataMake.TabIndex = 7;
            this.btnDataMake.Text = "データ作成(&G)";
            this.btnDataMake.UseVisualStyleBackColor = false;
            this.btnDataMake.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnErrCheck
            // 
            this.btnErrCheck.BackColor = System.Drawing.Color.Lavender;
            this.btnErrCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnErrCheck.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnErrCheck.Location = new System.Drawing.Point(1009, 549);
            this.btnErrCheck.Name = "btnErrCheck";
            this.btnErrCheck.Size = new System.Drawing.Size(112, 33);
            this.btnErrCheck.TabIndex = 6;
            this.btnErrCheck.Text = "エラーチェック(&C)";
            this.btnErrCheck.UseVisualStyleBackColor = false;
            this.btnErrCheck.Click += new System.EventHandler(this.button3_Click);
            // 
            // lblCnt
            // 
            this.lblCnt.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblCnt.Location = new System.Drawing.Point(1144, 485);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(93, 25);
            this.lblCnt.TabIndex = 303;
            this.lblCnt.Text = "lblCnt";
            this.lblCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnHold
            // 
            this.btnHold.BackColor = System.Drawing.Color.Lavender;
            this.btnHold.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnHold.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnHold.Location = new System.Drawing.Point(1009, 588);
            this.btnHold.Name = "btnHold";
            this.btnHold.Size = new System.Drawing.Size(112, 33);
            this.btnHold.TabIndex = 8;
            this.btnHold.Text = "保留(&L)";
            this.btnHold.UseVisualStyleBackColor = false;
            this.btnHold.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // txtMemo
            // 
            this.txtMemo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtMemo.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtMemo.Location = new System.Drawing.Point(1008, 376);
            this.txtMemo.Multiline = true;
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(228, 87);
            this.txtMemo.TabIndex = 4;
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
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
            this.gcMultiRow1.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellValueChanged);
            this.gcMultiRow1.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellEnter);
            this.gcMultiRow1.CellLeave += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellLeave);
            this.gcMultiRow1.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow1_EditingControlShowing);
            this.gcMultiRow1.CellContentClick += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellContentClick);
            this.gcMultiRow1.CurrentCellDirtyStateChanged += new System.EventHandler(this.gcMultiRow1_CurrentCellDirtyStateChanged);
            // 
            // frmCorrect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1244, 711);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.btnHold);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.btnErrCheck);
            this.Controls.Add(this.btnDataMake);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.gcMultiRow3);
            this.Controls.Add(this.gcMultiRow2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.gcMultiRow1);
            this.Controls.Add(this.hScrollBar1);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnBefore);
            this.Controls.Add(this.btnFirst);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.lblNoImage);
            this.Controls.Add(this.leadImg);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmCorrect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "勤怠データ登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCorrect_FormClosing);
            this.Load += new System.EventHandler(this.frmCorrect_Load);
            this.Shown += new System.EventHandler(this.frmCorrect_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.HScrollBar hScrollBar1;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnBefore;
        private System.Windows.Forms.Button btnFirst;
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
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow2;
        private Template1 template12;
        private Template2 template21;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow3;
        private Template3 template31;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnDataMake;
        private System.Windows.Forms.Button btnErrCheck;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.Button btnHold;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Drawing.Printing.PrintDocument printDocument1;
    }
}