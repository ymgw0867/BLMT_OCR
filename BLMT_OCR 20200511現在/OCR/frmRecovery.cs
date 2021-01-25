using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Odbc;
using BLMT_OCR.common;

namespace BLMT_OCR.OCR
{
    public partial class frmRecovery : Form
    {
        string appName = "出勤簿処理実施状況";      // アプリケーション表題

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter pmAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
      
        DataSet1TableAdapters.勤務票ヘッダTableAdapter hDtAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.勤務票明細TableAdapter mDtAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();

        DataSet1TableAdapters.保留勤務票ヘッダTableAdapter lAdp = new DataSet1TableAdapters.保留勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.保留勤務票明細TableAdapter lmAdp = new DataSet1TableAdapters.保留勤務票明細TableAdapter();

        clsStaff[] stf = null;              // スタッフクラス配列
        clsShop[] shp = null;               // 店舗マスタークラス
        clsArea[] ara = null;               // エリアマスタークラス
        clsXlsmst[] xlsArray = null;

        public frmRecovery()
        {
            InitializeComponent();

            // データセットへデータを読み込む
            phAdp.Fill(dts.過去勤務票ヘッダ);
            pmAdp.Fill(dts.過去勤務票明細);

            hDtAdp.Fill(dts.勤務票ヘッダ);
            mDtAdp.Fill(dts.勤務票明細);

            lAdp.Fill(dts.保留勤務票ヘッダ);
            lmAdp.Fill(dts.保留勤務票明細);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            //Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            // スタッフ配列クラス作成
            Utility u = new Utility();
            u.getXlsToArray(global.cnfMsPath, xlsArray, ref stf);

            //// 店舗マスター配列クラス作成
            //u.getShopMaster(stf, ref shp);

            // エリアマスター配列クラス作成
            u.getAreaMaster(stf, ref ara);

            ////所属コードの桁数を取得する
            //string sqlSTRING = string.Empty;
            
            //// 店舗コンボロード
            //Utility.comboShop.Load(cmbBumonS, shp);
            //cmbBumonS.MaxDropDownItems = 20;
            //cmbBumonS.SelectedIndex = -1;

            // 店舗コンボロード
            Utility.comboArea.Load(cmbBumonS, ara);
            cmbBumonS.MaxDropDownItems = 20;
            cmbBumonS.SelectedIndex = -1;

            //DataGridViewの設定
            GridViewSetting(dg1);

            // 対象年月を取得
            txtYear.Text = global.cnfYear.ToString();
            txtMonth.Text = global.cnfMonth.ToString();

            txtYear.Focus();
            
            // 表示選択コンボボックス
            comboBox1.Items.Add("すべて表示");
            comboBox1.Items.Add("処理済");
            comboBox1.Items.Add("処理中");
            comboBox1.Items.Add("未処理");
            comboBox1.SelectedIndex = 0;

            // CSV出力ボタン
            button1.Enabled = false;

            lblCnt.Visible = false;
        }

        string colBushoCode = "c2";
        string colBushoName = "c3";
        string colAreaCode = "c6";
        string colAreaName = "c7";
        string colStaffCode = "c8";
        string colStaffName = "c9";
        string colOcr = "c4";
        string colkbn = "c10";
        string colStatus = "c11";
        string colDate = "c12";
        string colID = "c5";
        string colShimebi = "c13";

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する
                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", (float)10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 582;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colAreaCode, "エリア№");
                tempDGV.Columns.Add(colAreaName, "エリア名");
                tempDGV.Columns.Add(colBushoCode, "店舗№");
                tempDGV.Columns.Add(colBushoName, "店舗名");
                tempDGV.Columns.Add(colStaffCode, "スタッフ№");
                tempDGV.Columns.Add(colStaffName, "氏名");
                tempDGV.Columns.Add(colkbn, "給与区分");
                tempDGV.Columns.Add(colShimebi, "締日");
                tempDGV.Columns.Add(colOcr, "OCR");
                tempDGV.Columns.Add(colStatus, "状況");
                tempDGV.Columns.Add(colDate, "作成日");

                tempDGV.Columns[colAreaCode].Width = 90;
                tempDGV.Columns[colAreaName].Width = 100;
                tempDGV.Columns[colBushoCode].Width = 80;
                tempDGV.Columns[colBushoName].Width = 220;
                tempDGV.Columns[colStaffCode].Width = 100;
                tempDGV.Columns[colStaffName].Width = 120;
                tempDGV.Columns[colkbn].Width = 100;
                tempDGV.Columns[colShimebi].Width = 54;
                tempDGV.Columns[colOcr].Width = 50;
                tempDGV.Columns[colStatus].Width = 200;
                tempDGV.Columns[colDate].Width = 140;

                //tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colAreaCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colBushoCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStaffCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colkbn].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colShimebi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colOcr].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                tempDGV.ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                ////ソート機能制限
                //for (int i = 0; i < tempDGV.Columns.Count; i++)
                //{
                //    tempDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ社員情報を表示する </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        /// <param name="sCode">
        ///     指定所属コード</param>
        /// ----------------------------------------------------------------------
        private void GridViewShowData(DataGridView g, int sCode, int sMode)
        {
            // カーソル待機中
            this.Cursor = Cursors.WaitCursor;

            // データグリッド行クリア
            g.Rows.Clear();
            int iX = 0;

            try 
	        {
                var sss = stf.OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード);

                if (cmbBumonS.SelectedIndex != -1)
                {
                    sss = sss.Where(a => a.エリアコード == sCode).OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード);
                }
                
                foreach (var t in sss)
                {
                    DateTime dt;
                    int sYY = Utility.StrtoInt(txtYear.Text);

                    if (sYY > 2000)
                    {
                        sYY -= 2000;
                    }

                    // 過去勤務票データ存在チェック
                    if (isPastData(sYY, Utility.StrtoInt(txtMonth.Text), t.スタッフコード, out dt))
                    {
                        if (sMode != 0 && sMode != 1)
                        {
                            continue;
                        }

                        g.Rows.Add();

                        // 勤怠データ作成済み
                        g[colOcr, g.Rows.Count - 1].Value = "〇";    // OCR
                        g[colStatus, g.Rows.Count - 1].Value = "PCA給与Xデータ作成済み"; // 状況
                        g[colDate, g.Rows.Count - 1].Value = dt.ToString("yyyy/MM/dd HH:mm:ss"); // 作成年月日

                    }
                    else if (isOcrData(sYY, Utility.StrtoInt(txtMonth.Text), t.スタッフコード, out dt))  // 勤務票データ存在チェック
                    {
                        if (sMode != 0 && sMode != 2)
                        {
                            continue;
                        }

                        g.Rows.Add();

                        // 現在処理中
                        g[colOcr, g.Rows.Count - 1].Value = "〇";        // OCR
                        g[colStatus, g.Rows.Count - 1].Value = "処理中"; // 状況
                        g[colDate, g.Rows.Count - 1].Value = dt.ToString("yyyy/MM/dd HH:mm:ss");  // 作成年月日

                        g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightBlue;
                    }
                    else if (isHoldData(sYY, Utility.StrtoInt(txtMonth.Text), t.スタッフコード, out dt))  // 保留勤務票データ存在チェック
                    {
                        if (sMode != 0 && sMode != 2)
                        {
                            continue;
                        }

                        g.Rows.Add();

                        // 現在処理中
                        g[colOcr, g.Rows.Count - 1].Value = "〇";        // OCR
                        g[colStatus, g.Rows.Count - 1].Value = "処理中（保留）";   // 状況
                        g[colDate, g.Rows.Count - 1].Value = dt.ToString("yyyy/MM/dd HH:mm:ss");    // 作成年月日

                        g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightBlue;
                    }
                    else
                    {
                        if (sMode != 0 && sMode != 3)
                        {
                            continue;
                        }

                        g.Rows.Add();

                        // OCR未実施
                        g[colOcr, g.Rows.Count - 1].Value = "×";                // OCR
                        g[colStatus, g.Rows.Count - 1].Value = "ＯＣＲ未実施";    // 状況
                        g[colDate, g.Rows.Count - 1].Value = string.Empty;   // 作成年月日

                        g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightPink;
                    }
                    
                    g[colAreaCode, g.Rows.Count - 1].Value = t.エリアコード.ToString("D3");
                    g[colAreaName, g.Rows.Count - 1].Value = t.エリア名;
                    g[colBushoCode, g.Rows.Count - 1].Value = t.店舗コード.ToString("D5");
                    g[colBushoName, g.Rows.Count - 1].Value = t.店舗名;
                    g[colStaffCode, g.Rows.Count - 1].Value = t.スタッフコード.ToString("D5");
                    g[colStaffName, g.Rows.Count - 1].Value = t.スタッフ名;

                    if (t.給与区分.ToString() == global.NIKKYU)
                    {
                        g[colkbn, g.Rows.Count - 1].Value = "日給";
                    }
                    else if (t.給与区分.ToString() == global.JIKKYU)
                    {
                        g[colkbn, g.Rows.Count - 1].Value = "時給";
                    }

                    g[colShimebi, g.Rows.Count - 1].Value = t.締日区分;

                }
            
                g.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                // カーソルを戻す
                this.Cursor = Cursors.Default;
            }

            // 該当するデータがないとき
            if (g.RowCount == 0)
            {
                MessageBox.Show("該当するデータはありませんでした", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button1.Enabled = false;
                lblCnt.Visible = false;
            }
            else
            {
                button1.Enabled = true;
                lblCnt.Visible = true;
                lblCnt.Text = g.RowCount.ToString("#,##0") + "件";
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータが存在するか </summary>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        /// <param name="sCode">
        ///     スタッフコード</param>
        /// <param name="dDate">
        ///     更新年月日</param>
        /// <returns>
        ///     true:データあり、false:データなし</returns>
        ///----------------------------------------------------------------------
        private bool isPastData(int sYY, int sMM, int sCode, out DateTime dDate)
        {
            dDate = DateTime.Today;

            if (dts.過去勤務票ヘッダ.Any(a => a.スタッフコード == sCode && a.年 == sYY && a.月 == sMM))
            {
                foreach (var item in dts.過去勤務票ヘッダ.Where(a => a.スタッフコード == sCode && a.年 == sYY && a.月 == sMM))
	            {
                    dDate = item.更新年月日;
	            }

                return true;
            }
            else
            {
                return false;
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     保留勤務票ヘッダデータが存在するか </summary>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        /// <param name="sCode">
        ///     スタッフコード</param>
        /// <returns>
        ///     true:データあり、false:データなし</returns>
        ///----------------------------------------------------------------------
        private bool isHoldData(int sYY, int sMM, int sCode, out DateTime dDate)
        {
            dDate = DateTime.Today;

            if (dts.保留勤務票ヘッダ.Any(a => a.スタッフコード == sCode && a.年 == sYY && a.月 == sMM))
            {
                foreach (var item in dts.保留勤務票ヘッダ.Where(a => a.スタッフコード == sCode && a.年 == sYY && a.月 == sMM))
                {
                    dDate = item.更新年月日;
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダデータが存在するか </summary>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        /// <param name="sCode">
        ///     スタッフコード</param>
        /// <returns>
        ///     true:データあり、false:データなし</returns>
        ///----------------------------------------------------------------------
        private bool isOcrData(int sYY, int sMM, int sCode, out DateTime dDate)
        {
            dDate = DateTime.Today;

            if (dts.勤務票ヘッダ.Any(a => a.スタッフコード == sCode && a.年 == sYY && a.月 == sMM))
            {
                foreach (var item in dts.勤務票ヘッダ.Where(a => a.スタッフコード == sCode && a.年 == sYY && a.月 == sMM))
                {
                    dDate = item.更新年月日;
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        private Boolean ErrCheck()
        {
            if (Utility.NumericCheck(txtYear.Text) == false)
            {
                MessageBox.Show("年は数字で入力してください", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.NumericCheck(txtMonth.Text) == false)
            {
                MessageBox.Show("月は数字で入力してください", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("月が正しくありません", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            return true;
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();
            
            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            txtObj.SelectAll();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("終了します。よろしいですか？",appName,MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }

            this.Dispose();
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
            DataSelect(comboBox1.SelectedIndex);
        }

        private void DataSelect(int sMode)
        {
            //エラーチェック
            if (ErrCheck() == false) return;

            int sArea = 0;

            if (cmbBumonS.SelectedIndex != -1)
            {
                Utility.comboArea cmb = (Utility.comboArea)cmbBumonS.SelectedItem;
                sArea = cmb.Code;
            }

            //データ表示
            GridViewShowData(dg1, sArea, sMode);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.KeyCode == Keys.Enter)
            //{
            //    if (!e.Control)
            //    {
            //        this.SelectNextControl(this.ActiveControl, !e.Shift, true, true, true);
            //    }
            //}
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {

            //if (e.KeyChar == (char)Keys.Enter)
            //{
            //    e.Handled = true;
            //}
        }

        private void rBtnPrn_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void prePrint_Shown(object sender, EventArgs e)
        {
            txtYear.Focus();
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b') 
                e.Handled = true;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyLibrary.CsvOut.GridView(dg1, "出勤簿処理状況");
        }

    }
}
