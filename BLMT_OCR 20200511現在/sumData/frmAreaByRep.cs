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

namespace BLMT_OCR.sumData
{
    public partial class frmAreaByRep : Form
    {
        //所属コード設定桁数
        int ShozokuLen = 0;
         
        string _ComNo = string.Empty;                   // 会社番号
        string _ComName = string.Empty;                 // 会社名
        string _ComDatabeseName = string.Empty;         // 会社データベース名

        const int MODE_ALL = 0;                         // 全員
        const int MODE_ZMIKAISHU = 1;                   // 前半未回収
        const int MODE_ZSHORICHU = 2;                   // 前半処理中
        const int MODE_ZSHORIZUMI = 3;                  // 前半処理済み
        const int MODE_KMIKAISHU = 4;                   // 後半未回収
        const int MODE_KSHORICHU = 5;                   // 後半処理中
        const int MODE_KSHORIZUMI = 6;                  // 後半処理済み
        const int MODE_SHORIZUMI = 7;                   // 月間処理済み
        const int MODE_MISHORI = 8;                     // 月間未処理

        string appName = "営業担当別勤怠明細一覧";          // アプリケーション表題

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter pmAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
      
        clsStaff[] stf = null;              // スタッフクラス配列
        //clsShop[] shp = null;               // 店舗マスタークラス
        //clsArea[] ara = null;               // エリアマスタークラス
        clsMana[] Mana = null;              // エリアマネージャークラス
        clsXlsmst[] xlsArray = null;

        public frmAreaByRep()
        {
            InitializeComponent();

            // データセットへデータを読み込む
            phAdp.Fill(dts.過去勤務票ヘッダ);
            pmAdp.Fill(dts.過去勤務票明細);
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

            //// エリアマスター配列クラス作成
            //u.getAreaMaster(stf, ref ara);

            // エリアマネージャー配列クラス作成
            u.getManaMaster(stf, ref Mana);

            //所属コードの桁数を取得する
            string sqlSTRING = string.Empty;
            
            //// 店舗コンボロード
            //Utility.comboShop.Load(cmbBumonS, shp);
            //cmbBumonS.MaxDropDownItems = 20;
            //cmbBumonS.SelectedIndex = -1;

            //// エリアコンボロード
            //Utility.comboArea.Load(cmbBumonS, ara);
            //cmbBumonS.MaxDropDownItems = 20;
            //cmbBumonS.SelectedIndex = -1;

            // エリアマネージャーコンボロード
            Utility.comboMana.Load(cmbBumonS, Mana);
            cmbBumonS.MaxDropDownItems = 20;
            cmbBumonS.SelectedIndex = -1;

            //DataGridViewの設定
            GridViewSetting(dg1);

            // 対象年月を取得
            txtYear.Text = global.cnfYear.ToString();
            txtMonth.Text = global.cnfMonth.ToString();

            txtYear.Focus();
            
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
        string colTlDay = "c14";
        string colWDay = "c15";
        string colYuDay = "c16";
        string colKDay = "c17";
        string colChisou = "c18";
        string colWTime = "c19";
        string colNaiZan = "c20";
        string colMashiZan = "c21";
        string col20Over = "c22";
        string col22Over = "c23";
        string colDonichi = "c24";
        string colKoutsuhi = "c25";
        string colSonota = "c26";
        string colYear = "c27";
        string colMonth = "c28";
        string colAreaMana = "c29";

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
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colYear, "年");
                tempDGV.Columns.Add(colMonth, "月");
                tempDGV.Columns.Add(colAreaMana, "担当");
                tempDGV.Columns.Add(colAreaCode, "エリア№");
                tempDGV.Columns.Add(colAreaName, "エリア名");
                tempDGV.Columns.Add(colBushoCode, "店舗№");
                tempDGV.Columns.Add(colBushoName, "店舗名");
                tempDGV.Columns.Add(colStaffCode, "スタッフ№");
                tempDGV.Columns.Add(colStaffName, "氏名");
                tempDGV.Columns.Add(colkbn, "給与形態");
                tempDGV.Columns.Add(colTlDay, "要出勤日数");
                tempDGV.Columns.Add(colWDay, "実労日数");
                tempDGV.Columns.Add(colYuDay, "有給日数");
                tempDGV.Columns.Add(colKDay, "公休日数");
                tempDGV.Columns.Add(colChisou, "遅早時間");
                tempDGV.Columns.Add(colWTime, "実労働時間");
                tempDGV.Columns.Add(colNaiZan, "時間内残業");
                tempDGV.Columns.Add(colMashiZan, "割増残業");
                tempDGV.Columns.Add(col20Over, "20時以降勤務");
                tempDGV.Columns.Add(col22Over, "22時以降勤務");
                tempDGV.Columns.Add(colDonichi, "土日祝勤務");
                tempDGV.Columns.Add(colKoutsuhi, "交通費");
                tempDGV.Columns.Add(colSonota, "その他");

                tempDGV.Columns[colYear].Width = 50;
                tempDGV.Columns[colMonth].Width = 50;
                tempDGV.Columns[colAreaCode].Width = 90;
                tempDGV.Columns[colAreaName].Width = 90;
                tempDGV.Columns[colBushoCode].Width = 90;
                tempDGV.Columns[colBushoName].Width = 200;
                tempDGV.Columns[colStaffCode].Width = 100;
                tempDGV.Columns[colStaffName].Width = 120;
                tempDGV.Columns[colkbn].Width = 80;
                tempDGV.Columns[colTlDay].Width = 90;
                tempDGV.Columns[colWDay].Width = 90;
                tempDGV.Columns[colYuDay].Width = 90;
                tempDGV.Columns[colKDay].Width = 90;
                tempDGV.Columns[colChisou].Width = 90;
                tempDGV.Columns[colWTime].Width = 90;
                tempDGV.Columns[colNaiZan].Width = 90;
                tempDGV.Columns[colMashiZan].Width = 90;
                tempDGV.Columns[col20Over].Width = 100;
                tempDGV.Columns[col22Over].Width = 100;
                tempDGV.Columns[colDonichi].Width = 90;
                tempDGV.Columns[colKoutsuhi].Width = 70;
                tempDGV.Columns[colSonota].Width = 70;

                //tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colYear].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colMonth].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colAreaCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colBushoCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStaffCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colkbn].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colWDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colTlDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colYuDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colKDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colChisou].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colWTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colNaiZan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colMashiZan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[col20Over].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[col22Over].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDonichi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colKoutsuhi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colSonota].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
 
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
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                //// 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

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
        private void GridViewShowData(DataGridView g, int sYY, int sMM, string sCode)
        {
            // カーソル待機中
            this.Cursor = Cursors.WaitCursor;

            // データグリッド行クリア
            g.Rows.Clear();

            try 
	        {
                var sss = dts.過去勤務票ヘッダ.Where(a => a.年 == sYY && a.月 == sMM).OrderBy(a => a.担当エリアマネージャー名).ThenBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);

                if (cmbBumonS.SelectedIndex != -1)
                {
                    sss = sss.Where(a => a.担当エリアマネージャー名 == sCode).OrderBy(a => a.担当エリアマネージャー名).ThenBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);
                }
                
                foreach (var t in sss)
                {
                    g.Rows.Add();

                    g[colYear, g.Rows.Count - 1].Value = "20" + t.年;
                    g[colMonth, g.Rows.Count - 1].Value = t.月;
                    g[colAreaMana, g.Rows.Count - 1].Value = t.担当エリアマネージャー名;
                    g[colAreaCode, g.Rows.Count - 1].Value = t.エリアコード.ToString("D3");
                    g[colAreaName, g.Rows.Count - 1].Value = t.エリア名;
                    g[colBushoCode, g.Rows.Count - 1].Value = t.店舗コード.ToString("D5");
                    g[colBushoName, g.Rows.Count - 1].Value = t.店舗名;
                    g[colStaffCode, g.Rows.Count - 1].Value = t.スタッフコード.ToString("D5");
                    g[colStaffName, g.Rows.Count - 1].Value = t.氏名;

                    if (t.給与形態.ToString() == global.NIKKYU)
                    {
                        g[colkbn, g.Rows.Count - 1].Value = "日給";
                    }
                    else if (t.給与形態.ToString() == global.JIKKYU)
                    {
                        g[colkbn, g.Rows.Count - 1].Value = "時給";
                    }

                    g[colTlDay, g.Rows.Count - 1].Value = t.要出勤日数;
                    g[colWDay, g.Rows.Count - 1].Value = t.実労日数;
                    g[colYuDay, g.Rows.Count - 1].Value = t.有休日数;
                    g[colKDay, g.Rows.Count - 1].Value = t.公休日数;
                    g[colChisou, g.Rows.Count - 1].Value = string.Format("{0, 3}", t.遅早時間時) + ":" + t.遅早時間分.ToString("D2");
                    g[colWTime, g.Rows.Count - 1].Value = string.Format("{0, 3}", t.実労働時間時) + ":" + t.実労働時間分.ToString("D2");
                    g[colNaiZan, g.Rows.Count - 1].Value = string.Format("{0, 3}", t.基本時間内残業時) + ":" + t.基本時間内残業分.ToString("D2");
                    g[colMashiZan, g.Rows.Count - 1].Value = string.Format("{0, 3}", t.割増残業時) + ":" + t.割増残業分.ToString("D2");
                    g[col20Over, g.Rows.Count - 1].Value = string.Format("{0, 3}", t._20時以降勤務時) + ":" + t._20時以降勤務分.ToString("D2");
                    g[col22Over, g.Rows.Count - 1].Value = string.Format("{0, 3}", t._22時以降勤務時) + ":" + t._22時以降勤務分.ToString("D2");
                    g[colDonichi, g.Rows.Count - 1].Value = string.Format("{0, 3}", t.土日祝日労働時間時) + ":" + t.土日祝日労働時間分.ToString("D2");
                    g[colKoutsuhi, g.Rows.Count - 1].Value = t.交通費.ToString("#,##0");
                    g[colSonota, g.Rows.Count - 1].Value = t.その他支給.ToString("#,##0");                    
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

        private Boolean ErrCheck()
        {
            if (Utility.NumericCheck(txtYear.Text) == false)
            {
                MessageBox.Show("年は数字で入力してください", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtYear.Text) < 2017)
            {
                MessageBox.Show("年を正しく入力してください", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
            DataSelect();
        }

        private void DataSelect()
        {
            //エラーチェック
            if (ErrCheck() == false) return;

            string sMana = string.Empty;

            if (cmbBumonS.SelectedIndex != -1)
            {
                Utility.comboMana cmb = (Utility.comboMana)cmbBumonS.SelectedItem;
                sMana = cmb.Code;
            }

            int sYY = Utility.StrtoInt(txtYear.Text) - 2000;
            int sMM = Utility.StrtoInt(txtMonth.Text);

            //データ表示
            GridViewShowData(dg1, sYY, sMM, sMana);
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
            MyLibrary.CsvOut.GridView(dg1, "営業担当別勤怠明細一覧");
        }

    }
}
