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
    public partial class frmEditLogRep : Form
    {
        string appName = "過去勤怠データ編集ログ一覧";          // アプリケーション表題

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去データ編集ログTableAdapter adp = new DataSet1TableAdapters.過去データ編集ログTableAdapter();
        DataSet1TableAdapters.ログインユーザーTableAdapter uAdp = new DataSet1TableAdapters.ログインユーザーTableAdapter();
      
        //clsStaff[] stf = null;              // スタッフクラス配列
        //clsShop[] shp = null;               // 店舗マスタークラス
        //clsArea[] ara = null;               // エリアマスタークラス
        //clsMana[] Mana = null;              // エリアマネージャークラス
        clsXlsmst[] xlsArray = null;

        public frmEditLogRep()
        {
            InitializeComponent();

            // データセットへデータを読み込む
            adp.Fill(dts.過去データ編集ログ);
            uAdp.Fill(dts.ログインユーザー);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            //Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //// スタッフ配列クラス作成
            //Utility u = new Utility();
            //u.getXlsToArray(global.cnfMsPath, xlsArray, ref stf);

            //// 店舗マスター配列クラス作成
            //u.getShopMaster(stf, ref shp);

            //// エリアマスター配列クラス作成
            //u.getAreaMaster(stf, ref ara);

            //// エリアマネージャー配列クラス作成
            //u.getManaMaster(stf, ref Mana);

            ////所属コードの桁数を取得する
            //string sqlSTRING = string.Empty;
            
            //// 店舗コンボロード
            //Utility.comboShop.Load(cmbBumonS, shp);
            //cmbBumonS.MaxDropDownItems = 20;
            //cmbBumonS.SelectedIndex = -1;

            // アカウントコンボロード
            Utility.comboAccount.Load(cmbBumonS, uAdp);
            cmbBumonS.MaxDropDownItems = 20;
            cmbBumonS.SelectedIndex = -1;

            //// エリアマネージャーコンボロード
            //Utility.comboMana.Load(cmbBumonS, Mana);
            //cmbBumonS.MaxDropDownItems = 20;
            //cmbBumonS.SelectedIndex = -1;

            //DataGridViewの設定
            GridViewSetting(dg1);

            // 対象年月を取得
            txtYear.Text = global.cnfYear.ToString();
            txtMonth.Text = global.cnfMonth.ToString();

            txtYear.Focus();
            
            button1.Enabled = false;    // CSV出力ボタン
            lblCnt.Visible = false;
        }

        string colStaffCode = "c1";
        string colStaffName = "c2";
        string colDate = "c3";
        string colEditDate = "c4";
        string colField = "c5";
        string colBefore = "c6";
        string colAfter = "c7";
        string colAccount = "c8";

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
                tempDGV.Columns.Add(colEditDate, "編集日時");
                tempDGV.Columns.Add(colStaffCode, "スタッフ№");
                tempDGV.Columns.Add(colStaffName, "スタッフ名");
                tempDGV.Columns.Add(colDate, "出勤簿日付");
                tempDGV.Columns.Add(colField, "項目");
                tempDGV.Columns.Add(colBefore, "編集前");
                tempDGV.Columns.Add(colAfter, "編集後");
                tempDGV.Columns.Add(colAccount, "編集者");

                tempDGV.Columns[colEditDate].Width = 160;
                tempDGV.Columns[colStaffCode].Width = 100;
                tempDGV.Columns[colStaffName].Width = 200;
                tempDGV.Columns[colDate].Width = 110;
                tempDGV.Columns[colField].Width = 130;
                tempDGV.Columns[colBefore].Width = 100;
                tempDGV.Columns[colAfter].Width = 100;
                tempDGV.Columns[colAccount].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                //tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colEditDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStaffCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colStaffName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colField].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colBefore].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colAfter].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colAccount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
             
                // 編集可否
                tempDGV.ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

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
        private void GridViewShowData(DataGridView g, int sYY, int sMM, int sYYto, int sMMto, int sCode)
        {
            // カーソル待機中
            this.Cursor = Cursors.WaitCursor;

            int sYYMM = sYY *100 + sMM;
            int eYYMM = sYYto * 100 + sMMto;

            // データグリッド行クリア
            g.Rows.Clear();

            try 
	        {
                var sss = dts.過去データ編集ログ.Where(a => (a.日時.Year * 100 + a.日時.Month) >= sYYMM && (a.日時.Year * 100 + a.日時.Month) <= eYYMM)
                    .OrderByDescending(a => a.日時);

                // 編集者指定
                if (sCode != global.flgOff)
                {
                    sss = sss.Where(a => a.編集アカウント == sCode).OrderByDescending(a => a.日時);
                }

                foreach (var t in sss)
                {
                    g.Rows.Add();

                    g[colEditDate, g.Rows.Count - 1].Value = t.日時;
                    g[colStaffCode, g.Rows.Count - 1].Value = t.出勤簿スタッフコード.ToString("D5");
                    g[colStaffName, g.Rows.Count - 1].Value = Utility.NulltoStr(t.出勤簿氏名);
                    if (t.出勤簿日 != string.Empty)
                    {
                        g[colDate, g.Rows.Count - 1].Value = t.出勤簿年 + "/" + t.出勤簿月.ToString("D2") + "/" +  t.出勤簿日.PadLeft(2, '0');
                    }
                    else
                    {
                        g[colDate, g.Rows.Count - 1].Value = t.出勤簿年 + "/" + t.出勤簿月.ToString("D2") + "   ";
                    }

                    if (t.Is項目名Null())
                    {
                        g[colField, g.Rows.Count - 1].Value = "";
                    }
                    else
                    {
                        g[colField, g.Rows.Count - 1].Value = Utility.NulltoStr(t.項目名);
                    }

                    g[colBefore, g.Rows.Count - 1].Value = Utility.NulltoStr(t.編集前値);
                    g[colAfter, g.Rows.Count - 1].Value = Utility.NulltoStr(t.編集後値);

                    if (t.ログインユーザーRow == null)
                    {
                        g[colAccount, g.Rows.Count - 1].Value = string.Empty;
                    }
                    else
                    {
                        g[colAccount, g.Rows.Count - 1].Value = t.ログインユーザーRow.名前;
                    }
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
            // 開始年月
            if (txtYear.Text != string.Empty || txtMonth.Text != string.Empty)
            {
                if (!Utility.NumericCheck(txtYear.Text))
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

                if (!Utility.NumericCheck(txtMonth.Text))
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
            }

            // 終了年月
            if (txtYearTo.Text != string.Empty || txtMonthTo.Text != string.Empty)
            {
                if (!Utility.NumericCheck(txtYearTo.Text))
                {
                    MessageBox.Show("年は数字で入力してください", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtYearTo.Focus();
                    return false;
                }

                if (!Utility.NumericCheck(txtMonthTo.Text))
                {
                    MessageBox.Show("月は数字で入力してください", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMonthTo.Focus();
                    return false;
                }

                if (int.Parse(txtMonthTo.Text) < 1 || int.Parse(txtMonthTo.Text) > 12)
                {
                    MessageBox.Show("月が正しくありません", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMonthTo.Focus();
                    return false;
                }
            }

            if (txtYear.Text != string.Empty && txtMonth.Text != string.Empty && 
                txtYearTo.Text != string.Empty && txtMonthTo.Text != string.Empty)
            {
                if ((Utility.StrtoInt(txtYear.Text) * 100 + Utility.StrtoInt(txtMonth.Text)) > (Utility.StrtoInt(txtYearTo.Text) * 100 + Utility.StrtoInt(txtMonthTo.Text)))
                {
                    MessageBox.Show("年月の範囲指定が正しくありません", appName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtYear.Focus();
                    return false;
                }
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
            //if (MessageBox.Show("終了します。よろしいですか？",appName,MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == DialogResult.No)
            //{
            //    e.Cancel = true;
            //    return;
            //}

            this.Dispose();
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
            DataSelect();
        }

        private void DataSelect()
        {
            // エラーチェック
            if (ErrCheck() == false) return;

            // エリア指定
            int sCode = 0;

            if (cmbBumonS.SelectedIndex != -1)
            {
                Utility.comboAccount cmb = (Utility.comboAccount)cmbBumonS.SelectedItem;
                sCode = cmb.code;
            }

            // 年月範囲
            int sYY = 0;
            int sYYto = 0;

            if (txtYear.Text == string.Empty)
            {
                sYY = 1900;
            }
            else
            {
                sYY = Utility.StrtoInt(txtYear.Text);
            }

            int sMM = Utility.StrtoInt(txtMonth.Text);

            if (txtYearTo.Text == string.Empty)
            {
                sYYto = 2999;
            }
            else
            {
                sYYto = Utility.StrtoInt(txtYearTo.Text);
            }

            int sMMto = Utility.StrtoInt(txtMonthTo.Text);

            // データ表示
            GridViewShowData(dg1, sYY, sMM, sYYto, sMMto, sCode);
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
            MyLibrary.CsvOut.GridView(dg1, "過去勤怠データ編集ログ一覧");
        }

    }
}
