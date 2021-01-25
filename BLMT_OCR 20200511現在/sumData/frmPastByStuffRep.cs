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
    public partial class frmPastByStuffRep : Form
    {
        string appName = "勤怠データ一覧";          // アプリケーション表題

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter pmAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
      
        clsStaff[] stf = null;              // スタッフクラス配列
        clsShop[] shp = null;               // 店舗マスタークラス
        clsArea[] ara = null;               // エリアマスタークラス
        //clsMana[] Mana = null;              // エリアマネージャークラス
        clsXlsmst[] xlsArray = null;

        public frmPastByStuffRep(int sYY, int sMM, int sYYto, int sMMto, int sArea, string sName)
        {
            InitializeComponent();

            // データセットへデータを読み込む
            phAdp.Fill(dts.過去勤務票ヘッダ);
            pmAdp.Fill(dts.過去勤務票明細);

            pYY = sYY;
            pMM = sMMto;
            pYYto = sYYto;
            pMMto = sMMto;
            pName = sName;
        }

        int pYY = 0;
        int pMM = 0;
        int pYYto = 0;
        int pMMto = 0;
        int pArea = 0;
        string pName = string.Empty;

        private void Form1_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            //Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            // スタッフ配列クラス作成
            Utility u = new Utility();
            u.getXlsToArray(global.cnfMsPath, xlsArray, ref stf);

            // 店舗マスター配列クラス作成
            u.getShopMaster(stf, ref shp);

            // エリアマスター配列クラス作成
            u.getAreaMaster(stf, ref ara);

            //// エリアマネージャー配列クラス作成
            //u.getManaMaster(stf, ref Mana);

            ////所属コードの桁数を取得する
            //string sqlSTRING = string.Empty;
            
            //// 店舗コンボロード
            //Utility.comboShop.Load(cmbBumonS, shp);
            //cmbBumonS.MaxDropDownItems = 20;
            //cmbBumonS.SelectedIndex = -1;

            // エリアコンボロード
            Utility.comboArea.Load(cmbBumonS, ara);
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
            
            // CSV出力ボタン
            button1.Enabled = false;

            lblCnt.Visible = false;

            if (pYY != 0)
            {
                //データ表示
                GridViewShowData(dg1, pYY, pMM, pYYto, pMMto, pArea, pName);
            }
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
        string colShhmm = "c30";
        string colEhhmm = "c31";
        string colRest = "c32";
        string colDayShopCode = "c33";
        string colDayShopName = "c34";

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
                tempDGV.Columns.Add(colAreaCode, "エリア№");
                tempDGV.Columns.Add(colAreaName, "エリア名");
                tempDGV.Columns.Add(colAreaMana, "担当");
                tempDGV.Columns.Add(colBushoCode, "店舗№");
                tempDGV.Columns.Add(colBushoName, "店舗名");
                tempDGV.Columns.Add(colStaffCode, "スタッフ№");
                tempDGV.Columns.Add(colStaffName, "氏名");
                tempDGV.Columns.Add(colkbn, "給与形態");
                tempDGV.Columns.Add(colDate, "年月日");
                tempDGV.Columns.Add(colStatus, "出勤状況");
                tempDGV.Columns.Add(colShhmm, "出勤時刻");
                tempDGV.Columns.Add(colEhhmm, "終業時刻");
                tempDGV.Columns.Add(colRest, "休憩");
                tempDGV.Columns.Add(colDayShopCode, "就労店舗№");
                tempDGV.Columns.Add(colDayShopName, "就労店舗名");

                tempDGV.Columns[colDate].Width = 120;
                tempDGV.Columns[colStaffCode].Width = 100;
                tempDGV.Columns[colStaffName].Width = 120;
                tempDGV.Columns[colAreaCode].Width = 90;
                tempDGV.Columns[colAreaName].Width = 90;
                tempDGV.Columns[colBushoCode].Width = 90;
                tempDGV.Columns[colBushoName].Width = 200;
                tempDGV.Columns[colkbn].Width = 80;
                tempDGV.Columns[colStatus].Width = 70;
                tempDGV.Columns[colShhmm].Width = 70;
                tempDGV.Columns[colEhhmm].Width = 70;
                tempDGV.Columns[colRest].Width = 60;
                tempDGV.Columns[colDayShopCode].Width = 80;
                //tempDGV.Columns[colDayShopName].Width = 200;

                tempDGV.Columns[colDayShopName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colAreaCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colBushoCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStaffCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colkbn].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStatus].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colShhmm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colEhhmm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colRest].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDayShopCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                tempDGV.Columns[colAreaCode].Visible = false;
                tempDGV.Columns[colAreaName].Visible = false;
                tempDGV.Columns[colBushoCode].Visible = false;
                tempDGV.Columns[colBushoName].Visible = false;
                tempDGV.Columns[colAreaMana].Visible = false;


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
        private void GridViewShowData(DataGridView g, int sYY, int sMM, int sYYto, int sMMto, int sArea, string sName)
        {
            // カーソル待機中
            this.Cursor = Cursors.WaitCursor;

            int sYYMM = sYY *100 + sMM;
            int eYYMM = sYYto * 100 + sMMto;

            // データグリッド行クリア
            g.Rows.Clear();

            try 
	        {
                var sss = dts.過去勤務票ヘッダ.Where(a => (a.年 * 100 + a.月) >= sYYMM && (a.年 * 100 + a.月) <=  eYYMM)
                                             .OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);

                // エリア指定
                if (sArea != 0)
                {
                    sss = sss.Where(a => a.エリアコード == sArea).OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);
                }

                // スタッフ指定
                if (sName != string.Empty)
                {
                    sss = sss.Where(a => a.氏名.Contains(sName)).OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);
                }

                foreach (var t in sss)
                {
                    if (t.Get過去勤務票明細Rows().Count() == 0)
                    {
                        continue;
                    }

                    foreach (var m in t.Get過去勤務票明細Rows().OrderBy(a => a.ID))
                    {
                        if (m.日 == 0)
                        {
                            continue;
                        }

                        g.Rows.Add();
                        g[colDate, g.Rows.Count - 1].Value = m.暦年 + "/" + m.暦月.ToString("D2") + "/" + m.日.ToString("D2");
                        g[colStaffCode, g.Rows.Count - 1].Value = t.スタッフコード.ToString("D5");
                        g[colStaffName, g.Rows.Count - 1].Value = t.氏名;
                        g[colAreaMana, g.Rows.Count - 1].Value = t.担当エリアマネージャー名;
                        g[colAreaCode, g.Rows.Count - 1].Value = t.エリアコード.ToString("D3");
                        g[colAreaName, g.Rows.Count - 1].Value = t.エリア名;
                        g[colBushoCode, g.Rows.Count - 1].Value = t.店舗コード.ToString("D5");
                        g[colBushoName, g.Rows.Count - 1].Value = t.店舗名;

                        if (t.給与形態.ToString() == global.NIKKYU)
                        {
                            g[colkbn, g.Rows.Count - 1].Value = "日給";
                        }
                        else if (t.給与形態.ToString() == global.JIKKYU)
                        {
                            g[colkbn, g.Rows.Count - 1].Value = "時給";
                        }

                        g[colStatus, g.Rows.Count - 1].Value = m.出勤状況;

                        // 有休または公休のとき
                        if (m.出勤状況 == global.STATUS_KOUKYU || m.出勤状況 == global.STATUS_YUKYU)
                        {
                            g[colShhmm, g.Rows.Count - 1].Value = string.Empty;
                            g[colEhhmm, g.Rows.Count - 1].Value = string.Empty;
                            g[colRest, g.Rows.Count - 1].Value = string.Empty; ;
                            g[colDayShopCode, g.Rows.Count - 1].Value = string.Empty;
                            g[colDayShopName, g.Rows.Count - 1].Value = string.Empty;
                        }
                        else
                        {
                            // 基本時間帯就労のとき
                            if (m.出勤状況 == global.STATUS_KIHON_1 || m.出勤状況 == global.STATUS_KIHON_2 || m.出勤状況 == global.STATUS_KIHON_3)
                            {
                                foreach (var item in stf)
                                {
                                    if (item.スタッフコード == t.スタッフコード)
                                    {
                                        if (m.出勤状況 == global.STATUS_KIHON_1)
                                        {
                                            g[colShhmm, g.Rows.Count - 1].Value = item.基本時間帯1.Substring(0, 5);
                                            g[colEhhmm, g.Rows.Count - 1].Value = item.基本時間帯1.Substring(6, 5);
                                            break;
                                        }
                                        else if (m.出勤状況 == global.STATUS_KIHON_2)
                                        {
                                            g[colShhmm, g.Rows.Count - 1].Value = item.基本時間帯2.Substring(0, 5);
                                            g[colEhhmm, g.Rows.Count - 1].Value = item.基本時間帯2.Substring(6, 5);
                                            break;
                                        }
                                        else if (m.出勤状況 == global.STATUS_KIHON_3)
                                        {
                                            g[colShhmm, g.Rows.Count - 1].Value = item.基本時間帯3.Substring(0, 5);
                                            g[colEhhmm, g.Rows.Count - 1].Value = item.基本時間帯3.Substring(6, 5);
                                            break;
                                        }
                                    }
                                }

                                // 基本就業時間帯のとき休憩は60分固定
                                g[colRest, g.Rows.Count - 1].Value = 60;
                            }
                            else
                            {
                                if (m.出勤時 == string.Empty && m.出勤分 == string.Empty)
                                {
                                    g[colShhmm, g.Rows.Count - 1].Value = string.Empty;
                                }
                                else
                                {
                                    g[colShhmm, g.Rows.Count - 1].Value = string.Format("{0, 2}", m.出勤時) + ":" + m.出勤分.PadLeft(2, '0');
                                }

                                if (m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                                {
                                    g[colEhhmm, g.Rows.Count - 1].Value = string.Empty;
                                }
                                else
                                {
                                    g[colEhhmm, g.Rows.Count - 1].Value = string.Format("{0, 2}", m.退勤時) + ":" + m.退勤分.PadLeft(2, '0');
                                }

                                g[colRest, g.Rows.Count - 1].Value = m.休憩;
                            }

                            // 就労店舗
                            if (m.店舗コード != global.flgOff)
                            {
                                g[colDayShopCode, g.Rows.Count - 1].Value = m.店舗コード.ToString("D5");

                                // 就労店舗名取得
                                foreach (var item in shp)
                                {
                                    if (item.店舗コード == m.店舗コード)
                                    {
                                        g[colDayShopName, g.Rows.Count - 1].Value = item.店舗名;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                g[colDayShopCode, g.Rows.Count - 1].Value = string.Empty;
                                g[colDayShopName, g.Rows.Count - 1].Value = string.Empty;
                            }
                        }
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
            int sArea = 0;

            if (cmbBumonS.SelectedIndex != -1)
            {
                Utility.comboArea cmb = (Utility.comboArea)cmbBumonS.SelectedItem;
                sArea = cmb.Code;
            }

            // 年月範囲
            int sYY = 0;
            int sYYto = 0;

            if (txtYear.Text == string.Empty)
            {
                sYY = 1;
            }
            else
            {
                sYY = Utility.StrtoInt(txtYear.Text) - 2000;
            }

            int sMM = Utility.StrtoInt(txtMonth.Text);

            if (txtYearTo.Text == string.Empty)
            {
                sYYto = 99;
            }
            else
            {
                sYYto = Utility.StrtoInt(txtYearTo.Text) - 2000;
            }

            int sMMto = Utility.StrtoInt(txtMonthTo.Text);

            //データ表示
            GridViewShowData(dg1, sYY, sMM, sYYto, sMMto, sArea, txtsName.Text);
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
            MyLibrary.CsvOut.GridView(dg1, "勤怠一覧");
        }

    }
}
