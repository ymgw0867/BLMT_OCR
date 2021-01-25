using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using BLMT_OCR.common;


namespace BLMT_OCR.prePrint
{
    public partial class frmPrePrint : Form
    {
        public frmPrePrint()
        {
            InitializeComponent();
            adp.Fill(dts.環境設定);
        }
        
        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.環境設定TableAdapter adp = new DataSet1TableAdapters.環境設定TableAdapter();
        
        private void frmPrePrint_Load(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Today.AddMonths(1);

            txtSyear.Text = dt.Year.ToString();
            txtSmonth.Text = dt.Month.ToString();
            txtEyear.Text = dt.Year.ToString();
            txtEmonth.Text = dt.Month.ToString();

            // 締日コンボ
            cmbShimebi.Items.Add("すべて");
            cmbShimebi.Items.Add(global.SHIME_15);
            cmbShimebi.Items.Add(global.SHIME_20);
            cmbShimebi.SelectedIndex = 0;

            // XLSマスターを配列に取り込み
            getXlsToArray();

            // エリアコンボ値ロード
            setAreaComboSet();

            // グリッドビュー設定
            gridViewSet(dataGridView1);

            // ボタン初期状態
            linkAllOn.Enabled = false;
            linkAllOff.Enabled = false;
            btnPrn.Enabled = false;
            lblCnt.Visible = false;
        }

        clsXlsmst[] xlsArray = null;
        //string[] csvArray = null;

        System.Collections.ArrayList al = new System.Collections.ArrayList();

        // カラム定義
        private string colChk = "c0";
        private string colAreaCode = "c1";
        private string colAreaName = "c2";
        private string colShopCode = "c3";
        private string colShopName = "c4";
        private string colSCode = "c5";
        private string colSName = "c6";
        private string colTime1 = "c7";
        private string colTime2 = "c8";
        private string colTime3 = "c9";
        private string colShimebi = "c10";
        private string colKyuyoKbn = "c11";
        private string ColID = "c12";
        private string colWorkTime = "c13";

        ///-------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        ///-------------------------------------------------------------
        private void gridViewSet(DataGridView dg)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                dg.EnableHeadersVisualStyles = false;
                dg.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                dg.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列ヘッダー表示位置指定
                dg.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                dg.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                dg.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 10, FontStyle.Regular);

                // 行の高さ
                dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                dg.ColumnHeadersHeight = 20;
                dg.RowTemplate.Height = 22;

                // 全体の高さ
                dg.Height = 547;

                // 奇数行の色
                dg.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列指定
                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                chk.Name = colChk;
                dg.Columns.Add(chk);
                dg.Columns[colChk].HeaderText = "";

                dg.Columns.Add(colAreaCode, "エリアID");
                dg.Columns.Add(colAreaName, "エリア名");
                dg.Columns.Add(colShopCode, "店舗ID");
                dg.Columns.Add(colShopName, "店舗名");
                dg.Columns.Add(colSCode, "スタッフ№");
                dg.Columns.Add(colSName, "氏名");
                dg.Columns.Add(colTime1, "基本就業時間帯１");
                dg.Columns.Add(colTime2, "基本就業時間帯２");
                dg.Columns.Add(colTime3, "基本就業時間帯３");
                dg.Columns.Add(colShimebi, "締日");
                dg.Columns.Add(colKyuyoKbn, "給与区分");
                dg.Columns.Add(colWorkTime, "基本実労働時間");
                dg.Columns.Add(ColID, "ID");

                dg.Columns[ColID].Visible = false;

                dg.Columns[colChk].Width = 30;
                dg.Columns[colAreaCode].Width = 70;
                dg.Columns[colAreaName].Width = 110;
                dg.Columns[colShopCode].Width = 70;
                dg.Columns[colShopName].Width = 220;
                dg.Columns[colSCode].Width = 70;
                dg.Columns[colSName].Width = 160;
                dg.Columns[colTime1].Width = 120;
                dg.Columns[colTime2].Width = 120;
                dg.Columns[colTime3].Width = 120;
                dg.Columns[colShimebi].Width = 60;
                dg.Columns[colKyuyoKbn].Width = 70;
                dg.Columns[colWorkTime].Width = 100;

                //dg.Columns[colShopName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                dg.Columns[colChk].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colAreaCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colShopCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colSCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colTime1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colTime2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colTime3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colShimebi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colKyuyoKbn].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[colWorkTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                dg.ReadOnly = false;
                foreach (DataGridViewColumn item in dg.Columns)
                {
                    // チェックボックスのみ使用可
                    if (item.Name == colChk)
                    {
                        dg.Columns[item.Name].ReadOnly = false;
                    }
                    else
                    {
                        dg.Columns[item.Name].ReadOnly = true;
                    }
                }

                // 行ヘッダを表示しない
                dg.RowHeadersVisible = false;

                // 選択モード
                dg.SelectionMode = DataGridViewSelectionMode.CellSelect;
                dg.MultiSelect = false;

                // 追加行表示しない
                dg.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                dg.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                dg.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                dg.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                dg.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 罫線
                dg.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                dg.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtSyear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     エリア別社員マスター（Excel)シート情報配列に取り込む </summary>
        ///-----------------------------------------------------------------------------
        private void getXlsToArray()
        {
            var s = dts.環境設定.Single(a => a.ID == global.flgOn);

            if (!System.IO.File.Exists(s.社員マスターパス))
            {
                MessageBox.Show("エリア別社員マスター：" + s.社員マスターパス + "は登録されていません", "ファイル未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                {
                    return;
                }
            }

            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(s.社員マスターパス, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = null;

            try
            {
                for (int i = 1; i <= oXlsBook.Sheets.Count; i++)
                {
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[i];
                    oxlsSheet.Select(Type.Missing);
                    
                    Excel.Range rng = null;

                    Array.Resize(ref xlsArray, i);
                    xlsArray[i - 1] = new clsXlsmst();

                    rng = oxlsSheet.Range[oxlsSheet.Cells[4, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                    xlsArray[i - 1].xlsMst = (object[,])rng.Value2;
                    xlsArray[i - 1].xlsMstID = i;

                    //cmbArea.Items.Add(oXlsBook.Sheets[i].Name);
                }

                //cmbArea.SelectedIndex = 0;

                // 2次元配列から1次元配列に取り込み
                array2ToArrayList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "エクセル社員マスター読み取りエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmPrePrint_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     2次元配列からArrayListにデータをセットする </summary>
        ///----------------------------------------------------------------------
        private void array2ToArrayList()
        {
            int iR = 0;

            for (int i = 0; i < xlsArray.Length; i++)
            {
                for (int iX = 1; iX <= xlsArray[i].xlsMst.GetLength(0); iX++)
                {
                    // 26列ないときはエリア別社員マスターシートとみなさない
                    if (xlsArray[i].xlsMst.GetLength(1) < 26)
                    {
                        continue;
                    }

                    // 使用区分が１以外はネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 26]).Trim() != global.FLGON)
                    {
                        continue;
                    }

                    // エリア№が空白のときネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).Trim() == string.Empty)
                    {
                        continue;
                    }

                    StringBuilder sb = new StringBuilder();
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).PadLeft(3, '0')).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 4]).PadLeft(5, '0')).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 20]).PadLeft(5, '0')).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 2])).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 5])).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 21])).Append(",");
                    sb.Append(getTimeString(xlsArray, i, iX, 8)).Append(",");
                    sb.Append(getTimeString(xlsArray, i, iX, 12)).Append(",");
                    sb.Append(getTimeString(xlsArray, i, iX, 16)).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 23])).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 22])).Append(",");

                    // 実労働時間
                    string wH = string.Empty;
                    string wM = string.Empty;
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]) != string.Empty)
                    {
                        wH = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]).PadLeft(2, ' ');
                    }

                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]) != string.Empty)
                    {
                        wM = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]).PadLeft(2, '0');
                    }

                    if (wH != string.Empty || wM != string.Empty)
                    {
                        sb.Append(wH).Append(":").Append(wM);
                    }
                    else
                    {
                        sb.Append(string.Empty);
                    }

                    //Array.Resize(ref csvArray, iR + 1);
                    //csvArray[iR] = sb.ToString();

                    al.Add(sb.ToString());
                    iR++;
                }
            }

            al.Sort();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     エリアコンボボックス値ロード </summary>
        ///---------------------------------------------------------------
        private void setAreaComboSet()
        {
            cmbArea.Items.Add("すべてのエリア");

            string arName = string.Empty;

            foreach (var t in al)
            {
                string [] j = t.ToString().Split(',');
                if (arName != j[3])
                {
                    cmbArea.Items.Add(j[3]);
                    arName = j[3];
                }
            }

            cmbArea.SelectedIndex = 0;
        }


        ///----------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューにスタッフ情報を表示する </summary>
        /// <param name="dg">
        ///     DataGridViewオブジェクト</param>
        ///----------------------------------------------------------------------
        private void xlsGridShow(DataGridView dg)
        {
            int iR = 0;

            dg.Rows.Clear();

            foreach (var t in al)
            {
                string[] f = t.ToString().Split(',');

                // 締日指定
                if (cmbShimebi.SelectedIndex == 1)
                {
                    if (Utility.NulltoStr(f[9]).Trim() != global.SHIME_15)
                    {
                        continue;
                    }
                }

                if (cmbShimebi.SelectedIndex == 2)
                {
                    if (Utility.NulltoStr(f[9]).Trim() != global.SHIME_20)
                    {
                        continue;
                    }
                }

                // エリア指定
                if (cmbArea.SelectedIndex > 0)
                {
                    if (f[3] != cmbArea.SelectedItem.ToString())
                    {
                        continue;
                    }
                }

                dg.Rows.Add();
                dg[colChk, iR].Value = true;
                dg[colAreaCode, iR].Value = f[0];
                dg[colAreaName, iR].Value = f[3];
                dg[colShopCode, iR].Value = f[1];
                dg[colShopName, iR].Value = f[4];
                dg[colSCode, iR].Value = f[2];
                dg[colSName, iR].Value = f[5];
                dg[colTime1, iR].Value = f[6];
                dg[colTime2, iR].Value = f[7];
                dg[colTime3, iR].Value = f[8];
                dg[colShimebi, iR].Value = f[9];
                dg[colKyuyoKbn, iR].Value = f[10];
                dg[colWorkTime, iR].Value = f[11];

                iR++;
            }

            // 表示結果
            if (dg.Rows.Count > 0)
            {
                dg.CurrentCell = null;

                linkAllOn.Enabled = true;
                linkAllOff.Enabled = true;
                btnPrn.Enabled = true;
                lblCnt.Visible = true;
                lblCnt.Text = dg.Rows.Count.ToString() + "名が抽出されました";
            }
            else
            {
                linkAllOn.Enabled = false;
                linkAllOff.Enabled = false;
                btnPrn.Enabled = false;
                lblCnt.Visible = false;
            }
        }
        
        private void xlsGridShow_XXX(DataGridView dg)
        {
            dg.Rows.Clear();
            int iR = 0;

            for (int i = 0; i < xlsArray.Length; i++)
            {
                for (int iX = 1; iX <= xlsArray[i].xlsMst.GetLength(0); iX++)
                {
                    // 使用区分が１以外はネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 29]).Trim() != global.FLGON)
                    {
                        continue;
                    }

                    // エリア№が空白のときネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).Trim() == string.Empty)
                    {
                        continue;
                    }

                    // 締日指定
                    if (cmbShimebi.SelectedIndex == 1)
                    {
                        if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 26]).Trim() != global.SHIME_15)
                        {
                            continue;
                        }
                    }

                    if (cmbShimebi.SelectedIndex == 2)
                    {
                        if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 26]).Trim() != global.SHIME_20)
                        {
                            continue;
                        }
                    }

                    // エリア指定
                    if (cmbArea.SelectedIndex > 0)
                    {
                        if (xlsArray[i].xlsMstID != cmbArea.SelectedIndex)
                        {
                            continue;
                        }
                    }

                    dg.Rows.Add();
                    dg[colChk, iR].Value = true;
                    dg[colAreaCode, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).PadLeft(3, '0');
                    dg[colAreaName, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 2]);
                    dg[colShopCode, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 4]).PadLeft(5, '0');
                    dg[colShopName, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 5]);
                    dg[colSCode, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 23]).PadLeft(5, '0');
                    dg[colSName, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]);
                    dg[colTime1, iR].Value = getTimeString(xlsArray, i, iX, 8);
                    dg[colTime2, iR].Value = getTimeString(xlsArray, i, iX, 13);
                    dg[colTime3, iR].Value = getTimeString(xlsArray, i, iX, 18);
                    dg[colShimebi, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 26]);
                    dg[colKyuyoKbn, iR].Value = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]);

                    iR++;
                }
            }

            // 表示結果
            if (dg.Rows.Count > 0)
            {
                dg.CurrentCell = null;

                linkAllOn.Enabled = true;
                linkAllOff.Enabled = true;
                btnPrn.Enabled = true;
            }
            else
            {
                linkAllOn.Enabled = false;
                linkAllOff.Enabled = false;
                btnPrn.Enabled = false;
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     基本就業時間帯文字列取得 </summary>
        /// <param name="x">
        ///     Excelシート配列クラス</param>
        /// <param name="i">
        ///     Excelシート配列クラスインデックス</param>
        /// <param name="iX">
        ///     シート配列１次元インデックス</param>
        /// <param name="r">
        ///     シート配列２次元インデックス</param>
        /// <returns>
        ///     基本就業時間帯文字列</returns>
        ///-----------------------------------------------------------
        private string getTimeString(clsXlsmst[] x, int i, int iX, int r)
        {
            string rtn = string.Empty;

            string sHH = Utility.NulltoStr(x[i].xlsMst[iX, r]);
            string sMM = Utility.NulltoStr(x[i].xlsMst[iX, r + 1]);
            string eHH = Utility.NulltoStr(x[i].xlsMst[iX, r + 2]);
            string eMM = Utility.NulltoStr(x[i].xlsMst[iX, r + 3]);

            if (sHH != string.Empty || sMM != string.Empty || eHH != string.Empty || eMM != string.Empty)
            {
                rtn = sHH.PadLeft(2, ' ') + ":" + sMM.PadLeft(2, '0') + "～" + eHH.PadLeft(2, ' ') + ":" + eMM.PadLeft(2, '0');
            }

            return rtn;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            xlsGridShow(dataGridView1);
        }

        private void linkAllOn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("全てのスタッフを印刷対象とします。よろしいですか。", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[colChk, i].Value = true;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("全てのスタッフを印刷対象外とします。よろしいですか。", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[colChk, i].Value = false;
            }
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            int pCnt = 0;

            // 印刷範囲
            if (Utility.StrtoInt(txtSyear.Text) < 2017)
            {
                MessageBox.Show("印刷開始年が正しくありません", "印刷期間", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSyear.Focus();
                return;
            }

            if (Utility.StrtoInt(txtSmonth.Text) < 1 || Utility.StrtoInt(txtSmonth.Text) > 12)
            {
                MessageBox.Show("印刷開始月が正しくありません", "印刷期間", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSmonth.Focus();
                return;
            }

            if (Utility.StrtoInt(txtEyear.Text) < 2017)
            {
                MessageBox.Show("印刷終了年が正しくありません", "印刷期間", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtEyear.Focus();
                return;
            }

            if (Utility.StrtoInt(txtEmonth.Text) < 1 || Utility.StrtoInt(txtEmonth.Text) > 12)
            {
                MessageBox.Show("印刷終了月が正しくありません", "印刷期間", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtEmonth.Focus();
                return;
            }

            int sym = Utility.StrtoInt(txtSyear.Text) * 100 + Utility.StrtoInt(txtSmonth.Text);
            int eym = Utility.StrtoInt(txtEyear.Text) * 100 + Utility.StrtoInt(txtEmonth.Text);

            if (sym > eym)
            {
                MessageBox.Show("印刷する期間が正しくありません", "印刷期間", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtSyear.Focus();
                return;
            }
            
            // 印刷するスタッフ
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                // チェックされている部署を対象とする
                if (dataGridView1[colChk, r.Index].Value.ToString() == "True")
                {
                    pCnt++;
                }
            }

            if (pCnt == 0)
            {
                MessageBox.Show("印刷するスタッフを選択してください", "印刷対象選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show(pCnt + "名のスタッフの出勤簿を印刷します。よろしいですか。","印刷確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 出勤簿印刷
            prnSheet(dataGridView1);
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     出勤簿印刷処理 </summary>
        /// <param name="dg">
        ///     DataGridViewオブジェクト</param>
        ///----------------------------------------------------------------------
        private void prnSheet(DataGridView dg)
        {
            string kihonWorkTime = Properties.Settings.Default.基本実労時 + Properties.Settings.Default.基本実労分.PadLeft(2, '0');

            // 印刷開始月
            DateTime nDt;

            // 印刷終了月
            DateTime eDt = DateTime.Parse(txtEyear.Text + "/" + txtEmonth.Text + "/01");

            DateTime zDt;

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excel起動
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsFAX出勤簿, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
            Excel.Range[] aRng = new Excel.Range[3];

            int pCnt = 1;   // ページカウント
            object[,] rtnArray = null;

            try
            {
                // スタッフ
                foreach (DataGridViewRow r in dg.Rows)
                {
                    // チェックされているスタッフを対象とする
                    if (dataGridView1[colChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    nDt = DateTime.Parse(txtSyear.Text + "/" + txtSmonth.Text + "/01");

                    // 年月指定範囲でループさせる
                    while (true)
                    {
                        // テンプレートシートを追加する
                        pCnt++;
                        oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                        oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                        // シートのセルを一括して配列に取得します
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[1, 1], oxlsMsSheet.Cells[oxlsMsSheet.UsedRange.Rows.Count, oxlsMsSheet.UsedRange.Columns.Count]];
                        rtnArray = (object[,])rng.Value2;

                        // 年
                        rtnArray[3, 14] = nDt.Year.ToString().Substring(2, 1);
                        rtnArray[3, 20] = nDt.Year.ToString().Substring(3, 1);

                        // 月
                        rtnArray[3, 32] = nDt.Month.ToString().PadLeft(2, '0').Substring(0, 1);
                        rtnArray[3, 38] = nDt.Month.ToString().PadLeft(2, '0').Substring(1, 1);

                        // 氏名
                        rtnArray[5, 18] = dg[colSName, r.Index].Value.ToString();

                        // スタッフ№
                        rtnArray[6, 106] = dg[colSCode, r.Index].Value.ToString().Substring(0, 1);
                        rtnArray[6, 112] = dg[colSCode, r.Index].Value.ToString().Substring(1, 1);
                        rtnArray[6, 118] = dg[colSCode, r.Index].Value.ToString().Substring(2, 1);
                        rtnArray[6, 124] = dg[colSCode, r.Index].Value.ToString().Substring(3, 1);
                        rtnArray[6, 130] = dg[colSCode, r.Index].Value.ToString().Substring(4, 1);

                        // 日給者
                        if (dg[colKyuyoKbn, r.Index].Value.ToString() == global.NIKKYU)
                        {
                            // 基本労働時間
                            if (dg[colWorkTime, r.Index].Value.ToString() != string.Empty)
                            {
                                rtnArray[6, 169] = dg[colWorkTime, r.Index].Value.ToString().Substring(1, 1);
                                rtnArray[6, 175] = dg[colWorkTime, r.Index].Value.ToString().Substring(3, 1);
                                rtnArray[6, 181] = dg[colWorkTime, r.Index].Value.ToString().Substring(4, 1);
                            }

                            // 基本就業時間帯１
                            string tm = dg[colTime1, r.Index].Value.ToString();
                            if (tm.Length > 10)
                            {
                                rtnArray[18, 29] = tm.Substring(0, 1);
                                rtnArray[18, 35] = tm.Substring(1, 1);
                                rtnArray[18, 41] = tm.Substring(3, 1);
                                rtnArray[18, 47] = tm.Substring(4, 1);

                                rtnArray[18, 57] = tm.Substring(6, 1);
                                rtnArray[18, 63] = tm.Substring(7, 1);
                                rtnArray[18, 69] = tm.Substring(9, 1);
                                rtnArray[18, 75] = tm.Substring(10, 1);
                            }

                            // 基本就業時間帯２
                            tm = dg[colTime2, r.Index].Value.ToString();
                            if (tm.Length > 10)
                            {
                                rtnArray[18, 88] = tm.Substring(0, 1);
                                rtnArray[18, 94] = tm.Substring(1, 1);
                                rtnArray[18, 100] = tm.Substring(3, 1);
                                rtnArray[18, 106] = tm.Substring(4, 1);

                                rtnArray[18, 116] = tm.Substring(6, 1);
                                rtnArray[18, 122] = tm.Substring(7, 1);
                                rtnArray[18, 128] = tm.Substring(9, 1);
                                rtnArray[18, 134] = tm.Substring(10, 1);
                            }

                            // 基本就業時間帯３
                            tm = dg[colTime3, r.Index].Value.ToString();
                            if (tm.Length > 10)
                            {
                                rtnArray[18, 147] = tm.Substring(0, 1);
                                rtnArray[18, 153] = tm.Substring(1, 1);
                                rtnArray[18, 159] = tm.Substring(3, 1);
                                rtnArray[18, 165] = tm.Substring(4, 1);

                                rtnArray[18, 175] = tm.Substring(6, 1);
                                rtnArray[18, 181] = tm.Substring(7, 1);
                                rtnArray[18, 187] = tm.Substring(9, 1);
                                rtnArray[18, 193] = tm.Substring(10, 1);
                            }
                        }
                        else if (dg[colKyuyoKbn, r.Index].Value.ToString() == global.JIKKYU) // 時給者
                        {
                            // 基本労働時間
                            rtnArray[6, 169] = "*";
                            rtnArray[6, 175] = "*";
                            rtnArray[6, 181] = "*";

                            // 基本就業時間帯１                            
                            rtnArray[18, 29] = "*";
                            rtnArray[18, 35] = "*";
                            rtnArray[18, 41] = "*";
                            rtnArray[18, 47] = "*";

                            rtnArray[18, 57] = "*";
                            rtnArray[18, 63] = "*";
                            rtnArray[18, 69] = "*";
                            rtnArray[18, 75] = "*";

                            // 基本就業時間帯２
                            rtnArray[18, 88] = "*";
                            rtnArray[18, 94] = "*";
                            rtnArray[18, 100] = "*";
                            rtnArray[18, 106] = "*";

                            rtnArray[18, 116] = "*";
                            rtnArray[18, 122] = "*";
                            rtnArray[18, 128] = "*";
                            rtnArray[18, 134] = "*";

                            // 基本就業時間帯３
                            rtnArray[18, 147] = "*";
                            rtnArray[18, 153] = "*";
                            rtnArray[18, 159] = "*";
                            rtnArray[18, 165] = "*";

                            rtnArray[18, 175] = "*";
                            rtnArray[18, 181] = "*";
                            rtnArray[18, 187] = "*";
                            rtnArray[18, 193] = "*";
                        }

                        // 所属店名
                        rtnArray[11, 25] = dg[colShopName, r.Index].Value.ToString();

                        // 所属店コード
                        rtnArray[12, 169] = dg[colShopCode, r.Index].Value.ToString().Substring(0, 1);
                        rtnArray[12, 175] = dg[colShopCode, r.Index].Value.ToString().Substring(1, 1);
                        rtnArray[12, 181] = dg[colShopCode, r.Index].Value.ToString().Substring(2, 1);
                        rtnArray[12, 187] = dg[colShopCode, r.Index].Value.ToString().Substring(3, 1);
                        rtnArray[12, 193] = dg[colShopCode, r.Index].Value.ToString().Substring(4, 1);

                        // 開始月
                        zDt = nDt.AddMonths(-1);

                        // 締日
                        string shime = dg[colShimebi, r.Index].Value.ToString();
                        if (shime != global.SHIME_15 && shime != global.SHIME_20)
                        {
                            // 締日記入がないものは15日締とする
                            shime = global.SHIME_15;
                        }

                        // 締日による制御
                        if (shime == global.SHIME_15)
                        {
                            // 15日締
                            // 締切文言
                            rtnArray[25, 103] = "※ 締切 毎月16日～翌15日　※ 送信締切日：17日迄";

                            // 訂正欄月区切り罫線
                            aRng[0] = (Excel.Range)oxlsSheet.Cells[128, 111];
                            aRng[1] = (Excel.Range)oxlsSheet.Cells[130, 111];
                            oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;

                            // 訂正欄月表示
                            rtnArray[128, 57] = zDt.Month.ToString();
                            rtnArray[128, 63] = "月";
                            rtnArray[128, 147] = nDt.Month.ToString();
                            rtnArray[128, 153] = "月";
                        }
                        else
                        {
                            // 20日締
                            // 締切文言
                            rtnArray[25, 103] = "※ 締切 毎月21日～翌20日　※ 送信締切日：22日迄";

                            // 訂正欄月区切り罫線
                            aRng[0] = (Excel.Range)oxlsSheet.Cells[128, 81];
                            aRng[1] = (Excel.Range)oxlsSheet.Cells[130, 81];
                            oxlsSheet.get_Range(aRng[0], aRng[1]).Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;

                            // 訂正欄月表示
                            rtnArray[128, 39] = zDt.Month.ToString();
                            rtnArray[128, 45] = "月";
                            rtnArray[128, 135] = nDt.Month.ToString();
                            rtnArray[128, 141] = "月";
                        }

                        // 日付配列クラスを作成
                        clsDays cld = new clsDays();
                        setDayArray(cld, shime, zDt, nDt);
                        
                        // 日付をセット
                        setDayToSheet(rtnArray, cld);

                        // 下部日付欄をセット
                        setDayToSheet2(rtnArray, cld, oxlsSheet);

                        // 配列からシートセルに一括してデータをセットします
                        rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                        rng.Value2 = rtnArray;

                        // 年月範囲を超えたらループからぬける
                        nDt = nDt.AddMonths(1);
                        bool nxt = true;
                        switch (nDt.CompareTo(eDt))
                        {
                            case -1:    // 期限内
                                nxt = true;
                                break;
                            case 0:     // 期限日と同日
                                nxt = true;
                                break;
                            case 1:     // 期限日超過
                                nxt = false;
                                break;
                        }

                        if (!nxt)
                        {
                            break;
                        }
                    }
                }

                // 確認のためExcelのウィンドウを表示する
                //oXls.Visible = true;

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                //System.Threading.Thread.Sleep(1000);

                // 印刷
                oXlsBook.PrintOut();

                // 終了メッセージ 
                MessageBox.Show("終了しました");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            finally
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;

                // 保存処理
                oXls.DisplayAlerts = false;

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsMsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;
                oxlsMsSheet = null;

                GC.Collect();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     シートの訂正欄日付をセットする </summary>
        /// <param name="rtnArray">
        ///     日付配列</param>
        /// <param name="cld">
        ///     日付配列クラス</param>
        /// <param name="oxls">
        ///     アクティブなエクセルシート</param>
        ///--------------------------------------------------------------------
        private void setDayToSheet2(object[,] rtnArray, clsDays cld, Excel.Worksheet oxls)
        {
            // 下部日付欄をセット
            int sCol = 15;
            for (int i = 0; i < 31; i++)
            {
                if (cld.dl[i] != string.Empty)
                {
                    rtnArray[129, sCol] = Utility.StrtoInt(cld.dl[i].Substring(3, 2)).ToString();
                }
                else
                {
                    // 空白欄の背景色を変える
                    Excel.Range aRng = (Excel.Range)oxls.Cells[129, sCol];
                    oxls.get_Range(aRng, aRng).Interior.ColorIndex = 15;
                }

                sCol += 6;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     シートの日付欄をセットする </summary>
        /// <param name="rtnArray">
        ///     日付配列</param>
        /// <param name="cld">
        ///     日付配列クラス</param>
        ///--------------------------------------------------------------
        private void setDayToSheet(object[,] rtnArray, clsDays cld)
        {
            int sRow = 31;

            for (int i = 0; i < 31; i++)
            {
                if (i < 16)
                {
                    rtnArray[sRow, 5] = cld.dl[i];
                    rtnArray[sRow + 3, 5] = cld.dw[i];
                }
                else
                {
                    rtnArray[sRow, 103] = cld.dl[i];
                    rtnArray[sRow + 3, 103] = cld.dw[i];
                }

                if (i == 15)
                {
                    sRow = 31;
                }
                else
                {
                    sRow += 6;
                }
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     日付配列を作成 </summary>
        /// <param name="cld">
        ///     日付配列クラス</param>
        /// <param name="shime">
        ///     締日</param>
        /// <param name="zDt">
        ///     開始月の前月（勤怠記入開始日の月）</param>
        /// <param name="nDt">
        ///     開始月</param>
        ///-------------------------------------------------------------------
        private void setDayArray(clsDays cld, string shime, DateTime zDt, DateTime nDt)
        {
            int idl = 0;
            int iZ = 1;
            DateTime gDt;
            
            while (idl < 31)
            {
                int day = Utility.StrtoInt(shime) + iZ;

                if (DateTime.TryParse(zDt.Year.ToString() + "/" + zDt.Month.ToString() + "/" + day.ToString(), out gDt))
                {
                    cld.dl[idl] = gDt.Month.ToString().PadLeft(2, ' ') + "/" + gDt.Day.ToString().PadLeft(2, '0');
                    cld.dw[idl] = "（" + ("日月火水木金土").Substring(int.Parse(gDt.DayOfWeek.ToString("d")), 1) + "）";
                }
                else
                {
                    cld.dl[idl] = string.Empty;
                    cld.dw[idl] = string.Empty;
                }

                if (day == 31)
                {
                    iZ = 1;
                    shime = "";
                    zDt = nDt;
                }
                else
                {
                    iZ++;
                }

                idl++;
            }
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     日付配列クラス </summary>
        ///--------------------------------------------------------
        class clsDays
        {
            public string[] dl = new string[31];
            public string[] dw = new string[31];
        }
    }
}
