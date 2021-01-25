using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;
using System.Data.SqlClient;

namespace BLMT_OCR.OCR
{
    public partial class frmOCRIndex : Form
    {
        public frmOCRIndex(DataSet1 _dts, DataSet1TableAdapters.勤務票ヘッダTableAdapter _hAdp, clsStaff [] stf, clsShop [] _shp)
        {
            InitializeComponent();

            dts = _dts;
            hAdp = _hAdp;
            shp = _shp;
        }

        DataSet1 dts = null;
        DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();

        clsShop[] shp = null;
        
        private void frmOCRIndex_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 店舗コンボボックスロード
            Utility.comboShop.Load(comboBox1, shp);

            // データグリッドビュー定義
            GridViewSetting(dataGridView1);

            // データグリッドビュー表示
            GridViewShowData(dataGridView1, 0);
        }


        #region グリッドカラム定義
        string colDate = "c1";
        string colBushoCode = "c2";
        string colBushoName = "c3";
        string colNinzu = "c4";
        string colID = "c5";
        string colAreaCode = "c6";
        string colAreaName = "c7";
        string colStaffCode = "c8";
        string colStaffName = "c9";
        #endregion

        /// ------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更するe

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 10, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", (float)10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 554;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // カラム定義
                tempDGV.Columns.Add(colAreaCode, "エリア№");
                tempDGV.Columns.Add(colAreaName, "エリア名");
                tempDGV.Columns.Add(colBushoCode, "店舗№");
                tempDGV.Columns.Add(colBushoName, "店舗名");
                tempDGV.Columns.Add(colStaffCode, "コード");
                tempDGV.Columns.Add(colStaffName, "氏名");
                tempDGV.Columns.Add(colID, "ID");

                // IDは非表示
                tempDGV.Columns[colID].Visible = false;

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

                // 表示位置
                tempDGV.Columns[colAreaCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colAreaName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colBushoCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colBushoName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStaffCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colStaffName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                              
                // 各列幅指定
                tempDGV.Columns[colAreaCode].Width = 90;
                tempDGV.Columns[colAreaName].Width = 120;
                tempDGV.Columns[colBushoCode].Width = 80;
                tempDGV.Columns[colBushoName].Width = 220;
                tempDGV.Columns[colStaffCode].Width = 80;
                tempDGV.Columns[colStaffName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                // 編集可否
                tempDGV.ReadOnly = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GridViewShowData(DataGridView g, int shopCode)
        {
            try
            {
                g.Rows.Clear();
                int i = 0;

                var s = dts.勤務票ヘッダ.OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);

                // スタッフ№
                if (txtSNum.Text != string.Empty)
                {
                    s = s.Where(a => a.スタッフコード == Utility.StrtoInt(txtSNum.Text)).OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);
                }

                // 店舗検索
                if (shopCode != global.flgOff)
                {
                    s = s.Where(a => a.店舗コード == shopCode).OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード).ThenBy(a => a.スタッフコード);
                }

                foreach (DataSet1.勤務票ヘッダRow t in s)
                {
                    g.Rows.Add();
                    g[colAreaCode, i].Value = t.エリアコード.ToString("D3");
                    g[colAreaName, i].Value = Utility.NulltoStr(t.エリア名);
                    g[colBushoCode, i].Value = t.店舗コード.ToString("D5");
                    g[colBushoName, i].Value = Utility.NulltoStr(t.店舗名);
                    g[colStaffCode, i].Value = t.スタッフコード.ToString("D5");
                    g[colStaffName, i].Value = Utility.NulltoStr(t.氏名);
                    g[colID, i].Value = t.ID.ToString();

                    i++;
                }

                g.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }

            // 勤務票情報がないとき
            if (g.RowCount == 0)
            {
                MessageBox.Show("出勤簿が存在しませんでした", "データなし", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // ID初期化
            hdID = string.Empty;

            // 確認
            string msg = dataGridView1[colStaffCode, e.RowIndex].Value.ToString() + " " + dataGridView1[colStaffName, e.RowIndex].Value.ToString() + "さん が選択されました。よろしいですか？";
            if (MessageBox.Show(msg,"確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // ID取得
            hdID = dataGridView1[colID, e.RowIndex].Value.ToString();

            // 閉じる
            this.Close();
        }

        // 選択したデータのID
        public string hdID { get; set; }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOCRIndex_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getSelectData();
        }

        private void getSelectData()
        {
            int sShop = 0;
            
            if (comboBox1.SelectedIndex != -1)
            {
                Utility.comboShop cmb = (Utility.comboShop)comboBox1.SelectedItem;
                sShop = cmb.Code; 
            }

            // データ検索
            GridViewShowData(dataGridView1, sShop);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtSNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
