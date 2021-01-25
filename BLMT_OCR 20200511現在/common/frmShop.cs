using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace BLMT_OCR.common
{
    public partial class frmShop : Form
    {
        public frmShop(clsShop [] shp)
        {
            InitializeComponent();
            _shp = shp;
        }

        clsShop[] _shp = null;

        // カラム定義
        private string colCode = "c0";
        private string colName = "c1";
        private string colSCode = "c2";
        private string colSName = "c3";

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        private void GridviewSet(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(225, 243, 190);
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 502;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 253, 230);

                // 各列幅指定
                tempDGV.Columns.Add(colCode, "エリア番号");
                tempDGV.Columns.Add(colName, "エリア名");
                tempDGV.Columns.Add(colSCode, "店舗コード");
                tempDGV.Columns.Add(colSName, "店舗名");

                tempDGV.Columns[colCode].Width = 80;
                tempDGV.Columns[colName].Width = 100;
                tempDGV.Columns[colSCode].Width = 90;
                tempDGV.Columns[colSName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colSCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
  
                // 編集可否
                tempDGV.ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                
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

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void showNouhin(DataGridView g)
        {
            this.Cursor = Cursors.WaitCursor;

            g.Rows.Clear();

            int cnt = 0;

            var s = _shp.OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード);

            if (sName.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.店舗名.Contains(sName.Text)).OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード);
            }

            foreach (var t in s)
            {
                g.Rows.Add();
                g[colCode, cnt].Value = t.エリアコード.ToString();
                g[colName, cnt].Value = Utility.NulltoStr(t.エリア名);
                g[colSCode, cnt].Value = t.店舗コード.ToString("D5");
                g[colSName, cnt].Value = Utility.NulltoStr(t.店舗名);

                cnt++;
            }

            g.CurrentCell = null;

            this.Cursor = Cursors.Default;

            if (cnt == 0)
            {
                MessageBox.Show("該当する店舗はありませんでした", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void frmTodoke_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            GridviewSet(dataGridView1);

            _nouCode = null;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (dataGridView1.SelectedRows.Count == 0)
            //{
            //    return;
            //}

            //int r = dataGridView1.SelectedRows[0].Index;

            //_nouCode = dataGridView1[colNouCode, r].Value.ToString();

            //Close();
        }

        public string[] _nouCode;

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            getSelectRows();
        }

        private void getSelectRows()
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }

            int iX = 0;

            Array.Resize(ref _nouCode, iX + 1);
            _nouCode[iX] = dataGridView1[colSCode, dataGridView1.SelectedRows[0].Index].Value.ToString();
           
            Close();
        }

        private void frmTodoke_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.F10 && button3.Enabled)
            {
                getSelectRows();
            }

            if (e.KeyData == Keys.F12)
            {
                Close();
            }
        }

        private void btnS_Click(object sender, EventArgs e)
        {
            showNouhin(dataGridView1);
        }
    }
}
