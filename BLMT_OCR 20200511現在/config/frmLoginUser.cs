using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;

namespace BLMT_OCR.config
{
    public partial class frmLoginUser : Form
    {
        public frmLoginUser()
        {
            InitializeComponent();

            // データ読み込み
            uAdp.Fill(dts.ログインユーザー);
        }

        // データセット・テーブルアダプタ
        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.ログインユーザーTableAdapter uAdp = new DataSet1TableAdapters.ログインユーザーTableAdapter();

        // フォームモード
        Utility.frmMode fMode = new Utility.frmMode();

        #region グリッドビューカラム定義
        string colID = "col1";
        string colAccount = "col2";
        string colName = "col7";
        string colType = "col6";
        string colMemo = "col3";
        string colAddDt = "col4";
        string colUpDt = "col5";
        #endregion

        // 管理者アカウント
        const string ADMINUSER = "administrator";

        private void frmLoginType_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ 
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // データグリッドビューの定義
            gridSetting(dataGridView1);

            // データグリッドビューデータ表示
            gridShow(dataGridView1);

            //// 受注データ保守コンボ
            //cmbOrderSet();

            // 画面初期化
            dispClear();
        }


        /// -------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// -------------------------------------------------------------------
        private void gridSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 181;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(colID, "ｺｰﾄﾞ");
                tempDGV.Columns.Add(colAccount, "ユーザーアカウント");
                tempDGV.Columns.Add(colName, "名前");
                tempDGV.Columns.Add(colMemo, "備考");
                tempDGV.Columns.Add(colUpDt, "更新年月日");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colAccount].Width = 140;
                tempDGV.Columns[colName].Width = 160;
                tempDGV.Columns[colMemo].Width = 160;
                tempDGV.Columns[colUpDt].Width = 140;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

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

        /// ------------------------------------------------------------------
        /// <summary>
        ///     グリッドデータを表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        private void gridShow(DataGridView g)
        {
            g.Rows.Clear();
            int iX = 0;

            // グリッドに表示
            foreach (var t in dts.ログインユーザー)
            {
                g.Rows.Add();
                g[colID, iX].Value = t.ID.ToString();
                g[colAccount, iX].Value = t.アカウント;
                g[colName, iX].Value = t.名前;
                g[colMemo, iX].Value = t.備考;
                g[colUpDt, iX].Value = t.更新年月日;

                iX++;
            }

            g.CurrentCell = null;
        }
        

        /// -------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        /// -------------------------------------------------------------
        private void dispClear()
        {
            fMode.Mode = 0;
            fMode.ID = global.flgOff;
            txtAccount.Text = string.Empty;
            txtAccount.Enabled = true;
            txtPassword.Text = string.Empty;
            txtPassword.Enabled = true;
            txtPassword2.Text = string.Empty;
            txtPassword2.Enabled = true;
            txtName.Text = string.Empty;
            txtMemo.Text = string.Empty;
            //checkBox1.Checked = false;
            checkBox1.Visible = false;

            button2.Enabled = false;
            button4.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // エラーチェック
            if (!errCheck(fMode.Mode))
            {
                return;
            }

            // 確認メッセージ
            if (fMode.Mode == global.FORM_ADDMODE)
            {
                if (MessageBox.Show("ユーザーアカウントを新規登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            else if (fMode.Mode == global.FORM_EDITMODE)
            {
                if (MessageBox.Show("ユーザーアカウントを更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            // 登録・更新処理
            dataUpdate(fMode.Mode, fMode.ID);

            // グリッド表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     登録時エラーチェック </summary>
        /// <param name="sMode">
        ///     処理モード</param>
        /// <returns>
        ///     エラーなし:true、エラーあり:false</returns>
        /// -------------------------------------------------------------------------
        private bool errCheck(int sMode)
        {
            // 新規登録のとき
            if (sMode == global.FORM_ADDMODE)
            {
                // IDチェック
                if (txtAccount.Text == string.Empty)
                {
                    MessageBox.Show("ユーザーアカウントが未入力です", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAccount.Focus();
                    return false;
                }

                if (txtAccount.Text.Length < 6 || txtAccount.Text.Length > 12)
                {
                    MessageBox.Show("ユーザーアカウントは6～12文字で登録して下さい", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAccount.Focus();
                    return false;
                }
                
                if (dts.ログインユーザー.Any(a => a.アカウント.Equals(txtAccount.Text)))
                {
                    MessageBox.Show("既に登録済みのユーザーアカウントです", "ユーザーアカウントエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAccount.Focus();
                    return false;
                }
            }

            // パスワード
            if (txtPassword.Enabled)
            {
                // adminのとき
                if (txtAccount.Text == global.ADMIN_USER)
                {
                    MessageBox.Show(global.ADMIN_USER + "のパスワードは変更できません", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    checkBox1.Checked = false;
                    return false;
                }

                if (txtPassword.Text.Trim() == string.Empty)
                {
                    MessageBox.Show("パスワードを入力して下さい", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPassword.Focus();
                    return false;
                }

                if (txtPassword.Text.Length < 8 || txtPassword.Text.Length > 12)
                {
                    MessageBox.Show("パスワードは8～12文字で登録して下さい", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPassword.Focus();
                    return false;
                }
                
                if (!txtPassword.Text.Equals(txtPassword2.Text))
                {
                    MessageBox.Show("パスワードと再入力パスワードが一致しません", "パスワードエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtPassword.Focus();
                    return false;
                }
            }
            
            return true;
        }
        
        /// -------------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ、タグデータ登録 </summary>
        /// <param name="sMode">
        ///     処理モード</param>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// -------------------------------------------------------------------------
        private void dataUpdate(int sMode, int sID)
        {
            // 新規登録
            if (sMode == global.FORM_ADDMODE)
            {
                DataSet1.ログインユーザーRow r = dts.ログインユーザー.NewログインユーザーRow();
                r.アカウント = txtAccount.Text;
                r.パスワード = txtPassword.Text;
                r.名前 = txtName.Text;
                r.備考 = txtMemo.Text;
                r.更新年月日 = DateTime.Now;
                dts.ログインユーザー.AddログインユーザーRow(r);                
            }
            else if (sMode == global.FORM_EDITMODE)    // 更新
            {
                DataSet1.ログインユーザーRow r = dts.ログインユーザー.Single(a => a.ID == sID);

                // パスワード変更のとき
                if (txtPassword.Visible)
                {
                    r.パスワード = txtPassword.Text;
                }

                r.名前 = txtName.Text;
                r.備考 = txtMemo.Text;
                r.更新年月日 = DateTime.Now;
            }
            
            // データベース更新
            uAdp.Update(dts.ログインユーザー);

            // データ読み込み
            uAdp.Fill(dts.ログインユーザー);
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmLoginType_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 画面初期化
            dispClear();
        }
        
        /// ----------------------------------------------------------------------
        /// <summary>
        ///     ログインタイプヘッダ、タグデータ表示 </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// ----------------------------------------------------------------------
        private void getData(int sID)
        {
            // ログインタイプヘッダ
            DataSet1.ログインユーザーRow r = dts.ログインユーザー.Single(a => a.ID == sID);

            txtAccount.Text = r.アカウント;
            txtPassword.Text = r.パスワード;
            txtName.Text = r.名前;
            txtMemo.Text = r.備考;
            
            // 処理モード
            fMode.Mode = 1;
            fMode.ID = r.ID;

            // ユーザーアカウント、パスワードの編集は不可とします
            txtAccount.Enabled = false;

            // パスワード編集チェック
            checkBox1.Visible = true;
            txtPassword.Enabled = false;
            txtPassword2.Enabled = false;

            // 削除、取消ボタンの使用を可能とします
            button2.Enabled = true;
            button4.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // adminのとき
            if (txtAccount.Text == global.ADMIN_USER)
            {
                MessageBox.Show(global.ADMIN_USER + "は削除できません", "削除不可", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 確認メッセージ
            if (MessageBox.Show("表示中のユーザーアカウントを削除します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            { 
                return; 
            }

            // データ削除
            delData(fMode.ID);

            // グリッド表示
            gridShow(dataGridView1);

            // 画面初期化
            dispClear();
        }
        
        /// ----------------------------------------------------------------------
        /// <summary>
        ///     データ削除 </summary>
        /// <param name="sID">
        ///     ヘッダID</param>
        /// ----------------------------------------------------------------------
        private void delData(int sID)
        {
            // ヘッダデータ削除
            DataSet1.ログインユーザーRow hr = dts.ログインユーザー.Single(a => a.ID == sID);
            hr.Delete();
            
            // データベース更新
            uAdp.Update(dts.ログインユーザー);

            // データ読み込み
            uAdp.Fill(dts.ログインユーザー);
        }

        private void txtID_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;

            //txtObj.BackColor = Color.LightSteelBlue;
            txtObj.SelectAll();
        }

        private void txtID_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;

            //txtObj.BackColor = Color.White;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // administratorは編集不可
            if (dataGridView1[colName, dataGridView1.SelectedRows[0].Index].Value.ToString() == ADMINUSER)
            {
                MessageBox.Show("管理者アカウントの変更、削除は出来ません。", "管理者アカウント", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // IDを取得
            int sID = int.Parse(dataGridView1[colID, dataGridView1.SelectedRows[0].Index].Value.ToString());
            
            // データ表示
            getData(sID);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                txtPassword.Enabled = true;
                txtPassword2.Enabled = true;
            }
            else
            {
                txtPassword.Enabled = false;
                txtPassword2.Enabled = false;
            }
        }
    }
}
