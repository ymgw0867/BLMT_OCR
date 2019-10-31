using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;

namespace BLMT_OCR
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();

            // ログインステータス
            global.loginStatus = false;

            // ログインユーザーデータ読み込み
            adp.Fill(dts.ログインユーザー);
        }
        
        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.ログインユーザーTableAdapter adp = new DataSet1TableAdapters.ログインユーザーTableAdapter();

        // ログイントライ回数
        int lTry = 0;

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmLogin_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 認証結果検証
            if (!SerachloginUser())
            {
                // 3回ログインエラーでシステムは終了します
                if (lTry > 2)
                {
                    MessageBox.Show("3回続けてログインに失敗しました。システムを終了します。","中止",MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Environment.Exit(0);
                }
            }
            else
            {
                global.loginStatus = true;
                this.Close();
            }
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     ユーザー認証 </summary>
        /// <returns>
        ///     認証成功：true、認証失敗：false</returns>
        /// ---------------------------------------------------------
        private bool SerachloginUser()
        {
            // ログインユーザー
            if (txtAccount.Text == string.Empty)
            {
                MessageBox.Show("ログインアカウントを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                lTry++;
                txtAccount.Focus();
                return false;
            }

            // パスワード
            if (txtPassword.Text == string.Empty)
            {
                MessageBox.Show("パスワードを入力してください", "エラー", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                lTry++;
                txtPassword.Focus();
                return false;
            }

            // ユーザー認証
            if (!dts.ログインユーザー.Any(a => a.アカウント == txtAccount.Text && a.パスワード == txtPassword.Text))
            {
                lTry++;
                MessageBox.Show("アカウント認証に失敗しました。" + Environment.NewLine + "ログインアカウントとパスワードを確認してください。", "認証エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAccount.Focus();
                return false;
            }
            else
            {
                var s = dts.ログインユーザー.Single(a => a.アカウント == txtAccount.Text && a.パスワード == txtPassword.Text);
                global.loginUserID = s.ID;
            }

            // 認証成功
            return true;
        }

        private void txtName_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;
            txtObj.SelectAll();
            //txtObj.BackColor = SystemColors.ControlLight;
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;
            //txtObj.BackColor = Color.White;
        }
    }
}
