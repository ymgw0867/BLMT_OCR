using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;
using System.Data.OleDb;

namespace BLMT_OCR.config
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();

            cAdp.Fill(dts.環境設定);
            uAdp.Fill(dts.ログインユーザー);
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.環境設定TableAdapter cAdp = new DataSet1TableAdapters.環境設定TableAdapter();
        DataSet1TableAdapters.ログインユーザーTableAdapter uAdp = new DataSet1TableAdapters.ログインユーザーTableAdapter();

        private void frmConfig_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            var s = dts.環境設定.Single(a => a.ID == global.configKEY);

            if (s.Is年Null())
            {
                txtYear.Text = string.Empty;
            }
            else
            {
                txtYear.Text = s.年.ToString();
            }

            if (s.Is月Null())
            {
                txtMonth.Text = string.Empty;
            }
            else
            {
                txtMonth.Text = s.月.ToString();
            }

            if (s.Is汎用データ出力先Null())
            {
                txtCsvPath.Text = string.Empty;
            }
            else
            {
                txtCsvPath.Text = s.汎用データ出力先;
            }

            if (s.Is社員マスターパスNull())
            {
                txtMstPath.Text = string.Empty;
            }
            else
            {
                txtMstPath.Text = s.社員マスターパス;
            }

            if (s.Isデータ保存月数Null())
            {
                txtDataSpan.Text = string.Empty;
            }
            else
            {
                txtDataSpan.Text = s.データ保存月数.ToString();
            }

            if (s.Is受信監視スパンNull())
            {
                txtRecSpan.Text = string.Empty;
            }
            else
            {
                txtRecSpan.Text = s.受信監視スパン.ToString();
            }
            
            // adminでログイン時のみログインユーザー登録可能とする
            btnUser.Visible = false;
            if (dts.ログインユーザー.Any(a => a.ID == global.loginUserID))
            {
                DataSet1.ログインユーザーRow ur = dts.ログインユーザー.Single(a => a.ID == global.loginUserID);

                if (ur.アカウント == global.ADMIN_USER)
                {
                    btnUser.Visible = true;
                }
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     フォルダダイアログ選択 </summary>
        /// <returns>
        ///     フォルダー名</returns>
        ///------------------------------------------------------------------------
        private string userFolderSelect()
        {
            string fName = string.Empty;

            //出力フォルダの選択ダイアログの表示
            // FolderBrowserDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            // ダイアログの説明を設定する
            folderBrowserDialog1.Description = "フォルダを選択してください";

            // ルートになる特殊フォルダを設定する (初期値 SpecialFolder.Desktop)
            folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.Desktop;

            // 初期選択するパスを設定する
            folderBrowserDialog1.SelectedPath = @"C:\BLMT_OCR";

            // [新しいフォルダ] ボタンを表示する (初期値 true)
            folderBrowserDialog1.ShowNewFolderButton = true;

            // ダイアログを表示し、戻り値が [OK] の場合は、選択したディレクトリを表示する
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                fName = folderBrowserDialog1.SelectedPath + @"\";
            }
            else
            {
                // 不要になった時点で破棄する
                folderBrowserDialog1.Dispose();
                return fName;
            }

            // 不要になった時点で破棄する
            folderBrowserDialog1.Dispose();

            return fName;
        }

        private string userFileSelect()
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = "エリア別社員マスターを選択してください";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excelファイル(*.xlsx)|*.xlsx|(*.xls)|*.xls|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "郵便番号CSV確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 画像保存先フォルダを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                //txtPath1.Text = sPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // データ更新
            DataUpdate();
        }

        private void DataUpdate()
        {
            // エラーチェック
            if (!errCheck())
            {
                return;
            }

            if (MessageBox.Show("データを更新してよろしいですか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;
            
            DataSet1.環境設定Row r = dts.環境設定.Single(a => a.ID == global.configKEY);

            r.年 = Utility.StrtoInt(txtYear.Text);
            r.月 = Utility.StrtoInt(txtMonth.Text);
            r.汎用データ出力先 = txtCsvPath.Text;
            r.社員マスターパス = txtMstPath.Text;
            r.データ保存月数 = Utility.StrtoInt(txtDataSpan.Text);
            r.受信監視スパン = Utility.StrtoInt(txtRecSpan.Text);
            r.編集アカウント = global.flgOff;
            r.更新年月日 = DateTime.Now;

            // データ更新
            cAdp.Update(r);

            //
            //global.cnfPath = r.受け渡しデータ作成パス;
            //global.cnfImgPath = r.画像保存先パス;
            //global.cnfLogPath = r.ログデータ出力先パス;
 
            // 終了
            this.Close();
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェック </summary>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// ------------------------------------------------------------------------------------
        private bool errCheck()
        {
            // 処理年月
            if (Utility.StrtoInt(txtYear.Text) == 0)
            {
                MessageBox.Show("処理年を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtYear.Text) < 2017)
            {
                MessageBox.Show("処理年が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtMonth.Text) == 0)
            {
                MessageBox.Show("処理月を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtMonth.Text) < 1 || Utility.StrtoInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("処理月が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            // 画像保存先パス
            if (txtMstPath.Text.Trim() == string.Empty)
            {
                MessageBox.Show("エリア別社員マスターのパスを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMstPath.Focus();
                return false;
            }

            if (!System.IO.File.Exists(txtMstPath.Text))
            {
                MessageBox.Show("指定したエリア別社員マスターは存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMstPath.Focus();
                return false;
            }
            
            // PCA給与X汎用データ出力先パス
            if (txtCsvPath.Text.Trim() == string.Empty)
            {
                MessageBox.Show("PCA給与X向けCSV出力先フォルダを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCsvPath.Focus();
                return false;
            }

            if (!System.IO.Directory.Exists(txtCsvPath.Text))
            {
                MessageBox.Show("指定したPCA給与X向けCSV出力先フォルダは存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCsvPath.Focus();
                return false;
            }

            // データ保存月数
            if (txtDataSpan.Text.Trim() == string.Empty)
            {
                MessageBox.Show("データ保存月数を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtDataSpan.Focus();
                return false;
            }

            // 受信監視スパン
            if (txtRecSpan.Text.Trim() == string.Empty)
            {
                MessageBox.Show("FAX受信監視の時間間隔（秒）を入力してください", "受信間隔エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtRecSpan.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtRecSpan.Text) < 120)
            {
                MessageBox.Show("FAX受信監視間隔（秒）は120以上で設定してください", "受信間隔エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtRecSpan.Focus();
                return false;
            }
            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                txtCsvPath.Text = sPath;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                //txtLogPath.Text = sPath;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // ファイルを選択する
            string sPath = userFileSelect();
            if (sPath != string.Empty)
            {
                txtMstPath.Text = sPath;
            }
        }

        private void btnUser_Click(object sender, EventArgs e)
        {
            Hide();
            config.frmLoginUser frm = new config.frmLoginUser();
            frm.ShowDialog();
            Show();
        }
    }
}
