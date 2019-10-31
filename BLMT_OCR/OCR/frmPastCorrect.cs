using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using BLMT_OCR.common;
using BLMT_OCR.OCR;
using GrapeCity.Win.MultiRow;
using Excel = Microsoft.Office.Interop.Excel;

namespace BLMT_OCR.OCR
{
    public partial class frmPastCorrect : Form
    {
        /// ------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ </summary>
        /// <param name="sID">
        ///     過去勤務票ヘッダID</param>
        /// ------------------------------------------------------------
        public frmPastCorrect(string sID)
        {
            InitializeComponent();

            dID = sID;              // 処理モード

            /* テーブルアダプターマネージャーに勤務票ヘッダ、明細テーブル、
             * 過去勤務票ヘッダ、過去明細テーブルアダプターを割り付ける */
            pAdpMn.過去勤務票ヘッダTableAdapter = phAdp;
            pAdpMn.過去勤務票明細TableAdapter = piAdp;
            pAdpMn.過去データ編集ログTableAdapter = adp;

            // テーブル読み込み
            kAdp.Fill(dts.休日);
            adp.Fill(dts.過去データ編集ログ);
        }

        // データアダプターオブジェクト
        DataSet1TableAdapters.TableAdapterManager pAdpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter piAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.過去データ編集ログTableAdapter adp = new DataSet1TableAdapters.過去データ編集ログTableAdapter();
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();
        //DataSet1TableAdapters.環境設定TableAdapter cAdp = new DataSet1TableAdapters.環境設定TableAdapter();

        // データセットオブジェクト
        DataSet1 dts = new DataSet1();

        // セル値
        private string cellName = string.Empty;         // セル名
        private string cellBeforeValue = string.Empty;  // 編集前
        private string cellAfterValue = string.Empty;   // 編集後

        #region 編集ログ・項目名 2015/09/08
        private const string LOG_YEAR = "年";
        private const string LOG_MONTH = "月";
        private const string LOG_DAY = "日";
        private const string LOG_TAIKEICD = "体系コード";
        private const string CELL_TORIKESHI = "取消";
        private const string CELL_NUMBER = "社員番号";
        private const string CELL_KIGOU = "記号";
        private const string CELL_FUTSU = "普通残業・時";
        private const string CELL_FUTSU_M = "普通残業・分";
        private const string CELL_SHINYA = "深夜残業・時";
        private const string CELL_SHINYA_M = "深夜残業・分";
        private const string CELL_SHIGYO = "始業時刻・時";
        private const string CELL_SHIGYO_M = "始業時刻・分";
        private const string CELL_SHUUGYO = "終業時刻・時";
        private const string CELL_SHUUGYO_M = "終業時刻・分";
        #endregion 編集ログ・項目名

        // カレント社員情報
        //SCCSDataSet.社員所属Row cSR = null;
        
        // 社員マスターより取得した所属コード
        string mSzCode = string.Empty;

        #region 終了ステータス定数
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";
        const string END_NODATA = "non Data";
        #endregion

        string dID = string.Empty;                  // 表示する過去データのID
        string sDBNM = string.Empty;                // データベース名

        string _dbName = string.Empty;           // 会社領域データベース識別番号
        string _comNo = string.Empty;            // 会社番号
        string _comName = string.Empty;          // 会社名

        bool _eMode = true;

        // dataGridView1_CellEnterステータス
        bool gridViewCellEnterStatus = true;
        bool WorkTotalSumStatus = true;

        clsXlsmst[] xlsArray = null;

        // カレントデータRowsインデックス
        string [] cID = null;
        int cI = 0;

        // グローバルクラス
        global gl = new global();

        System.Collections.ArrayList al = new System.Collections.ArrayList();

        clsStaff[] stf = null;              // スタッフクラス配列
        clsStaff sStf = new clsStaff();     // 画面表示したスタッフクラス
        clsShop[] shp = null;               // 店舗マスタークラス
        clsEditLog[] eLog = null;           // 編集ログクラス
        int eLogCnt = 0;

        string _imgFile = string.Empty;     // 画像名

        bool editLogStatus = true;
        const string lblEDITMODE_E = "現在、編集モードです";
        const string lblEDITMODE_R = "現在、閲覧モードです";

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // Tabキーの既定のショートカットキーを解除する。
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow2.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow3.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Enter);
            gcMultiRow2.ShortcutKeyManager.Unregister(Keys.Enter);
            gcMultiRow3.ShortcutKeyManager.Unregister(Keys.Enter);

            // Tabキーのショートカットキーにユーザー定義のショートカットキーを割り当てる。
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow2.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow3.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            gcMultiRow2.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            gcMultiRow3.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);

            // スタッフ配列クラス作成
            Utility u = new Utility();
            u.getXlsToArray(global.cnfMsPath, xlsArray, ref stf);

            // 店舗マスター配列クラス作成
            u.getShopMaster(stf, ref shp);

            // データセットへデータを読み込みます
            getDataSet();   // 出勤簿

            // キャプション
            this.Text = "過去出勤簿表示";

            // GCMultiRow初期化
            gcMrSetting();

            // データ表示
            showOcrData(dID);
            
            gcMultiRow1.ReadOnly = true;
            gcMultiRow3.ReadOnly = true;

            // 閲覧モード
            lblEdit.Text = lblEDITMODE_R;
            lblEdit.ForeColor = Color.Blue;

            // tagを初期化
            this.Tag = string.Empty;

            // 現在の表示倍率を初期化
            gl.miMdlZoomRate = 0f;

            // 書換結果初期値
            reWhite = false;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     キー配列作成 </summary>
        ///-------------------------------------------------------------
        private void keyArrayCreate()
        {
            int iX = 0;
            foreach (var t in dts.勤務票ヘッダ.OrderBy(a => a.ID))
            {
                Array.Resize(ref cID, iX + 1);
                cID[iX] = t.ID;
                iX++;
            }
        }

        #region データグリッドビューカラム定義
        private static string cCheck = "col1";      // 取消
        private static string cShainNum = "col2";   // 社員番号
        private static string cName = "col3";       // 氏名
        private static string cKinmu = "col4";      // 勤務記号
        private static string cZH = "col5";         // 残業時
        private static string cZE = "col6";         // :
        private static string cZM = "col7";         // 残業分
        private static string cSIH = "col8";        // 深夜時
        private static string cSIE = "col9";        // :
        private static string cSIM = "col10";       // 深夜分
        private static string cSH = "col11";        // 開始時
        private static string cSE = "col12";        // :
        private static string cSM = "col13";        // 開始分
        private static string cEH = "col14";        // 終了時
        private static string cEE = "col15";        // :
        private static string cEM = "col16";        // 終了分
        //private static string cID = "colID";        // ID
        private static string cSzCode = "colSzCode";  // 所属コード
        private static string cSzName = "colSzName";  // 所属名

        #endregion

        private void gcMrSetting()
        {
            //multirow編集モード
            gcMultiRow3.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow3.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow3.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow3.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow3.RowCount = 1;                                  // 行数を設定
            this.gcMultiRow3.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする

            //multirow編集モード
            gcMultiRow2.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow2.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow2.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow2.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow2.RowCount = 1;                                  // 行数を設定
            this.gcMultiRow2.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする

            //multirow編集モード
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow1.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow1.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow1.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow1.RowCount = global.MAX_GYO;                     // 行数を設定
            this.gcMultiRow1.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void GetCsvDataToMDB()
        {
            // CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv");

            // CSVファイルがなければ終了
            if (inCsv.Length == 0) return;

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // OCRのCSVデータをMDBへ取り込む
            OCRData ocr = new OCRData(_dbName);
            ocr.CsvToMdb(Properties.Settings.Default.dataPath, frmP, _dbName);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //if (e.Control is DataGridViewTextBoxEditingControl)
            //{
            //    // 数字のみ入力可能とする
            //    if (dGV.CurrentCell.ColumnIndex != 0 && dGV.CurrentCell.ColumnIndex != 2)
            //    {
            //        //イベントハンドラが複数回追加されてしまうので最初に削除する
            //        e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            //        e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

            //        //イベントハンドラを追加する
            //        e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            //    }
            //}
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'a' && e.KeyChar <= 'z') ||
                e.KeyChar == '\b' || e.KeyChar == '\t')
                e.Handled = false;
            else e.Handled = true;
        }

        void Control_KeyPress3(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '0' && e.KeyChar != '5' && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void Control_KeyDownShop(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Space)
            {
                gcMultiRow1.EndEdit();

                frmShop frm = new frmShop(shp);
                frm.ShowDialog();

                if (frm._nouCode != null)
                {
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCell.RowIndex, gcMultiRow1.CurrentCellPosition.CellName, frm._nouCode[0]);

                    if (gcMultiRow1.CurrentCellPosition.CellName == "txtShopCode")
                    {
                        gcMultiRow1.CurrentCell = null;
                    }
                }

                // 後片付け
                frm.Dispose();
            }
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty)
            {
                btnRtn.Focus();
            }
        }

        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        ///-----------------------------------------------------------------------------------
        /// <summary>
        ///     カレントデータを更新する</summary>
        /// <param name="iX">
        ///     カレントレコードのインデックス</param>
        ///-----------------------------------------------------------------------------------
        private void CurDataUpDate(string iX)
        {
            // エラーメッセージ
            string errMsg = "過去出勤簿テーブル更新";

            try
            {
                // 過去勤務票ヘッダテーブル行を取得
                DataSet1.過去勤務票ヘッダRow r = dts.過去勤務票ヘッダ.Single(a => a.ID == iX);

                // 勤務票ヘッダテーブルセット更新
                //r.年 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtYear"].Value));
                //r.月 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtMonth"].Value));
                //r.スタッフコード = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtSNum"].Value));
                //r.氏名 = Utility.NulltoStr(gcMultiRow2[0, "lblName"].Value);

                //if (sStf != null)
                //{
                //    r.担当エリアマネージャー名 = sStf.エリアマネージャー名;
                //    r.エリアコード = sStf.エリアコード;
                //    r.エリア名 = sStf.エリア名;
                //    r.給与形態 = sStf.給与区分;
                //}
                //else
                //{
                //    r.担当エリアマネージャー名 = string.Empty;
                //    r.エリアコード = 0;
                //    r.エリア名 = string.Empty;
                //    r.給与形態 = 0;
                //}

                //r.店舗コード = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "lblShopCode"].Value));
                //r.店舗名 = Utility.NulltoStr(gcMultiRow2[0, "lblShopName"].Value);

                r.基本就業時間帯1開始時 = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Sh"].Value);
                r.基本就業時間帯1開始分 = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Sm"].Value);
                r.基本就業時間帯1終了時 = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Eh"].Value);
                r.基本就業時間帯1終了分 = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Em"].Value);

                r.基本就業時間帯2開始時 = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Sh"].Value);
                r.基本就業時間帯2開始分 = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Sm"].Value);
                r.基本就業時間帯2終了時 = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Eh"].Value);
                r.基本就業時間帯2終了分 = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Em"].Value);

                r.基本就業時間帯3開始時 = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Sh"].Value);
                r.基本就業時間帯3開始分 = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Sm"].Value);
                r.基本就業時間帯3終了時 = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Eh"].Value);
                r.基本就業時間帯3終了分 = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Em"].Value);

                r.要出勤日数 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtTlDays"].Value));
                r.実労日数 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtZdays"].Value));
                r.公休日数 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtKdays"].Value));
                r.有休日数 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtYuDays"].Value));

                r.遅早時間時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtChisou"].Value));
                r.遅早時間分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtChisou2"].Value));

                r.実労働時間時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtWorkTime"].Value));
                r.実労働時間分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtWorkTime2"].Value));

                r.基本時間内残業時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtNaiZan"].Value));
                r.基本時間内残業分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtNaiZan2"].Value));

                r.割増残業時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtMashiZan"].Value));
                r.割増残業分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtMashiZan2"].Value));

                r._20時以降勤務時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txt20Zan"].Value));
                r._20時以降勤務分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txt20Zan2"].Value));

                r._22時以降勤務時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txt22Zan"].Value));
                r._22時以降勤務分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txt22Zan2"].Value));

                r.土日祝日労働時間時 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtHolZan"].Value));
                r.土日祝日労働時間分 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtHolZan2"].Value)); 

                r.交通費 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtKotuhi"].Value));
                r.その他支給 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtSonota"].Value));
                r.更新年月日 = DateTime.Now;

                //if (checkBox1.Checked)
                //{
                //    r.確認 = global.flgOn;
                //}
                //else
                //{
                //    r.確認 = global.flgOff;
                //}

                r.備考 = txtMemo.Text;


                // 過去勤務票明細テーブルセット更新
                for (int i = 0; i < gcMultiRow1.RowCount; i++)
                {
                    int sID = int.Parse(gcMultiRow1[i, "txtID"].Value.ToString());
                    DataSet1.過去勤務票明細Row m = (DataSet1.過去勤務票明細Row)dts.過去勤務票明細.FindByID(sID);

                    // 無効な日付
                    if (Utility.NulltoStr(gcMultiRow1[i, "lblDay"].Value) == string.Empty)
                    {
                        m.日 = global.flgOff;
                        //m.出勤状況 = string.Empty;
                        //m.出勤時 = string.Empty;
                        //m.出勤分 = string.Empty;
                        //m.退勤時 = string.Empty;
                        //m.退勤分 = string.Empty;
                        //m.休憩 = string.Empty;
                        //m.有給申請 = global.flgOff;
                        //m.店舗コード = global.flgOff;
                        //m.編集アカウント = global.loginUserID;
                        //m.更新年月日 = DateTime.Now;
                        //m.暦年 = global.flgOff;
                        //m.暦月 = global.flgOff;

                        continue;
                    }

                    m.日 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "lblDay"].Value));
                    m.出勤状況 = Utility.NulltoStr(gcMultiRow1[i, "txtStatus"].Value);
                    m.出勤時 = Utility.NulltoStr(gcMultiRow1[i, "txtSh"].Value);
                    m.出勤分 = Utility.NulltoStr(gcMultiRow1[i, "txtSm"].Value);
                    m.退勤時 = Utility.NulltoStr(gcMultiRow1[i, "txtEh"].Value);
                    m.退勤分 = Utility.NulltoStr(gcMultiRow1[i, "txtEm"].Value);
                    m.休憩 = Utility.NulltoStr(gcMultiRow1[i, "txtRest"].Value);

                    if (gcMultiRow1[i, "chkZan"].Value.ToString() == "True")
                    {
                        m.有給申請 = global.flgOn;
                    }
                    else
                    {
                        m.有給申請 = global.flgOff;
                    }

                    if (gcMultiRow1[i, "chkTeisei"].Value.ToString() == "True")
                    {
                        if (i == 0)
                        {
                            r.訂正1 = global.flgOn;
                        }

                        if (i == 1)
                        {
                            r.訂正2 = global.flgOn;
                        }

                        if (i == 2)
                        {
                            r.訂正3 = global.flgOn;
                        }

                        if (i == 3)
                        {
                            r.訂正4 = global.flgOn;
                        }

                        if (i == 4)
                        {
                            r.訂正5 = global.flgOn;
                        }

                        if (i == 5)
                        {
                            r.訂正6 = global.flgOn;
                        }

                        if (i == 6)
                        {
                            r.訂正7 = global.flgOn;
                        }

                        if (i == 7)
                        {
                            r.訂正8 = global.flgOn;
                        }

                        if (i == 8)
                        {
                            r.訂正9 = global.flgOn;
                        }

                        if (i == 9)
                        {
                            r.訂正10 = global.flgOn;
                        }

                        if (i == 10)
                        {
                            r.訂正11 = global.flgOn;
                        }

                        if (i == 11)
                        {
                            r.訂正12 = global.flgOn;
                        }

                        if (i == 12)
                        {
                            r.訂正13 = global.flgOn;
                        }

                        if (i == 13)
                        {
                            r.訂正14 = global.flgOn;
                        }

                        if (i == 14)
                        {
                            r.訂正15 = global.flgOn;
                        }

                        if (i == 15)
                        {
                            r.訂正16 = global.flgOn;
                        }

                        if (i == 16)
                        {
                            r.訂正17 = global.flgOn;
                        }

                        if (i == 17)
                        {
                            r.訂正18 = global.flgOn;
                        }

                        if (i == 18)
                        {
                            r.訂正19 = global.flgOn;
                        }

                        if (i == 19)
                        {
                            r.訂正20 = global.flgOn;
                        }

                        if (i == 20)
                        {
                            r.訂正21 = global.flgOn;
                        }

                        if (i == 21)
                        {
                            r.訂正22 = global.flgOn;
                        }

                        if (i == 22)
                        {
                            r.訂正23 = global.flgOn;
                        }

                        if (i == 23)
                        {
                            r.訂正24 = global.flgOn;
                        }

                        if (i == 24)
                        {
                            r.訂正25 = global.flgOn;
                        }

                        if (i == 25)
                        {
                            r.訂正26 = global.flgOn;
                        }

                        if (i == 26)
                        {
                            r.訂正27 = global.flgOn;
                        }

                        if (i == 27)
                        {
                            r.訂正28 = global.flgOn;
                        }

                        if (i == 28)
                        {
                            r.訂正29 = global.flgOn;
                        }

                        if (i == 29)
                        {
                            r.訂正30 = global.flgOn;
                        }

                        if (i == 30)
                        {
                            r.訂正31 = global.flgOn;
                        }
                    }
                    else
                    {

                        if (i == 0)
                        {
                            r.訂正1 = global.flgOff;
                        }

                        if (i == 1)
                        {
                            r.訂正2 = global.flgOff;
                        }

                        if (i == 2)
                        {
                            r.訂正3 = global.flgOff;
                        }

                        if (i == 3)
                        {
                            r.訂正4 = global.flgOff;
                        }

                        if (i == 4)
                        {
                            r.訂正5 = global.flgOff;
                        }

                        if (i == 5)
                        {
                            r.訂正6 = global.flgOff;
                        }

                        if (i == 6)
                        {
                            r.訂正7 = global.flgOff;
                        }

                        if (i == 7)
                        {
                            r.訂正8 = global.flgOff;
                        }

                        if (i == 8)
                        {
                            r.訂正9 = global.flgOff;
                        }

                        if (i == 9)
                        {
                            r.訂正10 = global.flgOff;
                        }

                        if (i == 10)
                        {
                            r.訂正11 = global.flgOff;
                        }

                        if (i == 11)
                        {
                            r.訂正12 = global.flgOff;
                        }

                        if (i == 12)
                        {
                            r.訂正13 = global.flgOff;
                        }

                        if (i == 13)
                        {
                            r.訂正14 = global.flgOff;
                        }

                        if (i == 14)
                        {
                            r.訂正15 = global.flgOff;
                        }

                        if (i == 15)
                        {
                            r.訂正16 = global.flgOff;
                        }

                        if (i == 16)
                        {
                            r.訂正17 = global.flgOff;
                        }

                        if (i == 17)
                        {
                            r.訂正18 = global.flgOff;
                        }

                        if (i == 18)
                        {
                            r.訂正19 = global.flgOff;
                        }

                        if (i == 19)
                        {
                            r.訂正20 = global.flgOff;
                        }

                        if (i == 20)
                        {
                            r.訂正21 = global.flgOff;
                        }

                        if (i == 21)
                        {
                            r.訂正22 = global.flgOff;
                        }

                        if (i == 22)
                        {
                            r.訂正23 = global.flgOff;
                        }

                        if (i == 23)
                        {
                            r.訂正24 = global.flgOff;
                        }

                        if (i == 24)
                        {
                            r.訂正25 = global.flgOff;
                        }

                        if (i == 25)
                        {
                            r.訂正26 = global.flgOff;
                        }

                        if (i == 26)
                        {
                            r.訂正27 = global.flgOff;
                        }

                        if (i == 27)
                        {
                            r.訂正28 = global.flgOff;
                        }

                        if (i == 28)
                        {
                            r.訂正29 = global.flgOff;
                        }

                        if (i == 29)
                        {
                            r.訂正30 = global.flgOff;
                        }

                        if (i == 30)
                        {
                            r.訂正31 = global.flgOff;
                        }
                    }

                    m.店舗コード = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtShopCode"].Value));
                    m.編集アカウント = global.loginUserID;

                    m.更新年月日 = DateTime.Now;
                    m.暦年 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtYear"].Value));
                    m.暦月 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtMonth"].Value));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、指定された文字数になるまで左側に０を埋めこみ、右寄せした文字列を返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <param name="len">
        ///     文字列の長さ</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeVal(object tm, int len)
        {
            string t = Utility.NulltoStr(tm);
            if (t != string.Empty) return t.PadLeft(len, '0');
            else return t;
        }

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、先頭文字が０のとき先頭文字を削除した文字列を返す　
        ///     先頭文字が０以外のときはそのまま返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeValH(object tm)
        {
            string t = Utility.NulltoStr(tm);

            if (t != string.Empty)
            {
                t = t.PadLeft(2, '0');
                if (t.Substring(0, 1) == "0")
                {
                    t = t.Substring(1, 1);
                }
            }

            return t;
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     Bool値を数値に変換する </summary>
        /// <param name="b">
        ///     True or False</param>
        /// <returns>
        ///     true:1, false:0</returns>
        /// ------------------------------------------------------------------------------------
        private int booltoFlg(string b)
        {
            if (b == "True") return global.flgOn;
            else return global.flgOff;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     エラーチェックボタン </summary>
        /// <param name="sender">
        ///     </param>
        /// <param name="e">
        ///     </param>
        ///-----------------------------------------------------------------
        private void btnErrCheck_Click(object sender, EventArgs e)
        {
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            ////カレントデータの更新
            //CurDataUpDate(cID[cI]);

            ////レコードの移動
            //cI = hScrollBar1.Value;
            //showOcrData(cI);
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            ////「受入データ作成終了」「勤務票データなし」以外での終了のとき
            //if (this.Tag.ToString() != END_MAKEDATA && this.Tag.ToString() != END_NODATA)
            //{
            //    //if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //    //{
            //    //    e.Cancel = true;
            //    //    return;
            //    //}

            //    //// カレントデータ更新
            //    //CurDataUpDate(dID);
            //}

            //// データベース更新
            //pAdpMn.UpdateAll(dts);

            // 解放する
            //this.Dispose();
        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックを実行する</summary>
        /// <param name="cIdx">
        ///     現在表示中の勤務票ヘッダデータインデックス</param>
        /// <param name="ocr">
        ///     OCRDATAクラスインスタンス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// -----------------------------------------------------------------------------------
        private bool getErrData(int cIdx, OCRPastData ocr)
        {
            //// カレントレコード更新
            //CurDataUpDate(dID);

            //// エラー番号初期化
            //ocr._errNumber = ocr.eNothing;

            //// エラーメッセージクリーン
            //ocr._errMsg = string.Empty;

            //// エラーチェック実行①:カレントレコードから最終レコードまで
            //if (!ocr.errCheckMain(cIdx, (dts.過去勤務票ヘッダ.Rows.Count - 1), this, dts, cID, sStf, stf))
            //{
            //    return false;
            //}

            //// エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            //if (cIdx > 0)
            //{
            //    if (!ocr.errCheckMain(0, (cIdx - 1), this, dts, cID, sStf, stf))
            //    {
            //        return false;
            //    }
            //}

            //// エラーなし
            //lblErrMsg.Text = string.Empty;

            return true;
        }

        /// ---------------------------------------------------------------------
        /// <summary>
        ///     MDBファイルを最適化する </summary>
        /// ---------------------------------------------------------------------
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();
                string OldDb = Properties.Settings.Default.mdbOlePath;
                string NewDb = Properties.Settings.Default.mdbPathTemp;

                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.mdbPath + global.MDBBACK);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBFILE, Properties.Settings.Default.mdbPath + global.MDBBACK);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBTEMP, Properties.Settings.Default.mdbPath + global.MDBFILE);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < gl.ZOOM_MAX)
            {
                leadImg.ScaleFactor += gl.ZOOM_STEP;
            }
            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;

            //if (dGV.RowCount == global.NIPPOU_TATE)
            //{
            //    global.miMdlZoomRate_TATE = (float)leadImg.ScaleFactor;
            //}
            //else if (dGV.RowCount == global.NIPPOU_YOKO)
            //{
            //    global.miMdlZoomRate_YOKO = (float)leadImg.ScaleFactor;
            //}
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > gl.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= gl.ZOOM_STEP;
            }
            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;

            //if (dGV.RowCount == global.NIPPOU_TATE)
            //{
            //    global.miMdlZoomRate_TATE = (float)leadImg.ScaleFactor;
            //}
            //else if (dGV.RowCount == global.NIPPOU_YOKO)
            //{
            //    global.miMdlZoomRate_YOKO = (float)leadImg.ScaleFactor;
            //}
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                //global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (gl.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= gl.ZOOM_RATE;
                }
                else
                {
                    leadImg.ScaleFactor *= gl.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                //global.pblImagePath = string.Empty;
            }
        }

        private void leadImg_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void leadImg_MouseMove(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }
        
        private void gcMultiRow1_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            //// 過去データ表示のときは終了
            //if (dID != string.Empty) return;

            // 休憩チェック
            if (e.CellName == "txtStatus" || e.CellName == "txtSh" || e.CellName == "txtSm" || 
                e.CellName == "txtEh" || e.CellName == "txtEm" || e.CellName == "txtRest")
            {
                // 実働日で
                if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtSh"].Value) != string.Empty || 
                    Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtSm"].Value) != string.Empty || 
                    Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtEh"].Value) != string.Empty || 
                    Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtEm"].Value) != string.Empty)
                {
                    int rst = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtRest"].Value));

                    // 休憩が60ではないとき
                    if (rst != 60)
                    {
                        gcMultiRow1[e.RowIndex, "txtRest"].Style.BackColor = Color.FromName("LightGray");
                    }
                    else
                    {
                        gcMultiRow1[e.RowIndex, "txtRest"].Style.BackColor = Color.Empty;
                    }
                }
                else
                {
                    // 訂正チェックではない行のとき背景色をEmptyとする
                    if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, "chkTeisei"].Value) != "True")
                    {
                        gcMultiRow1[e.RowIndex, "txtRest"].Style.BackColor = Color.Empty;
                    }
                }
            }

            // 訂正チェック
            if (e.CellName == "chkTeisei")
            {
                if (gcMultiRow1[e.RowIndex, e.CellName].Value.ToString() == "True")
                {
                    gcMultiRow1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromName("MistyRose");
                }
                else
                {
                    gcMultiRow1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromName("White");
                }
            }

            // 集計値再計算
            if (WorkTotalSumStatus)
            {
                getWorkTimeSection();
            }
        }

        private void gcMultiRow1_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyDown -= new KeyEventHandler(Control_KeyDownShop);

                // 数字のみ入力可能とする
                if (gcMultiRow1.CurrentCell.Name == "txtSh" || gcMultiRow1.CurrentCell.Name == "txtSm" || 
                    gcMultiRow1.CurrentCell.Name == "txtEh" || gcMultiRow1.CurrentCell.Name == "txtEm" || 
                    gcMultiRow1.CurrentCell.Name == "txtRest" || gcMultiRow1.CurrentCell.Name == "txtStatus" ||
                    gcMultiRow1.CurrentCell.Name == "txtShopCode")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDownShop);
                }
                
                // 店舗検索画面呼出
                if (gcMultiRow1.CurrentCell.Name == "txtShopCode")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyDown += new KeyEventHandler(Control_KeyDownShop);
                }
            }
        }

        private void gcMultiRow2_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

                // 数字のみ入力可能とする
                if (gcMultiRow2.CurrentCell.Name == "txtSNum" ||
                    gcMultiRow2.CurrentCell.Name == "txtYear" || gcMultiRow2.CurrentCell.Name == "txtMonth" ||
                    gcMultiRow2.CurrentCell.Name == "txtTm1Sh" || gcMultiRow2.CurrentCell.Name == "txtTm1Sm" || 
                    gcMultiRow2.CurrentCell.Name == "txtTm1Eh" || gcMultiRow2.CurrentCell.Name == "txtTm1Em" || 
                    gcMultiRow2.CurrentCell.Name == "txtTm2Sh" || gcMultiRow2.CurrentCell.Name == "txtTm2Sm" || 
                    gcMultiRow2.CurrentCell.Name == "txtTm2Eh" || gcMultiRow2.CurrentCell.Name == "txtTm2Em" || 
                    gcMultiRow2.CurrentCell.Name == "txtTm3Sh" || gcMultiRow2.CurrentCell.Name == "txtTm3Sm" || 
                    gcMultiRow2.CurrentCell.Name == "txtTm3Eh" || gcMultiRow2.CurrentCell.Name == "txtTm3Em")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }
            }
        }

        private void gcMultiRow2_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            
            // 社員番号のとき社員名を表示します
            if (e.CellName == "txtSNum")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 氏名を初期化
                gcMultiRow2[e.RowIndex, "lblName"].Value = string.Empty;

                if (Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtSNum"].Value) != string.Empty)
                {
                    // スタッフ情報取得
                    getStaffData(Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtSNum"].Value)), out sStf);

                    if (sStf != null)
                    {
                        gcMultiRow2.SetValue(e.RowIndex, "lblName", sStf.スタッフ名);
                        gcMultiRow2.SetValue(e.RowIndex, "lblShopCode", sStf.店舗コード);
                        gcMultiRow2.SetValue(e.RowIndex, "lblShopName", sStf.店舗名);

                        if (sStf.店舗区分 == 1)
                        {
                            gcMultiRow3.SetValue(e.RowIndex, "txtTenpoKbn", "直営店");
                        }
                        else if (sStf.店舗区分 == 2)
                        {
                            gcMultiRow3.SetValue(e.RowIndex, "txtTenpoKbn", "消化店");
                        }
                        else if (sStf.店舗区分 == 3)
                        {
                            gcMultiRow3.SetValue(e.RowIndex, "txtTenpoKbn", "本社");
                        }
                        else
                        {
                            gcMultiRow3.SetValue(e.RowIndex, "txtTenpoKbn", "");
                        }
                    }
                }
            }




            //// 過去データ表示のときは終了
            //if (dID != string.Empty) return;

            //// 社員番号のとき社員名を表示します
            //if (e.CellName == "txtSNum")
            //{
            //    // ChangeValueイベントを発生させない
            //    gl.ChangeValueStatus = false;

            //    // 氏名を初期化
            //    gcMultiRow2[e.RowIndex, "lblName"].Value = string.Empty;

            //    if (Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtSNum"].Value) != string.Empty)
            //    {
            //        // スタッフ情報取得
            //        getStaffData(Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtSNum"].Value)), out sStf);

            //        if (sStf != null)
            //        {
            //            gcMultiRow2.SetValue(e.RowIndex, "lblName", sStf.スタッフ名);
            //            gcMultiRow2.SetValue(e.RowIndex, "lblShopCode", sStf.店舗コード);
            //            gcMultiRow2.SetValue(e.RowIndex, "lblShopName", sStf.店舗名);

            //            if (sStf.実労時間.Length > 4)
            //            {
            //                gcMultiRow2.SetValue(e.RowIndex, "txtWh", sStf.実労時間.Substring(0, 2));
            //                gcMultiRow2.SetValue(e.RowIndex, "txtWm", sStf.実労時間.Substring(3, 2));
            //            }
            //            else
            //            {
            //                gcMultiRow2.SetValue(e.RowIndex, "txtWh", "");
            //                gcMultiRow2.SetValue(e.RowIndex, "txtWm", "");
            //            }

            //            // 日給者のとき
            //            if (sStf.給与区分.ToString() == global.NIKKYU)
            //            {
            //                // 基本実労働時間
            //                gcMultiRow2[0, "txtWh"].Selectable = true;
            //                gcMultiRow2[0, "txtWm"].Selectable = true;

            //                // 基本就業時間帯
            //                gcMultiRow2[0, "txtTm1Sh"].Selectable = true;
            //                gcMultiRow2[0, "txtTm1Sm"].Selectable = true;
            //                gcMultiRow2[0, "txtTm1Eh"].Selectable = true;
            //                gcMultiRow2[0, "txtTm1Em"].Selectable = true;

            //                gcMultiRow2[0, "txtTm2Sh"].Selectable = true;
            //                gcMultiRow2[0, "txtTm2Sm"].Selectable = true;
            //                gcMultiRow2[0, "txtTm2Eh"].Selectable = true;
            //                gcMultiRow2[0, "txtTm2Em"].Selectable = true;

            //                gcMultiRow2[0, "txtTm3Sh"].Selectable = true;
            //                gcMultiRow2[0, "txtTm3Sm"].Selectable = true;
            //                gcMultiRow2[0, "txtTm3Eh"].Selectable = true;
            //                gcMultiRow2[0, "txtTm3Em"].Selectable = true;
            //            }
            //            else if (sStf.給与区分.ToString() == global.JIKKYU)
            //            {
            //                // 時給者のとき
            //                // 基本実労働時間
            //                gcMultiRow2[0, "txtWh"].Selectable = false;
            //                gcMultiRow2[0, "txtWm"].Selectable = false;

            //                // 基本就業時間帯
            //                gcMultiRow2[0, "txtTm1Sh"].Selectable = false;
            //                gcMultiRow2[0, "txtTm1Sm"].Selectable = false;
            //                gcMultiRow2[0, "txtTm1Eh"].Selectable = false;
            //                gcMultiRow2[0, "txtTm1Em"].Selectable = false;

            //                gcMultiRow2[0, "txtTm2Sh"].Selectable = false;
            //                gcMultiRow2[0, "txtTm2Sm"].Selectable = false;
            //                gcMultiRow2[0, "txtTm2Eh"].Selectable = false;
            //                gcMultiRow2[0, "txtTm2Em"].Selectable = false;

            //                gcMultiRow2[0, "txtTm3Sh"].Selectable = false;
            //                gcMultiRow2[0, "txtTm3Sm"].Selectable = false;
            //                gcMultiRow2[0, "txtTm3Eh"].Selectable = false;
            //                gcMultiRow2[0, "txtTm3Em"].Selectable = false;
            //            }
            //        }
            //        else
            //        {
            //            // マスター未登録
            //            gcMultiRow2.SetValue(e.RowIndex, "lblName", "");
            //            gcMultiRow2.SetValue(e.RowIndex, "lblShopCode", "");
            //            gcMultiRow2.SetValue(e.RowIndex, "lblShopName", "");

            //            gcMultiRow2.SetValue(e.RowIndex, "txtWh", "");
            //            gcMultiRow2.SetValue(e.RowIndex, "txtWm", "");
            //        }
            //    }

            //    // ChangeValueイベントステータスをtrueに戻す
            //    gl.ChangeValueStatus = true;

            //    // 再表示：締日が異なることがあるため
            //    if (WorkTotalSumStatus)
            //    {
            //        //カレントデータの更新
            //        CurDataUpDate(dID);

            //        // 再表示
            //        showOcrData(dID);
            //    }
            //}

            //// 基本就労時間帯
            //if (e.CellName == "txtTm1Sh" || e.CellName == "txtTm1Sm" || e.CellName == "txtTm1Eh" || e.CellName == "txtTm1Em" || 
            //    e.CellName == "txtTm2Sh" || e.CellName == "txtTm2Sm" || e.CellName == "txtTm2Eh" || e.CellName == "txtTm2Em" || 
            //    e.CellName == "txtTm3Sh" || e.CellName == "txtTm3Sm" || e.CellName == "txtTm3Eh" || e.CellName == "txtTm3Em")
            //{
            //    // 再表示
            //    if (WorkTotalSumStatus)
            //    {
            //        //カレントデータの更新
            //        CurDataUpDate(dID);

            //        // 再表示
            //        showOcrData(dID);
            //    }
            //}
        }

        private void gcMultiRow2_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow2.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow2.BeginEdit(true);
            }
        }

        private void gcMultiRow1_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow1.BeginEdit(true);
            }
        }

        private void gcMultiRow1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = gcMultiRow1.CurrentCell.Name;            

            if (colName == "chkTeisei" || colName == "chkZan")
            {
                if (gcMultiRow1.IsCurrentCellDirty)
                {
                    gcMultiRow1.CommitEdit(DataErrorContexts.Commit);
                    gcMultiRow1.Refresh();
                }
            }
        }

        private void gcMultiRow1_CellLeave(object sender, CellEventArgs e)
        {
            //if (e.CellName == "txtSm" || e.CellName == "txtEm")
            //{
            //    gl.ChangeValueStatus = false;

            //    if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, e.CellName].Value) != string.Empty)
            //    {

            //        gcMultiRow1.SetValue(e.RowIndex, e.CellName, Utility.NulltoStr(gcMultiRow1[e.RowIndex, e.CellName].Value).PadLeft(2, '0'));
            //    }

            //    gl.ChangeValueStatus = true;
            //}
        }

        private void gcMultiRow1_CellContentClick(object sender, CellEventArgs e)
        {
            //if (gcMultiRow1[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
            //{
            //    return;
            //}

            //if (e.CellName == "btnCell")
            //{
            //    //カレントデータの更新
            //    CurDataUpDate(cID[cI]);
                
            //    int sMID = Utility.StrtoInt(gcMultiRow1[e.RowIndex, "txtID"].Value.ToString());

            //    if (dts.勤務票明細.Any(a => a.ID == sMID))
            //    {
            //        var s = dts.勤務票明細.Single(a => a.ID == sMID);
            //        string kID = s.帰宅後勤務ID;
            //        frmKitakugo frm = new frmKitakugo(_dbName, sMID, kID, hArray, bs, true);
            //        frm.ShowDialog();

            //        // 帰宅後勤務データ再読み込み
            //        tAdp.Fill(dts.帰宅後勤務);

            //        //// 勤務票明細再読み込み
            //        //adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);

            //        // データ再表示
            //        showOcrData(cI);
            //    }
            //}
        }
        
        ///-------------------------------------------------------------------
        /// <summary>
        ///     スタッフ配列クラスより任意のスタッフ情報を取得する </summary>
        /// <param name="sNum">
        ///     スタッフコード</param>
        /// <param name="_stf">
        ///     スタッフ配列クラス</param>
        /// <returns>
        ///     スタッフコードに該当：true、該当者なし：false</returns>
        ///-------------------------------------------------------------------
        //private bool getStaffData(int sNum, out clsStaff _stf)
        //{
        //    bool rtn = false;

        //    _stf = null;

        //    for (int i = 0; i < stf.Length; i++)
        //    {
        //        if (stf[i].スタッフコード == sNum)
        //        {
        //            _stf = stf[i];
        //            rtn = true;
        //            break;
        //        }
        //    }

        //    return rtn;
        //}

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void lnkLblClr_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void lnkLblDelete_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void gcMultiRow3_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);

                // 数字のみ入力可能とする
                if (gcMultiRow3.CurrentCell.Name == "txtZdays" || gcMultiRow3.CurrentCell.Name == "txtKdays" ||
                    gcMultiRow3.CurrentCell.Name == "txtTlDays" || gcMultiRow3.CurrentCell.Name == "txtYuDays" ||
                    gcMultiRow3.CurrentCell.Name == "txtWorkTime" || gcMultiRow3.CurrentCell.Name == "txtWorkTime2" ||
                    gcMultiRow3.CurrentCell.Name == "txtNaiZan" || gcMultiRow3.CurrentCell.Name == "txtNaiZan2" ||
                    gcMultiRow3.CurrentCell.Name == "txtMashiZan" || gcMultiRow3.CurrentCell.Name == "txtMashiZan2" || 
                    gcMultiRow3.CurrentCell.Name == "txt20Zan" || gcMultiRow3.CurrentCell.Name == "txt20Zan2" || 
                    gcMultiRow3.CurrentCell.Name == "txtHolZan" || gcMultiRow3.CurrentCell.Name == "txtHolZan2" || 
                    gcMultiRow3.CurrentCell.Name == "txtChisou" || gcMultiRow3.CurrentCell.Name == "txtChisou2" || 
                    gcMultiRow3.CurrentCell.Name == "txtSonota" || gcMultiRow3.CurrentCell.Name == "txtSonota2" || 
                    gcMultiRow3.CurrentCell.Name == "txtKotuhi" || gcMultiRow3.CurrentCell.Name == "txtKotuhi2")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }
            }
        }

        private void gcMultiRow3_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow3.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow3.BeginEdit(true);
            }
        }

        private void gcMultiRow1_CellLeave_1(object sender, CellEventArgs e)
        {
        }

        private void btnRtn_Click_1(object sender, EventArgs e)
        {
            if (chkEdit.Checked)
            {
                if (MessageBox.Show("過去データを書き換えてよろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    // エラーチェック～データ書き換え
                    errChkPastData();
                }
                else
                {
                    MessageBox.Show("データを更新しないで終了します", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // フォームを閉じる
                    reWhite = false;
                    this.Tag = END_BUTTON;
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("データを更新しないで終了します", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // フォームを閉じる
                reWhite = false;
                this.Tag = END_BUTTON;
                this.Close();
            }
        }

        private void editLogUpdate()
        {
            if (eLog != null)
            {
                foreach (var t in eLog)
                {
                    var s = dts.過去データ編集ログ.New過去データ編集ログRow();
                    s.日時 = t.upDt;
                    s.編集アカウント = t.aID;
                    s.出勤簿スタッフコード = t.stuffCode;
                    s.出勤簿氏名 = t.stuffName;
                    s.項目名 = t.fieldName;
                    s.編集前値 = t.beforeValue;
                    s.編集後値 = t.afterValue;
                    s.出勤簿年 = Utility.StrtoInt(t.sYear);
                    s.出勤簿月 = Utility.StrtoInt(t.sMonth);
                    s.出勤簿日 = t.sDay;

                    dts.過去データ編集ログ.Add過去データ編集ログRow(s);
                }
            }
        }
        
        private void errChkPastData()
        {
            // OCRDataクラス生成
            OCRPastData ocr = new OCRPastData();

            // カレントレコード更新
            CurDataUpDate(dID);

            // エラー番号初期化
            ocr._errNumber = ocr.eNothing;

            // エラーメッセージクリーン
            ocr._errMsg = string.Empty;

            // 勤務票ヘッダ行のコレクションを取得します
            DataSet1.過去勤務票ヘッダRow r = dts.過去勤務票ヘッダ.Single(a => a.ID == dID);

            // エラーチェックを実行する
            if (ocr.errCheckData(dts, r, stf))
            {
                // エラーがなかったとき
                lblErrMsg.Text = string.Empty;

                // 編集ログデータ書き込み
                editLogUpdate();

                // 書換結果
                reWhite = true;

                // データベース更新
                pAdpMn.UpdateAll(dts);

                // フォームを閉じる
                this.Tag = END_BUTTON;
                this.Close();
            }
            else
            {
                // データ表示
                showOcrData(dID);

                // エラー表示
                ErrShow(ocr);

                if (chkEdit.Checked)
                {
                    gcMultiRow1.ReadOnly = false;
                    gcMultiRow3.ReadOnly = false;
                    txtMemo.ReadOnly = false;
                }
            }
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            //// 非ログ書き込み状態とする：2015/09/25
            //editLogStatus = false;

            //// OCRDataクラス生成
            //OCRData ocr = new OCRData(_dbName);

            //// エラーチェックを実行
            //if (getErrData(cI, ocr))
            //{
            //    MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    gcMultiRow1.CurrentCell = null;
            //    gcMultiRow2.CurrentCell = null;
            //    gcMultiRow3.CurrentCell = null;

            //    // データ表示
            //    showOcrData(dID);
            //}
            //else
            //{
            //    // カレントインデックスをエラーありインデックスで更新
            //    cI = ocr._errHeaderIndex;

            //    // データ表示
            //    showOcrData(dID);

            //    // エラー表示
            //    ErrShow(ocr);
            //}
        }
        
        private void gcMultiRow3_CellValueChanged(object sender, CellEventArgs e)
        {            
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            if (e.CellName == "txtZdays" || e.CellName == "txtKdays" || e.CellName == "txtYuDays")
            {
                // 
                int tDays = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtZdays"].Value)) +
                            Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtYuDays"].Value));

                gcMultiRow3.SetValue(0, "txtTlDays", tDays);
            }
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkEdit_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEdit.Checked)
            {
                gcMultiRow1.ReadOnly = false;
                gcMultiRow3.ReadOnly = false;
                txtMemo.ReadOnly = false;
                lblEdit.Text = lblEDITMODE_E;
                lblEdit.ForeColor = Color.Red;
            }
            else
            {
                gcMultiRow1.ReadOnly = true;
                gcMultiRow3.ReadOnly = true;
                txtMemo.ReadOnly = true;
                lblEdit.Text = lblEDITMODE_R;
                lblEdit.ForeColor = Color.Blue;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("画像を印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
            {
                return;
            }

            // 印刷実行
            printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Image img;

            img = Image.FromFile(_imgFile);
            e.Graphics.DrawImage(img, 0, 0);
            e.HasMorePages = false;

            img.Dispose();  // 2019/01/15
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("勤怠データを印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
            {
                return;
            }

            dataPrint();
        }

        private void dataPrint()
        {
            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excel起動
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsKintaiSheet, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            //Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
            //Excel.Range[] aRng = new Excel.Range[3];

            //int pCnt = 1;   // ページカウント
            object[,] rtnArray = null;

            try
            {
                // シートのセルを一括して配列に取得します
                rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                rtnArray = (object[,])rng.Value2;

                rtnArray[2, 2] = "20" + Utility.NulltoStr(gcMultiRow2[0, "txtYear"].Value) + "年" + Utility.NulltoStr(gcMultiRow2[0, "txtMonth"].Value) + "月度";
                rtnArray[2, 5] = Utility.NulltoStr(gcMultiRow2[0, "txtSNum"].Value);
                rtnArray[2, 6] = Utility.NulltoStr(gcMultiRow2[0, "lblName"].Value);
                rtnArray[2, 8] = Utility.NulltoStr(gcMultiRow2[0, "lblShopCode"].Value);
                rtnArray[2, 9] = Utility.NulltoStr(gcMultiRow2[0, "lblShopName"].Value);

                DateTime dt = DateTime.Today;

                int mRow = 5;

                for (int i = 0; i < gcMultiRow1.RowCount; i++)
                {
                    if (!DateTime.TryParse(Utility.NulltoStr(gcMultiRow1[i, "txtYear"].Value) + "/" + Utility.NulltoStr(gcMultiRow1[i, "txtMonth"].Value) + "/" + Utility.NulltoStr(gcMultiRow1[i, "lblDay"].Value), out dt))
                    {
                        continue;
                    }

                    rtnArray[mRow, 2] = dt.ToShortDateString();
                    rtnArray[mRow, 3] = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                    string gStatus = Utility.NulltoStr(gcMultiRow1[i, "txtStatus"].Value);
                    rtnArray[mRow, 4] = gStatus;

                    if (gStatus == global.STATUS_KIHON_1)
                    {
                        rtnArray[mRow, 5] = "基本就業時間１";
                        rtnArray[mRow, 6] = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Sh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow2[0, "txtTm1Sm"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 7] = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Eh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow2[0, "txtTm1Em"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 8] = "60";
                    }
                    else if (gStatus == global.STATUS_KIHON_2)
                    {
                        rtnArray[mRow, 5] = "基本就業時間２";
                        rtnArray[mRow, 6] = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Sh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow2[0, "txtTm2Sm"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 7] = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Eh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow2[0, "txtTm2Em"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 8] = "60";
                    }
                    else if (gStatus == global.STATUS_KIHON_3)
                    {
                        rtnArray[mRow, 5] = "基本就業時間３";
                        rtnArray[mRow, 6] = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Sh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow2[0, "txtTm3Sm"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 7] = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Eh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow2[0, "txtTm3Em"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 8] = "60";
                    }
                    else if (gStatus == global.STATUS_KOUKYU)
                    {
                        rtnArray[mRow, 5] = "公休";
                        rtnArray[mRow, 6] = "";
                        rtnArray[mRow, 7] = "";
                        rtnArray[mRow, 8] = "";
                    }
                    else if (gStatus == global.STATUS_YUKYU)
                    {
                        rtnArray[mRow, 5] = "有給休暇";
                        rtnArray[mRow, 6] = "";
                        rtnArray[mRow, 7] = "";
                        rtnArray[mRow, 8] = "";
                    }
                    else if (gStatus == string.Empty && Utility.NulltoStr(gcMultiRow1[i, "txtSh"].Value) == string.Empty)
                    {
                        rtnArray[mRow, 5] = "";
                        rtnArray[mRow, 6] = "";
                        rtnArray[mRow, 7] = "";
                        rtnArray[mRow, 8] = "";
                    }
                    else
                    {
                        rtnArray[mRow, 5] = "";
                        rtnArray[mRow, 6] = Utility.NulltoStr(gcMultiRow1[i, "txtSh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow1[i, "txtSm"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 7] = Utility.NulltoStr(gcMultiRow1[i, "txtEh"].Value).PadLeft(2, ' ') + ":" + Utility.NulltoStr(gcMultiRow1[i, "txtEm"].Value).PadLeft(2, '0');
                        rtnArray[mRow, 8] = Utility.NulltoStr(gcMultiRow1[i, "txtRest"].Value);
                    }

                    mRow++;
                }
                
                // 合計
                rtnArray[5, 11] = Utility.NulltoStr(gcMultiRow3[0, "txtTlDays"].Value);
                rtnArray[6, 11] = Utility.NulltoStr(gcMultiRow3[0, "txtZdays"].Value);
                rtnArray[7, 11] = Utility.NulltoStr(gcMultiRow3[0, "txtKdays"].Value);
                rtnArray[8, 11] = Utility.NulltoStr(gcMultiRow3[0, "txtYuDays"].Value);
                rtnArray[9, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txtWorkTime"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txtWorkTime2"].Value).PadLeft(2, '0');
                rtnArray[10, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txtNaiZan"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txtNaiZan2"].Value).PadLeft(2, '0');
                rtnArray[11, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txtMashiZan"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txtMashiZan2"].Value).PadLeft(2, '0');
                rtnArray[12, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txt20Zan"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txt20Zan2"].Value).PadLeft(2, '0');
                rtnArray[13, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txt22Zan"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txt22Zan2"].Value).PadLeft(2, '0');
                rtnArray[14, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txtHolZan"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txtHolZan2"].Value).PadLeft(2, '0');
                rtnArray[15, 11] = string.Format("{0, 3}", Utility.NulltoStr(gcMultiRow3[0, "txtChisou"].Value)) + ":" + Utility.NulltoStr(gcMultiRow3[0, "txtChisou2"].Value).PadLeft(2, '0');

                rtnArray[37, 4] = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtKotuhi"].Value)).ToString("#,##0");
                rtnArray[38, 4] = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtSonota"].Value)).ToString("#,##0");
                rtnArray[37, 7] = txtMemo.Text.Replace("\r\n", " ");

                // 配列からシートセルに一括してデータをセットします
                rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                rng.Value2 = rtnArray;

                // 確認のためExcelのウィンドウを表示する
                //oXls.Visible = true;

                oXls.DisplayAlerts = false;

                // 印刷
                oXlsBook.PrintOut();

                // 終了メッセージ 
                MessageBox.Show("終了しました");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsMsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;
                //oxlsMsSheet = null;

                GC.Collect();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;
            }
        }

        private void gcMultiRow1_CellValidating(object sender, CellValidatingEventArgs e)
        {
            string beforeValue = string.Empty;
            string afterValue = string.Empty;

            if (chkEdit.Checked && editLogStatus)
            {
                if (gcMultiRow1.CurrentCell is TextBoxCell)
                {
                    object beforeEdit = gcMultiRow1.Rows[e.RowIndex].Cells[e.CellIndex].Value;

                    beforeValue = Utility.NulltoStr(beforeEdit);

                    if (gcMultiRow1.EditingControl != null)
                    {
                        afterValue = Utility.NulltoStr(gcMultiRow1.EditingControl.Text);
                    }

                    //if (beforeEdit != null)
                    //    MessageBox.Show(string.Format("編集前の値:{0}", beforeEdit.ToString()));
                    //else
                    //    MessageBox.Show("編集前の値:(なし)");

                    //if (gcMultiRow1.EditingControl != null)
                    //    MessageBox.Show(string.Format("編集中の値:{0}", gcMultiRow1.EditingControl.Text));
                    //else
                    //    MessageBox.Show("編集中の値:(なし)");

                    if (beforeValue != afterValue)
                    {
                        // 編集ログクラス               
                        eLogCnt++;
                        Array.Resize(ref eLog, eLogCnt);
                        eLog[eLogCnt - 1] = new clsEditLog();

                        eLog[eLogCnt - 1].sYear = Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtYear"].Value);
                        eLog[eLogCnt - 1].sMonth = Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtMonth"].Value);
                        eLog[eLogCnt - 1].sDay = Utility.NulltoStr(gcMultiRow1[e.RowIndex, "lblDay"].Value);
                        eLog[eLogCnt - 1].beforeValue = beforeValue;
                        eLog[eLogCnt - 1].afterValue = afterValue;
                        eLog[eLogCnt - 1].aID = global.loginUserID;
                        eLog[eLogCnt - 1].upDt = DateTime.Now;
                        eLog[eLogCnt - 1].stuffCode = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtSNum"].Value));
                        eLog[eLogCnt - 1].stuffName = Utility.NulltoStr(gcMultiRow2[0, "lblName"].Value);

                        switch (e.CellName)
                        {
                            case "txtStatus":
                                eLog[eLogCnt - 1].fieldName = "出勤状況";
                                break;

                            case "txtSh":
                                eLog[eLogCnt - 1].fieldName = "出勤時";
                                break;

                            case "txtSm":
                                eLog[eLogCnt - 1].fieldName = "出勤分";
                                break;

                            case "txtEh":
                                eLog[eLogCnt - 1].fieldName = "退勤時";
                                break;

                            case "txtEm":
                                eLog[eLogCnt - 1].fieldName = "退勤分";
                                break;

                            case "txtRest":
                                eLog[eLogCnt - 1].fieldName = "休憩";
                                break;

                            case "txtShopCode":
                                eLog[eLogCnt - 1].fieldName = "就労店舗コード";
                                break;

                            default:
                                break;
                        }
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("現在のセルは文字列型ではありません");
            //}

        }

        private void gcMultiRow3_CellValidating(object sender, CellValidatingEventArgs e)
        {
            string beforeValue = string.Empty;
            string afterValue = string.Empty;

            if (chkEdit.Checked && editLogStatus)
            {
                if (gcMultiRow3.CurrentCell is TextBoxCell)
                {
                    object beforeEdit = gcMultiRow3.Rows[e.RowIndex].Cells[e.CellIndex].Value;

                    beforeValue = Utility.NulltoStr(beforeEdit);

                    if (gcMultiRow3.EditingControl != null)
                    {
                        afterValue = Utility.NulltoStr(gcMultiRow3.EditingControl.Text);
                    }

                    if (beforeValue != afterValue)
                    {
                        // 編集ログクラス               
                        eLogCnt++;
                        Array.Resize(ref eLog, eLogCnt);
                        eLog[eLogCnt - 1] = new clsEditLog();

                        eLog[eLogCnt - 1].sYear = "20" + Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtYear"].Value);
                        eLog[eLogCnt - 1].sMonth = Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtMonth"].Value);
                        eLog[eLogCnt - 1].sDay = "";
                        eLog[eLogCnt - 1].beforeValue = beforeValue;
                        eLog[eLogCnt - 1].afterValue = afterValue;
                        eLog[eLogCnt - 1].aID = global.loginUserID;
                        eLog[eLogCnt - 1].upDt = DateTime.Now;
                        eLog[eLogCnt - 1].stuffCode = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtSNum"].Value));
                        eLog[eLogCnt - 1].stuffName = Utility.NulltoStr(gcMultiRow2[0, "lblName"].Value);

                        switch (e.CellName)
                        {
                            case "txtZdays":
                                eLog[eLogCnt - 1].fieldName = "実労日数";
                                break;

                            case "txtKdays":
                                eLog[eLogCnt - 1].fieldName = "公休日数";
                                break;

                            case "txtYuDays":
                                eLog[eLogCnt - 1].fieldName = "有休日数";
                                break;

                            case "txtChisou":
                                eLog[eLogCnt - 1].fieldName = "遅早回数";
                                break;

                            case "txtChisou2":
                                eLog[eLogCnt - 1].fieldName = "遅早回数";
                                break;

                            case "txtKotuhi":
                                eLog[eLogCnt - 1].fieldName = "交通費";
                                break;

                            case "txtSonota":
                                eLog[eLogCnt - 1].fieldName = "その他支給";
                                break;

                            default:
                                break;
                        }
                    }
                }
            }
        }

        // 書換え結果
        public bool reWhite { get; set; }


        ///-------------------------------------------------------------------
        /// <summary>
        ///     スタッフ配列クラスより任意のスタッフ情報を取得する </summary>
        /// <param name="sNum">
        ///     スタッフコード</param>
        /// <param name="_stf">
        ///     スタッフ配列クラス</param>
        /// <returns>
        ///     スタッフコードに該当：true、該当者なし：false</returns>
        ///-------------------------------------------------------------------
        private bool getStaffData(int sNum, out clsStaff _stf)
        {
            bool rtn = false;

            _stf = null;

            for (int i = 0; i < stf.Length; i++)
            {
                if (stf[i].スタッフコード == sNum)
                {
                    _stf = stf[i];
                    rtn = true;
                    break;
                }
            }

            return rtn;
        }
    }
}
