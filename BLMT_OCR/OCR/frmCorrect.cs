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
    public partial class frmCorrect : Form
    {
        /// ------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ </summary>
        /// <param name="dbName">
        ///     会社領域データベース名</param>
        /// <param name="comName">
        ///     会社名</param>
        /// <param name="sID">
        ///     処理モード</param>
        /// <param name="eMode">
        ///     true:通常処理, false:応援移動票画面からの呼出し
        ///     （エラーチェック、汎用データ作成機能なし）</param>
        /// ------------------------------------------------------------
        public frmCorrect(string sID)
        {
            InitializeComponent();

            //_dbName = dbName;       // データベース名
            //_comName = comName;     // 会社名
            dID = sID;              // 処理モード
            //_eMode = eMode;         // 処理モード2

            /* テーブルアダプターマネージャーに勤務票ヘッダ、明細テーブル、
             * 過去勤務票ヘッダ、過去明細テーブルアダプターを割り付ける */
            adpMn.勤務票ヘッダTableAdapter = hAdp;
            adpMn.勤務票明細TableAdapter = iAdp;

            pAdpMn.過去勤務票ヘッダTableAdapter = phAdp;
            pAdpMn.過去勤務票明細TableAdapter = piAdp;

            dAdpMn.保留勤務票ヘッダTableAdapter = dhAdp;
            dAdpMn.保留勤務票明細TableAdapter = diAdp;

            // 休日テーブル読み込み
            kAdp.Fill(dts.休日);

            // 環境設定読み込み
            //cAdp.Fill(dts.環境設定);
        }

        // データアダプターオブジェクト
        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();

        DataSet1TableAdapters.TableAdapterManager pAdpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter piAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();
        //DataSet1TableAdapters.環境設定TableAdapter cAdp = new DataSet1TableAdapters.環境設定TableAdapter();

        DataSet1TableAdapters.TableAdapterManager dAdpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.保留勤務票ヘッダTableAdapter dhAdp = new DataSet1TableAdapters.保留勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.保留勤務票明細TableAdapter diAdp = new DataSet1TableAdapters.保留勤務票明細TableAdapter();

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

        string[] idArray = null;            // 印刷用ヘッダＩＤ配列 2019/01/15
        string b_img = string.Empty;        // 印刷用画像名 2019/01/15

        bool editLogStatus = true;

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
            //var s = dts.環境設定.Single(a => a.ID == global.configKEY);
            Utility u = new Utility();
            u.getXlsToArray(global.cnfMsPath, xlsArray, ref stf);

            // 店舗マスター配列クラス作成
            u.getShopMaster(stf, ref shp);

            // 勤務データ登録
            if (dID == string.Empty)
            {
                // CSVデータをMDBへ読み込みます
                GetCsvDataToMDB();

                // データセットへデータを読み込みます
                getDataSet();   // 出勤簿

                // データテーブル件数カウント
                if (dts.勤務票ヘッダ.Count == 0)
                {
                    MessageBox.Show("出勤簿がありません", "出勤簿登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //終了処理
                    Environment.Exit(0);
                }

                // キー配列作成
                keyArrayCreate();
            }

            // キャプション
            this.Text = "出勤簿";

            // GCMultiRow初期化
            gcMrSetting();

            // 編集作業、過去データ表示の判断
            if (dID == string.Empty) // パラメータのヘッダIDがないときは編集作業
            {
                // 最初のレコードを表示
                cI = 0;
                showOcrData(cI);
            }

            // tagを初期化
            this.Tag = string.Empty;

            // 現在の表示倍率を初期化
            gl.miMdlZoomRate = 0f;
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

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            if (cI + 1 < dts.勤務票ヘッダ.Rows.Count)
            {
                cI++;
                showOcrData(cI);
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
            string errMsg = "出勤簿テーブル更新";

            try
            {
                // 勤務票ヘッダテーブル行を取得
                DataSet1.勤務票ヘッダRow r = dts.勤務票ヘッダ.Single(a => a.ID == iX);

                // 勤務票ヘッダテーブルセット更新
                r.年 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtYear"].Value));
                r.月 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtMonth"].Value));
                r.スタッフコード = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtSNum"].Value));
                r.氏名 = Utility.NulltoStr(gcMultiRow2[0, "lblName"].Value);

                if (sStf != null)
                {
                    r.担当エリアマネージャー名 = sStf.エリアマネージャー名;
                    r.エリアコード = sStf.エリアコード;
                    r.エリア名 = sStf.エリア名;
                    r.給与形態 = sStf.給与区分;
                }
                else
                {
                    r.担当エリアマネージャー名 = string.Empty;
                    r.エリアコード = 0;
                    r.エリア名 = string.Empty;
                    r.給与形態 = 0;
                }

                r.店舗コード = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "lblShopCode"].Value));
                r.店舗名 = Utility.NulltoStr(gcMultiRow2[0, "lblShopName"].Value);

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
                r.特休日数 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtTokuDays"].Value));  // 特休日数：2020/05/12

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

                if (checkBox1.Checked)
                {
                    r.確認 = global.flgOn;
                }
                else
                {
                    r.確認 = global.flgOff;
                }

                r.備考 = txtMemo.Text;
                r.基本実労働時 = Utility.NulltoStr(gcMultiRow2[0, "txtWh"].Value);
                r.基本実労働分 = Utility.NulltoStr(gcMultiRow2[0, "txtWm"].Value);

                // 勤務票明細テーブルセット更新
                for (int i = 0; i < gcMultiRow1.RowCount; i++)
                {
                    int sID = int.Parse(gcMultiRow1[i, "txtID"].Value.ToString());
                    DataSet1.勤務票明細Row m = (DataSet1.勤務票明細Row)dts.勤務票明細.FindByID(sID);

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

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            cI = dts.勤務票ヘッダ.Rows.Count - 1;
            showOcrData(cI);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            if (cI > 0)
            {
                cI--;
                showOcrData(cI);
            }
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            cI = 0;
            showOcrData(cI);
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
            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            cI = hScrollBar1.Value;
            showOcrData(cI);
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
        }

        ///-------------------------------------------------------------------------------
        /// <summary>
        ///     １．指定した勤務票ヘッダデータと勤務票明細データを削除する　
        ///     ２．該当する画像データを削除する</summary>
        /// <param name="i">
        ///     勤務票ヘッダRow インデックス</param>
        ///-------------------------------------------------------------------------------
        private void DataDelete(int i)
        {
            string sImgNm = string.Empty;
            string errMsg = string.Empty;

            // 勤務票データ削除
            try
            {
                // ヘッダIDを取得します
                DataSet1.勤務票ヘッダRow r = dts.勤務票ヘッダ.Single(a => a.ID == cID[i]);

                // 画像ファイル名を取得します
                sImgNm = r.画像名;

                // データテーブルからヘッダIDが一致する勤務票明細データを削除します。
                errMsg = "勤務票明細データ";
                foreach (DataSet1.勤務票明細Row item in dts.勤務票明細.Rows)
                {
                    if (item.RowState != DataRowState.Deleted && item.ヘッダID == r.ID)
                    {
                        item.Delete();
                    }
                }

                // データテーブルから勤務票ヘッダデータを削除します
                errMsg = "勤務票ヘッダデータ";
                r.Delete();

                // データベース更新
                adpMn.UpdateAll(dts);

                // 画像ファイルを削除します
                errMsg = "勤務管理表画像";
                if (sImgNm != string.Empty)
                {
                    if (System.IO.File.Exists(Properties.Settings.Default.dataPath + sImgNm))
                    {
                        System.IO.File.Delete(Properties.Settings.Default.dataPath + sImgNm);
                    }
                }

                // 配列キー再構築
                keyArrayCreate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(errMsg + "の削除に失敗しました" + Environment.NewLine + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
            }

        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            //「受入データ作成終了」「勤務票データなし」以外での終了のとき
            if (this.Tag.ToString() != END_MAKEDATA && this.Tag.ToString() != END_NODATA)
            {
                //if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                //{
                //    e.Cancel = true;
                //    return;
                //}

                // カレントデータ更新
                if (dID == string.Empty)
                {
                    CurDataUpDate(cID[cI]);
                }

                //// 勤務表データのない帰宅後勤務データを削除する
                //kitakuClean();
            }

            // データベース更新
            adpMn.UpdateAll(dts);

            // 解放する
            this.Dispose();
        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
        }

        /// -----------------------------------------------------------------------
        /// <summary>
        ///     PCA給与X受入CSVデータ出力 </summary>
        /// -----------------------------------------------------------------------
        private void textDataMake()
        {
            if (MessageBox.Show("PCA給与X用データを作成します。" + Environment.NewLine + Environment.NewLine +
              "対象となる出勤簿データと画像の印刷も行います。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName);

            // エラーチェックを実行する
            if (getErrData(cI, ocr)) // エラーがなかったとき
            {
                // OCROutputクラス インスタンス生成
                OCROutput kd = new OCROutput(this, dts);

                // 汎用データ作成
                kd.SaveData();

                // 画像ファイル退避
                tifFileMove();

                // 過去出勤簿データ作成
                saveLastData();

                // 設定月数分経過した過去の画像と出勤簿データを削除する
                deleteArchived();

                // 勤務票データ＆画像印刷：2019/01/15
                batch_ImagePrint(); // 画像
                batch_DataPrint();  // データ

                //if (MessageBox.Show("出勤簿データと画像の印刷を行います。よろしいですか", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Yes)
                //{
                //    // 画像印刷
                //    batch_ImagePrint();

                //    // データ印刷
                //    batch_DataPrint();
                //}

                // 勤務票データ削除
                deleteDataAll();

                // MDBファイル最適化
                mdbCompact();

                //終了
                MessageBox.Show("終了しました。ＰＣＡ給与Ｘで給与データ受け入れを行ってください。", "ＰＣＡ給与Ｘ給与データ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // エラーあり
                showOcrData(cI);    // データ表示
                ErrShow(ocr);   // エラー表示
            }
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
        private bool getErrData(int cIdx, OCRData ocr)
        {
            // カレントレコード更新
            CurDataUpDate(cID[cIdx]);

            // エラー番号初期化
            ocr._errNumber = ocr.eNothing;

            // エラーメッセージクリーン
            ocr._errMsg = string.Empty;

            // エラーチェック実行①:カレントレコードから最終レコードまで
            if (!ocr.errCheckMain(cIdx, (dts.勤務票ヘッダ.Rows.Count - 1), this, dts, cID, sStf, stf))
            {
                return false;
            }

            // エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (cIdx > 0)
            {
                if (!ocr.errCheckMain(0, (cIdx - 1), this, dts, cID, sStf, stf))
                {
                    return false;
                }
            }

            // エラーなし
            lblErrMsg.Text = string.Empty;

            return true;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     画像ファイル退避処理 勤務票・応援移動票</summary>
        ///----------------------------------------------------------------------------------
        private void tifFileMove()
        {
            // 移動先フォルダがあるか？なければ作成する（TIFフォルダ）
            if (!System.IO.Directory.Exists(Properties.Settings.Default.tifPath))
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.tifPath);
            
            string fromImg = string.Empty;
            string toImg = string.Empty;

            // 出勤簿ヘッダデータを取得する
            foreach (var t in dts.勤務票ヘッダ.OrderBy(a => a.ID))
            {              
                // 出勤簿画像ファイルパスを取得する
                fromImg = Properties.Settings.Default.dataPath + t.画像名;

                // 出勤簿移動先ファイルパス
                toImg = Properties.Settings.Default.tifPath + t.画像名;

                // 同名ファイルが既に登録済みのときは削除する
                if (System.IO.File.Exists(toImg)) System.IO.File.Delete(toImg);

                // ファイルを移動する
                if (System.IO.File.Exists(fromImg)) System.IO.File.Move(fromImg, toImg);
            }
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

        /// ---------------------------------------------------------------------------------
        /// <summary>
        ///     設定月数分経過した過去画像と過去勤務データ、過去応援移動票データを削除する </summary> 
        /// ---------------------------------------------------------------------------------
        private void deleteArchived()
        {
            // 削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (global.cnfArchived == global.flgOff) return;

            try
            {
                // 削除年月の取得
                DateTime dt = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");
                DateTime delDate = dt.AddMonths(global.cnfArchived * (-1));

                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月

                //int _waYYMM = (delDate.Year - Properties.Settings.Default.rekiHosei) * 100 + _dMM;   //基準年月(和暦）

                // 設定月数分経過した過去画像・過去勤務票データを削除する
                deleteLastDataArchived(_dYYMM);
            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像・過去勤務票データ削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
            finally
            {
                //if (ocr.sCom.Connection.State == ConnectionState.Open) ocr.sCom.Connection.Close();
            }
        }

        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票データ削除～登録 </summary>
        /// ---------------------------------------------------------------------------
        private void saveLastData()
        {
            try
            {
                // データベース更新
                adpMn.UpdateAll(dts);
                pAdpMn.UpdateAll(dts);

                //  過去勤務票ヘッダデータとその明細データを削除します
                //deleteLastData();
                delPastData();

                // データセットへデータを再読み込みします
                getDataSet();

                // 過去勤務票ヘッダデータと過去勤務票明細データを作成します
                addLastdata();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "過去勤務票データ作成エラー", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }


        ///------------------------------------------------------
        /// <summary>
        ///     過去勤務票データ削除 </summary>
        ///------------------------------------------------------
        private void delPastData()
        {
            // 過去勤務票ヘッダデータ削除
            foreach (var t in dts.勤務票ヘッダ)
            {
                string sBusho = t.スタッフコード.ToString();
                int sYY = t.年;
                int sMM = t.月;

                // 過去勤務票ヘッダ削除
                delPastHeader(sBusho, sYY, sMM);
            }

            // 過去勤務票明細データ削除
            delPastItem();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータ削除 </summary>
        /// <param name="bCode">
        ///     スタッフコード</param>
        /// <param name="syy">
        ///     対象年</param>
        /// <param name="smm">
        ///     対象月</param>
        ///----------------------------------------------------------------
        private void delPastHeader(string bCode, int syy, int smm)
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("delete from 過去勤務票ヘッダ ");
                sb.Append("where スタッフコード = ? and 年 = ? and 月 = ?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@b", bCode);
                sCom.Parameters.AddWithValue("@y", syy);
                sCom.Parameters.AddWithValue("@m", smm);

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     過去勤務票明細データ削除 </summary>
        ///--------------------------------------------------------
        private void delPastItem()
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("delete a.ヘッダID from  過去勤務票明細 as a ");
                sb.Append("where not EXISTS (select * from 過去勤務票ヘッダ ");
                sb.Append("WHERE 過去勤務票ヘッダ.ID = a.ヘッダID)");
                
                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータとその明細データを削除します</summary>    
        ///     
        /// -------------------------------------------------------------------------
        private void deleteLastData()
        {
            OleDbCommand sCom = new OleDbCommand();
            OleDbCommand sCom2 = new OleDbCommand();
            OleDbCommand sCom3 = new OleDbCommand();

            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);
            mdb.dbConnect(sCom2);
            mdb.dbConnect(sCom3);

            OleDbDataReader dR = null;
            OleDbDataReader dR2 = null;

            StringBuilder sb = new StringBuilder();
            StringBuilder sbd = new StringBuilder();

            try
            {
                // 対象データ : 取消は対象外とする
                sb.Clear();
                sb.Append("Select 勤務票明細.ヘッダID, 勤務票明細.ID,");
                sb.Append("勤務票ヘッダ.年, 勤務票ヘッダ.月, 勤務票ヘッダ.日,");
                sb.Append("勤務票明細.社員番号 from 勤務票ヘッダ inner join 勤務票明細 ");
                sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID ");
                sb.Append("where 勤務票明細.取消 = '").Append(global.FLGOFF).Append("'");
                sb.Append("order by 勤務票明細.ヘッダID, 勤務票明細.ID");

                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    // ヘッダID
                    string hdID = string.Empty;

                    // 日付と社員番号で過去データを抽出（該当するのは1件）
                    sb.Clear();
                    sb.Append("Select 過去勤務票明細.ヘッダID,過去勤務票明細.ID,");
                    sb.Append("過去勤務票ヘッダ.年, 過去勤務票ヘッダ.月, 過去勤務票ヘッダ.日,");
                    sb.Append("過去勤務票明細.社員番号 from 過去勤務票ヘッダ inner join 過去勤務票明細 ");
                    sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
                    sb.Append("where ");
                    sb.Append("過去勤務票ヘッダ.年 = ? and ");
                    sb.Append("過去勤務票ヘッダ.月 = ? and ");
                    sb.Append("過去勤務票ヘッダ.日 = ? and ");
                    sb.Append("過去勤務票ヘッダ.データ領域名 = ? and ");
                    sb.Append("過去勤務票明細.社員番号 = ?");

                    sCom2.CommandText = sb.ToString();
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@yy", dR["年"].ToString());
                    sCom2.Parameters.AddWithValue("@mm", dR["月"].ToString());
                    sCom2.Parameters.AddWithValue("@dd", dR["日"].ToString());
                    sCom2.Parameters.AddWithValue("@db", _dbName);
                    sCom2.Parameters.AddWithValue("@n", dR["社員番号"].ToString());

                    dR2 = sCom2.ExecuteReader();

                    while (dR2.Read())
                    {
                        //// ヘッダIDを取得
                        //if (hdID == string.Empty)
                        //{
                        //    hdID = dR2["ヘッダID"].ToString();
                        //}

                        // 過去勤務票明細レコード削除
                        sbd.Clear();
                        sbd.Append("delete from 過去勤務票明細 ");
                        sbd.Append("where ID = ?");

                        sCom3.CommandText = sbd.ToString();
                        sCom3.Parameters.Clear();
                        sCom3.Parameters.AddWithValue("@id", dR2["ID"].ToString());

                        sCom3.ExecuteNonQuery();
                    }

                    dR2.Close();
                }

                dR.Close();

                // データベース接続解除
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }

                if (sCom2.Connection.State == ConnectionState.Open)
                {
                    sCom2.Connection.Close();
                }

                if (sCom3.Connection.State == ConnectionState.Open)
                {
                    sCom3.Connection.Close();
                }

                // データベース再接続
                mdb.dbConnect(sCom);
                mdb.dbConnect(sCom2);

                // 明細データのない過去勤務票ヘッダデータを抽出
                sb.Clear();
                sb.Append("Select 過去勤務票ヘッダ.ID,過去勤務票明細.ヘッダID ");
                sb.Append("from 過去勤務票ヘッダ left join 過去勤務票明細 ");
                sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
                sb.Append("where ");
                sb.Append("過去勤務票明細.ヘッダID is null");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    // 過去勤務票ヘッダレコード削除
                    sbd.Clear();

                    sbd.Append("delete from 過去勤務票ヘッダ ");
                    sbd.Append("where ID = ?");

                    sCom2.CommandText = sbd.ToString();
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@id", dR["ID"].ToString());

                    sCom2.ExecuteNonQuery();
                }

                dR.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }

                if (sCom2.Connection.State == ConnectionState.Open)
                {
                    sCom2.Connection.Close();
                }

                if (sCom3.Connection.State == ConnectionState.Open)
                {
                    sCom3.Connection.Close();
                }
            }
        }


        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータと過去勤務票明細データを作成します</summary>
        ///     
        /// -------------------------------------------------------------------------
        private void addLastdata()
        {
            for (int i = 0; i < dts.勤務票ヘッダ.Rows.Count; i++)
            {
                // -------------------------------------------------------------------------
                //      過去勤務票ヘッダレコードを作成します
                // -------------------------------------------------------------------------
                DataSet1.勤務票ヘッダRow hr = (DataSet1.勤務票ヘッダRow)dts.勤務票ヘッダ.Rows[i];
                DataSet1.過去勤務票ヘッダRow nr = dts.過去勤務票ヘッダ.New過去勤務票ヘッダRow();

                #region テーブルカラム名比較～データコピー

                // 勤務票ヘッダのカラムを順番に読む
                for (int j = 0; j < dts.勤務票ヘッダ.Columns.Count; j++)
                {
                    // 過去勤務票ヘッダのカラムを順番に読む
                    for (int k = 0; k < dts.過去勤務票ヘッダ.Columns.Count; k++)
                    {
                        // フィールド名が同じであること
                        if (dts.勤務票ヘッダ.Columns[j].ColumnName == dts.過去勤務票ヘッダ.Columns[k].ColumnName)
                        {
                            if (dts.過去勤務票ヘッダ.Columns[k].ColumnName == "更新年月日")
                            {
                                nr[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                            }
                            else
                            {
                                nr[k] = hr[j];          // データをコピー
                            }
                            break;
                        }
                    }
                }
                #endregion

                // 過去勤務票ヘッダデータテーブルに追加
                dts.過去勤務票ヘッダ.Add過去勤務票ヘッダRow(nr);

                // -------------------------------------------------------------------------
                //      過去勤務票明細レコードを作成します
                // -------------------------------------------------------------------------
                var mm = dts.勤務票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                           a.ヘッダID == hr.ID)
                    .OrderBy(a => a.ID);

                foreach (var item in mm)
                {
                    DataSet1.勤務票明細Row m = (DataSet1.勤務票明細Row)dts.勤務票明細.Rows.Find(item.ID);
                    DataSet1.過去勤務票明細Row nm = dts.過去勤務票明細.New過去勤務票明細Row();

                    //// 社員番号が空白のレコードは対象外とします
                    //if (m.社員番号 == string.Empty) continue;

                    #region  テーブルカラム名比較～データコピー

                    // 勤務票明細のカラムを順番に読む
                    for (int j = 0; j < dts.勤務票明細.Columns.Count; j++)
                    {
                        // IDはオートナンバーのため値はコピーしない
                        if (dts.勤務票明細.Columns[j].ColumnName != "ID")
                        {
                            // 過去勤務票ヘッダのカラムを順番に読む
                            for (int k = 0; k < dts.過去勤務票明細.Columns.Count; k++)
                            {
                                // フィールド名が同じであること
                                if (dts.勤務票明細.Columns[j].ColumnName == dts.過去勤務票明細.Columns[k].ColumnName)
                                {
                                    if (dts.過去勤務票明細.Columns[k].ColumnName == "更新年月日")
                                    {
                                        nm[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                                    }
                                    else
                                    {
                                        nm[k] = m[j];          // データをコピー
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    #endregion

                    // 過去勤務票明細データテーブルに追加
                    dts.過去勤務票明細.Add過去勤務票明細Row(nm);
                }
            }

            // データベース更新
            pAdpMn.UpdateAll(dts);
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     出勤簿画像印刷 : 2019/01/15 </summary>
        ///-----------------------------------------------------------
        private void batch_ImagePrint()
        {
            //foreach (var t in dts.勤務票ヘッダ.OrderBy(a => a.ID)) // 2019/05/07 コメント化
            foreach (var t in dts.勤務票ヘッダ.OrderBy(a => a.スタッフコード)) // スタッフコードでソート 2019/05/07
            {
                b_img = Properties.Settings.Default.tifPath + t.画像名;
                printDocument1.Print();
            }

            // 後片付け
            printDocument1.Dispose();
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     出勤簿データ印刷 : 2019/01/15 </summary>
        ///-------------------------------------------------------
        private void batch_DataPrint()
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
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート

            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            object[,] rtnArray = null;

            int pCnt = 1;   // ページカウント

            try
            {
                //foreach (var t in dts.勤務票ヘッダ) // 2019/05/07 コメント化
                foreach (var t in dts.勤務票ヘッダ.OrderBy(a => a.スタッフコード)) // スタッフコード順で出力 2019/05/07
                {
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートのセルを一括して配列に取得します
                    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                    rtnArray = (object[,])rng.Value2;

                    rtnArray[2, 2] = "20" + t.年 + "年" + t.月 + "月度";
                    rtnArray[2, 5] = t.スタッフコード.ToString("D5");
                    rtnArray[2, 6] = t.氏名;
                    rtnArray[2, 8] = t.店舗コード.ToString("D5");
                    rtnArray[2, 9] = t.店舗名;

                    DateTime dt = DateTime.Today;

                    int mRow = 5;

                    var s = t.Get勤務票明細Rows();

                    foreach (var m in s.OrderBy(a => a.ID))
                    {                      
                        if (!DateTime.TryParse(m.暦年 + "/" + m.暦月 + "/" + m.日, out dt))
                        {
                            continue;
                        }

                        rtnArray[mRow, 2] = dt.ToShortDateString();
                        rtnArray[mRow, 3] = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                        string gStatus = m.出勤状況;
                        rtnArray[mRow, 4] = gStatus;

                        if (gStatus == global.STATUS_KIHON_1)
                        {
                            rtnArray[mRow, 5] = "基本就業時間１";
                            rtnArray[mRow, 6] = t.基本就業時間帯1開始時.PadLeft(2, ' ') + ":" + t.基本就業時間帯1開始分.PadLeft(2, '0');
                            rtnArray[mRow, 7] = t.基本就業時間帯1終了時.PadLeft(2, ' ') + ":" + t.基本就業時間帯1終了分.PadLeft(2, '0');
                            rtnArray[mRow, 8] = "60";
                        }
                        else if (gStatus == global.STATUS_KIHON_2)
                        {
                            rtnArray[mRow, 5] = "基本就業時間２";
                            rtnArray[mRow, 6] = t.基本就業時間帯2開始時.PadLeft(2, ' ') + ":" + t.基本就業時間帯2開始分.PadLeft(2, '0');
                            rtnArray[mRow, 7] = t.基本就業時間帯2終了時.PadLeft(2, ' ') + ":" + t.基本就業時間帯2終了分.PadLeft(2, '0');
                            rtnArray[mRow, 8] = "60";
                        }
                        else if (gStatus == global.STATUS_KIHON_3)
                        {
                            rtnArray[mRow, 5] = "基本就業時間３";
                            rtnArray[mRow, 6] = t.基本就業時間帯3開始時.PadLeft(2, ' ') + ":" + t.基本就業時間帯3開始分.PadLeft(2, '0');
                            rtnArray[mRow, 7] = t.基本就業時間帯3終了時.PadLeft(2, ' ') + ":" + t.基本就業時間帯3終了分.PadLeft(2, '0');
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
                        else if (gStatus == string.Empty && m.出勤時 == string.Empty)
                        {
                            rtnArray[mRow, 5] = "";
                            rtnArray[mRow, 6] = "";
                            rtnArray[mRow, 7] = "";
                            rtnArray[mRow, 8] = "";
                        }
                        else
                        {
                            rtnArray[mRow, 5] = "";
                            rtnArray[mRow, 6] = m.出勤時.PadLeft(2, ' ') + ":" + m.出勤分.PadLeft(2, '0');
                            rtnArray[mRow, 7] = m.退勤時.PadLeft(2, ' ') + ":" + m.退勤分.PadLeft(2, '0');
                            rtnArray[mRow, 8] = m.休憩;
                        }

                        mRow++;
                    }

                    // 合計

                    // 2020/05/18
                    if (t.Is特休日数Null())
                    {
                        rtnArray[5, 11] = t.実労日数 + t.有休日数;
                        rtnArray[16, 11] = global.FLGOFF;
                    }
                    else
                    {
                        rtnArray[5, 11] = t.実労日数 + t.有休日数 + t.特休日数;
                        rtnArray[16, 11] = t.特休日数;
                    }

                    //rtnArray[5, 11] = t.実労日数 + t.有休日数 + t.特休日数;

                    rtnArray[6, 11] = t.実労日数;
                    rtnArray[7, 11] = t.公休日数;
                    rtnArray[8, 11] = t.有休日数;
                    
                    double kihon = global.cnfKihonWh * 60 + global.cnfKihonWm;

                    double _w20Time = 0;
                    double _w22Time = 0;
                    double _naiZan = 0;
                    double _mashiZan = 0;
                    double _doniShuku = 0;

                    double wt = getTotalWorkTime_n(out _w20Time, out _w22Time, out _naiZan, out _mashiZan, out _doniShuku, kihon, t, s);

                    rtnArray[9, 11] = string.Format("{0, 3}", (int)(wt / 60) + ":" + ((int)(wt % 60)).ToString("D2"));
                    rtnArray[10, 11] = string.Format("{0, 3}", (int)(_naiZan / 60) + ":" + ((int)(_naiZan % 60)).ToString("D2"));
                    rtnArray[11, 11] = string.Format("{0, 3}", (int)(_mashiZan / 60) + ":" + ((int)(_mashiZan % 60)).ToString("D2"));
                    rtnArray[12, 11] = string.Format("{0, 3}", (int)(_w20Time / 60) + ":" + ((int)(_w20Time % 60)).ToString("D2"));
                    rtnArray[13, 11] = string.Format("{0, 3}", (int)(_w22Time / 60) + ":" + ((int)(_w22Time % 60)).ToString("D2"));
                    rtnArray[14, 11] = string.Format("{0, 3}", (int)(_doniShuku / 60) + ":" + ((int)(_doniShuku % 60)).ToString("D2"));
                    rtnArray[15, 11] = string.Format("{0, 3}", t.遅早時間時 + ":" + t.遅早時間分.ToString("D2"));
                    
                    rtnArray[37, 4] = t.交通費.ToString("#,##0");
                    rtnArray[38, 4] = t.その他支給.ToString("#,##0");
                    //rtnArray[37, 7] = txtMemo.Text.Replace("\r\n", " "); 
                    rtnArray[37, 7] = t.備考; // 2019/02/18 改修

                    // 配列からシートセルに一括してデータをセットします
                    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;
                }
                
                // 確認のためExcelのウィンドウを表示する
                //oXls.Visible = true;

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                // 印刷
                oXlsBook.PrintOut();

                //// 終了メッセージ 
                //MessageBox.Show("終了しました");
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


        ///----------------------------------------------------------
        /// <summary>
        ///     総労働時間取得 </summary>
        /// <returns>
        ///     総労働時間・分</returns>
        ///----------------------------------------------------------
        private double getTotalWorkTime_n(out double w20Total, out double w22Total, out double naiZan, out double mashiZan, out double doniShuku, double kihon, DataSet1.勤務票ヘッダRow t, DataSet1.勤務票明細Row [] _m)
        {
            DateTime sDt;
            DateTime eDt;
            DateTime tDt;
            DateTime gDt;

            string sStatus = "";
            string sSh = "";
            string sSm = "";
            string sEh = "";
            string sEm = "";
            string sRest = "";
            const string REST60 = "60";
            //const double kihonWork = 480; 

            double wTotal = 0;
            w20Total = 0;
            w22Total = 0;
            naiZan = 0;
            mashiZan = 0;
            doniShuku = 0;

            int stfX = 0;

            for (int i = 0; i < stf.Length; i++)
            {
                if (stf[i].スタッフコード == t.スタッフコード)
                {
                    stfX = i;
                    break;
                }
            }

            // 個人基本労働時間
            double pKihon = (Utility.StrtoInt(t.基本実労働時)) * 60 + Utility.StrtoInt(t.基本実労働分);

            // 勤務票明細
            foreach (var m in _m)
            {
                if (!DateTime.TryParse(m.暦年 + "/" + m.暦月 + "/" + m.日, out gDt))
                {
                    continue;
                }

                sStatus = m.出勤状況;

                // 有給休暇取得日：2017/12/05
                if (sStatus == global.STATUS_YUKYU)
                {
                    // 時給者のとき基本労働時間を労働時間に加算
                    if (stf[stfX].給与区分.ToString() == global.JIKKYU)
                    {
                        wTotal += pKihon;
                    }

                    continue;
                }

                // 基本就業時間１～３か？
                if (sStatus == global.STATUS_KIHON_1 || sStatus == global.STATUS_KIHON_2 || sStatus == global.STATUS_KIHON_3)
                {
                    if (sStatus == global.STATUS_KIHON_1)
                    {
                        // 基本就業時間1
                        sSh = t.基本就業時間帯1開始時;
                        sSm = t.基本就業時間帯1開始分;
                        sEh = t.基本就業時間帯1終了時;
                        sEm = t.基本就業時間帯1終了分;
                        sRest = REST60;
                    }
                    else if (sStatus == global.STATUS_KIHON_2)
                    {
                        // 基本就業時間2
                        sSh = t.基本就業時間帯2開始時;
                        sSm = t.基本就業時間帯2開始分;
                        sEh = t.基本就業時間帯2終了時;
                        sEm = t.基本就業時間帯2終了分;
                        sRest = REST60;
                    }
                    else if (sStatus == global.STATUS_KIHON_3)
                    {
                        // 基本就業時間3
                        sSh = t.基本就業時間帯3開始時;
                        sSm = t.基本就業時間帯3開始分;
                        sEh = t.基本就業時間帯3終了時;
                        sEm = t.基本就業時間帯3終了分;
                        sRest = REST60;
                    }
                }
                else
                {
                    sSh = m.出勤時;
                    sSm = m.出勤分;
                    sEh = m.退勤時;
                    sEm = m.退勤分;
                    sRest = m.休憩;
                }

                // 開始時刻
                bool st = DateTime.TryParse(sSh + ":" + sSm + ":00", out sDt);

                if (!st)
                {
                    continue;
                }

                // 終了時刻
                bool et;
                bool over23 = false;
                int eh = Utility.StrtoInt(sEh);
                if (eh > 23)
                {
                    eh -= 24;
                    over23 = true;
                }

                et = DateTime.TryParse(eh + ":" + sEm + ":00", out eDt);

                if (et)
                {
                    // 翌日
                    if (over23)
                    {
                        eDt = eDt.AddDays(1);
                    }
                }
                else
                {
                    continue;
                }

                // 開始時刻と終了時刻が正しく登録されているとき
                double wTime = 0;
                if (st && et)
                {
                    // 労働時間
                    wTime = Utility.GetTimeSpan(sDt, eDt).TotalMinutes; // 開始～終了
                    wTime -= Utility.StrtoInt(sRest);  // 休憩時間を差し引く
                    wTotal += wTime;

                    // 個人基本実労働時間がゼロでないとき
                    if (pKihon > 0)
                    {
                        // 基本時間内残業時間：8時間以上のとき
                        if (wTime > pKihon && wTime > kihon)
                        {
                            // 2017/12/07
                            if (kihon > pKihon)
                            {
                                naiZan += (kihon - pKihon);
                            }
                        }

                        // 基本時間内残業時間：個人基本超え、8時間以内
                        if (wTime > pKihon && wTime <= kihon)
                        {
                            naiZan += (wTime - pKihon);
                        }
                    }

                    // 割増残業時間
                    if (wTime > kihon)
                    {
                        mashiZan += (wTime - kihon);
                    }

                    // 20時以降勤務時間：直営店の日給者＆直営店の15日締時給者のみ対象 2017/12/05
                    if (stf[stfX].店舗区分.ToString() == global.TENPO_CHOKUEI)
                    {
                        if (stf[stfX].給与区分.ToString() == global.NIKKYU ||
                            (stf[stfX].給与区分.ToString() == global.JIKKYU && stf[stfX].締日区分.ToString() == global.SHIME_15))
                        {
                            if (eDt > global.dt2000)
                            {
                                w20Total += Utility.GetTimeSpan(global.dt2000, eDt).TotalMinutes;
                            }
                        }
                    }

                    // 22時以降勤務時間
                    if (eDt > global.dt2200)
                    {
                        w22Total += Utility.GetTimeSpan(global.dt2200, eDt).TotalMinutes;
                    }

                    // 最初の日付
                    int sd = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "lblDay"].Value));
                    int ed = 31;

                    if (sd > 1)
                    {
                        ed = sd - 1;
                    }

                    // 土日・祝日の勤務時間：直営店の日給者＆直営店の15日締時給者のみ対象 2017/12/05
                    if (stf[stfX].店舗区分.ToString() == global.TENPO_CHOKUEI)
                    {
                        string wd = ("日月火水木金土").Substring(int.Parse(gDt.DayOfWeek.ToString("d")), 1);
             
                        if (stf[stfX].給与区分.ToString() == global.NIKKYU ||
                            (stf[stfX].給与区分.ToString() == global.JIKKYU && stf[stfX].締日区分.ToString() == global.SHIME_15))
                        {
                            if (wd == "土" || wd == "日")
                            {
                                // 土日
                                doniShuku += wTime;
                            }
                            else
                            {
                                // 土日以外の休日
                                if (dts.休日.Any(a => a.年月日 == gDt))
                                {
                                    doniShuku += wTime;
                                }
                            }
                        }
                    }
                }
            }

            return wTotal;
        }


        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
        //    //if (e.RowIndex < 0) return;

        //    string colName = dGV.Columns[e.ColumnIndex].Name;

        //    if (colName == cSH || colName == cSE || colName == cEH || colName == cEE ||
        //        colName == cZH || colName == cZE || colName == cSIH || colName == cSIE)
        //    {
        //        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
        //    }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //string colName = dGV.Columns[dGV.CurrentCell.ColumnIndex].Name;
            ////if (colName == cKyuka || colName == cCheck)
            ////{
            ////    if (dGV.IsCurrentCellDirty)
            ////    {
            ////        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            ////        dGV.RefreshEdit();
            ////    }
            ////}
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellEnter_1(object sender, DataGridViewCellEventArgs e)
        {
            //// 時が入力済みで分が未入力のとき分に"00"を表示します
            //if (dGV[ColH, dGV.CurrentRow.Index].Value != null)
            //{
            //    if (dGV[ColH, dGV.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
            //    {
            //        if (dGV[ColM, dGV.CurrentRow.Index].Value == null)
            //        {
            //            dGV[ColM, dGV.CurrentRow.Index].Value = "00";
            //        }
            //        else if (dGV[ColM, dGV.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
            //        {
            //            dGV[ColM, dGV.CurrentRow.Index].Value = "00";
            //        }
            //    }
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

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     基準年月以前の過去勤務票ヘッダデータとその明細データを削除します</summary>
        /// <param name="sYYMM">
        ///     基準年月</param>     
        /// -------------------------------------------------------------------------
        private void deleteLastDataArchived(int sYYMM)
        {
            // データ読み込み
            getDataSet();

            // 基準年月以前の過去勤務票ヘッダデータを取得します
            var h = dts.過去勤務票ヘッダ
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                ((2000 + a.年) * 100 + a.月) < sYYMM);

            // foreach用の配列を作成
            var hLst = h.ToList();

            foreach (var lh in hLst)
            {
                // ヘッダIDが一致する過去勤務票明細を取得します
                var m = dts.過去勤務票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.ヘッダID == lh.ID);

                // foreach用の配列を作成
                var list = m.ToList();

                // 該当過去勤務票明細を削除します
                foreach (var lm in list)
                {
                    DataSet1.過去勤務票明細Row lRow = (DataSet1.過去勤務票明細Row)dts.過去勤務票明細.Rows.Find(lm.ID);
                    lRow.Delete();
                }

                // 画像ファイルを削除します
                string imgPath = Properties.Settings.Default.tifPath + lh.画像名;
                File.Delete(imgPath);

                // 過去勤務票ヘッダを削除します
                lh.Delete();
            }

            // データベース更新
            pAdpMn.UpdateAll(dts);
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     設定月数分経過した過去画像を削除する</summary>
        /// <param name="_dYYMM">
        ///     基準年月 (例：201401)</param>
        /// -----------------------------------------------------------------------------
        private void deleteImageArchived(int _dYYMM)
        {
            int _DataYYMM;
            string fileYYMM;

            // 設定月数分経過した過去画像を削除する            
            foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.tifPath, "*.tif"))
            {
                // ファイル名が規定外のファイルは読み飛ばします
                if (System.IO.Path.GetFileName(files).Length < 21) continue;

                //ファイル名より年月を取得する
                fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                if (Utility.NumericCheck(fileYYMM))
                {
                    _DataYYMM = int.Parse(fileYYMM);

                    //基準年月以前なら削除する
                    if (_DataYYMM <= _dYYMM) File.Delete(files);
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダデータと勤務票明細データを全件削除します</summary>
        /// -------------------------------------------------------------------
        private void deleteDataAll()
        {
            // 出勤簿データ読み込み
            getDataSet();

            // 勤務票明細全行削除
            var m = dts.勤務票明細.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in m)
            {
                t.Delete();
            }

            // 勤務票ヘッダ全行削除
            var h = dts.勤務票ヘッダ.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in h)
            {
                t.Delete();
            }

            // データベース更新
            adpMn.UpdateAll(dts);

            // 後片付け
            dts.勤務票明細.Dispose();
            dts.勤務票ヘッダ.Dispose();
        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            //// 曜日
            //DateTime eDate;
            //int tYY = Utility.StrtoInt(txtYear.Text);
            //string sDate = tYY.ToString() + "/" + Utility.EmptytoZero(txtMonth.Text) + "/" +
            //        Utility.EmptytoZero(txtDay.Text);

            //// 存在する日付と認識された場合、曜日を表示する
            //if (DateTime.TryParse(sDate, out eDate))
            //{
            //    txtWeekDay.Text = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);
            //}
            //else
            //{
            //    txtWeekDay.Text = string.Empty;
            //}
        }

        private void dGV_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //if (editLogStatus)
            //{
            //    if (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 3 || e.ColumnIndex == 4 ||
            //        e.ColumnIndex == 6 || e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 10 ||
            //        e.ColumnIndex == 12 || e.ColumnIndex == 13 || e.ColumnIndex == 15)
            //    {
            //        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //        cellAfterValue = Utility.NulltoStr(dGV[e.ColumnIndex, e.RowIndex].Value);

            //        //// 変更のとき編集ログデータを書き込み
            //        //if (cellBeforeValue != cellAfterValue)
            //        //{
            //        //    logDataUpdate(e.RowIndex, cI, global.flgOn);
            //        //}
            //    }
            //}
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            //if (editLogStatus)
            //{
            //    if (sender == txtYear) cellName = LOG_YEAR;
            //    if (sender == txtMonth) cellName = LOG_MONTH;
            //    if (sender == txtDay) cellName = LOG_DAY;
            //    //if (sender == txtSftCode) cellName = LOG_TAIKEICD;

            //    TextBox tb = (TextBox)sender;

            //    // 値を保持
            //    cellBeforeValue = Utility.NulltoStr(tb.Text);
            //}
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            if (editLogStatus)
            {
                TextBox tb = (TextBox)sender;
                cellAfterValue = Utility.NulltoStr(tb.Text);

                //// 変更のとき編集ログデータを書き込み
                //if (cellBeforeValue != cellAfterValue)
                //{
                //    logDataUpdate(0, cI, global.flgOff);
                //}
            }
        }

        private void gcMultiRow1_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            // 過去データ表示のときは終了
            if (dID != string.Empty) return;

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

            // 過去データ表示のときは終了
            if (dID != string.Empty) return;

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

                        if (sStf.実労時間.Length > 4)
                        {
                            gcMultiRow2.SetValue(e.RowIndex, "txtWh", sStf.実労時間.Substring(0, 2));
                            gcMultiRow2.SetValue(e.RowIndex, "txtWm", sStf.実労時間.Substring(3, 2));
                        }
                        else
                        {
                            gcMultiRow2.SetValue(e.RowIndex, "txtWh", "");
                            gcMultiRow2.SetValue(e.RowIndex, "txtWm", "");
                        }

                        // 日給者のとき
                        if (sStf.給与区分.ToString() == global.NIKKYU)
                        {
                            // 基本実労働時間
                            gcMultiRow2[0, "txtWh"].Selectable = true;
                            gcMultiRow2[0, "txtWm"].Selectable = true;

                            // 基本就業時間帯
                            gcMultiRow2[0, "txtTm1Sh"].Selectable = true;
                            gcMultiRow2[0, "txtTm1Sm"].Selectable = true;
                            gcMultiRow2[0, "txtTm1Eh"].Selectable = true;
                            gcMultiRow2[0, "txtTm1Em"].Selectable = true;

                            gcMultiRow2[0, "txtTm2Sh"].Selectable = true;
                            gcMultiRow2[0, "txtTm2Sm"].Selectable = true;
                            gcMultiRow2[0, "txtTm2Eh"].Selectable = true;
                            gcMultiRow2[0, "txtTm2Em"].Selectable = true;

                            gcMultiRow2[0, "txtTm3Sh"].Selectable = true;
                            gcMultiRow2[0, "txtTm3Sm"].Selectable = true;
                            gcMultiRow2[0, "txtTm3Eh"].Selectable = true;
                            gcMultiRow2[0, "txtTm3Em"].Selectable = true;
                        }
                        else if (sStf.給与区分.ToString() == global.JIKKYU)
                        {
                            // 時給者のとき
                            // 基本実労働時間
                            gcMultiRow2[0, "txtWh"].Selectable = false;
                            gcMultiRow2[0, "txtWm"].Selectable = false;

                            // 基本就業時間帯
                            gcMultiRow2[0, "txtTm1Sh"].Selectable = false;
                            gcMultiRow2[0, "txtTm1Sm"].Selectable = false;
                            gcMultiRow2[0, "txtTm1Eh"].Selectable = false;
                            gcMultiRow2[0, "txtTm1Em"].Selectable = false;

                            gcMultiRow2[0, "txtTm2Sh"].Selectable = false;
                            gcMultiRow2[0, "txtTm2Sm"].Selectable = false;
                            gcMultiRow2[0, "txtTm2Eh"].Selectable = false;
                            gcMultiRow2[0, "txtTm2Em"].Selectable = false;

                            gcMultiRow2[0, "txtTm3Sh"].Selectable = false;
                            gcMultiRow2[0, "txtTm3Sm"].Selectable = false;
                            gcMultiRow2[0, "txtTm3Eh"].Selectable = false;
                            gcMultiRow2[0, "txtTm3Em"].Selectable = false;
                        }
                    }
                    else
                    {
                        // マスター未登録
                        gcMultiRow2.SetValue(e.RowIndex, "lblName", "");
                        gcMultiRow2.SetValue(e.RowIndex, "lblShopCode", "");
                        gcMultiRow2.SetValue(e.RowIndex, "lblShopName", "");

                        gcMultiRow2.SetValue(e.RowIndex, "txtWh", "");
                        gcMultiRow2.SetValue(e.RowIndex, "txtWm", "");
                    }
                }

                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;

                // 再表示：締日が異なることがあるため
                if (WorkTotalSumStatus)
                {
                    //カレントデータの更新
                    CurDataUpDate(cID[cI]);

                    // 再表示
                    showOcrData(cI);
                }
            }

            // 基本就労時間帯
            if (e.CellName == "txtTm1Sh" || e.CellName == "txtTm1Sm" || e.CellName == "txtTm1Eh" || e.CellName == "txtTm1Em" || 
                e.CellName == "txtTm2Sh" || e.CellName == "txtTm2Sm" || e.CellName == "txtTm2Eh" || e.CellName == "txtTm2Em" || 
                e.CellName == "txtTm3Sh" || e.CellName == "txtTm3Sm" || e.CellName == "txtTm3Eh" || e.CellName == "txtTm3Em")
            {
                // 再表示
                if (WorkTotalSumStatus)
                {
                    //カレントデータの更新
                    CurDataUpDate(cID[cI]);

                    // 再表示
                    showOcrData(cI);
                }
            }
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

        private void button1_Click(object sender, EventArgs e)
        {
            frmOCRIndex frm = new frmOCRIndex(dts, hAdp, stf, shp);
            frm.ShowDialog();
            string hID = frm.hdID;
            frm.Dispose();

            if (hID != string.Empty)
            {
                //カレントデータの更新
                CurDataUpDate(cID[cI]);

                // レコード検索
                for (int i = 0; i < cID.Length; i++)
                {
                    if (cID[i] == hID)
                    {
                        cI = i;
                        showOcrData(cI);
                        break;
                    }
                }
            }
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

                // 数字のみ入力可能とする：特休日数を含める 2020/05/12
                if (gcMultiRow3.CurrentCell.Name == "txtZdays" || gcMultiRow3.CurrentCell.Name == "txtKdays" ||
                    gcMultiRow3.CurrentCell.Name == "txtTlDays" || gcMultiRow3.CurrentCell.Name == "txtYuDays" ||
                    gcMultiRow3.CurrentCell.Name == "txtWorkTime" || gcMultiRow3.CurrentCell.Name == "txtWorkTime2" ||
                    gcMultiRow3.CurrentCell.Name == "txtNaiZan" || gcMultiRow3.CurrentCell.Name == "txtNaiZan2" ||
                    gcMultiRow3.CurrentCell.Name == "txtMashiZan" || gcMultiRow3.CurrentCell.Name == "txtMashiZan2" || 
                    gcMultiRow3.CurrentCell.Name == "txt20Zan" || gcMultiRow3.CurrentCell.Name == "txt20Zan2" || 
                    gcMultiRow3.CurrentCell.Name == "txtHolZan" || gcMultiRow3.CurrentCell.Name == "txtHolZan2" || 
                    gcMultiRow3.CurrentCell.Name == "txtChisou" || gcMultiRow3.CurrentCell.Name == "txtChisou2" || 
                    gcMultiRow3.CurrentCell.Name == "txtSonota" || gcMultiRow3.CurrentCell.Name == "txtSonota2" || 
                    gcMultiRow3.CurrentCell.Name == "txtKotuhi" || gcMultiRow3.CurrentCell.Name == "txtKotuhi2" ||
                    gcMultiRow3.CurrentCell.Name == "txtTokuDays")
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

        private void btnRtn_Click_1(object sender, EventArgs e)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 非ログ書き込み状態とする：2015/09/25
            editLogStatus = false;

            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName);

            // エラーチェックを実行
            if (getErrData(cI, ocr))
            {
                MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gcMultiRow1.CurrentCell = null;
                gcMultiRow2.CurrentCell = null;
                gcMultiRow3.CurrentCell = null;

                // データ表示
                showOcrData(cI);
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // データ表示
                showOcrData(cI);

                // エラー表示
                ErrShow(ocr);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の出勤簿を削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // 非ログ書き込み状態とする
            editLogStatus = false;

            // レコードと画像ファイルを削除する
            DataDelete(cI);

            // 勤務票ヘッダテーブル件数カウント
            if (dts.勤務票ヘッダ.Count() > 0)
            {
                // カレントレコードインデックスを再設定
                if (dts.勤務票ヘッダ.Count() - 1 < cI) cI = dts.勤務票ヘッダ.Count() - 1;

                // データ画面表示
                showOcrData(cI);
            }
            else
            {
                // ゼロならばプログラム終了
                MessageBox.Show("全ての出勤簿データが削除されました。処理を終了します。", "出勤簿削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Tag = END_NODATA;
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // PCA給与X用CSVデータ出力
            textDataMake();
        }

        private void gcMultiRow3_CellValueChanged(object sender, CellEventArgs e)
        {            
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            // 特休日数を含める：2020/05/12
            if (e.CellName == "txtZdays" || e.CellName == "txtKdays" || e.CellName == "txtYuDays" || e.CellName == "txtTokuDays")
            {
                // 
                int tDays = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtZdays"].Value)) +
                            Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtYuDays"].Value));

                // 特休日数を含める：2020/05/12
                tDays += Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[0, "txtTokuDays"].Value));

                gcMultiRow3.SetValue(0, "txtTlDays", tDays);
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の出勤簿を保留にします。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
            {
                return;
            }

            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //// データベース更新
            //adpMn.UpdateAll(dts);

            // 保留処理
            setHoldData(cID[cI]);
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     保留処理 </summary>
        /// <param name="iX">
        ///     データインデックス</param>
        ///----------------------------------------------------------
        private void setHoldData(string iX)
        {
            try
            {
                //var t = dts.勤務票ヘッダ.FindByID(iX);
                var t = dts.勤務票ヘッダ.Single(a => a.ID == iX);

                dAdpMn.保留勤務票ヘッダTableAdapter.Fill(dts.保留勤務票ヘッダ);

                DataSet1.保留勤務票ヘッダRow hr = dts.保留勤務票ヘッダ.New保留勤務票ヘッダRow();
                hr.ID = t.ID;
                hr.年 = t.年;
                hr.月 = t.月;
                hr.担当エリアマネージャー名 = t.担当エリアマネージャー名;
                hr.エリアコード = t.エリアコード;
                hr.エリア名 = t.エリア名;
                hr.店舗コード = t.店舗コード;
                hr.店舗名 = t.店舗名;
                hr.スタッフコード = t.スタッフコード;
                hr.氏名 = t.氏名;
                hr.給与形態 = t.給与形態;
                hr.要出勤日数 = t.要出勤日数;
                hr.実労日数 = t.実労日数;
                hr.有休日数 = t.有休日数;
                hr.公休日数 = t.公休日数;
                hr.遅早時間時 = t.遅早時間時;
                hr.遅早時間分 = t.遅早時間分;
                hr.実労働時間時 = t.実労働時間時;
                hr.実労働時間分 = t.実労働時間分;
                hr.基本時間内残業時 = t.基本時間内残業分;
                hr.基本時間内残業分 = t.基本時間内残業分;
                hr.割増残業時 = t.割増残業時;
                hr.割増残業分 = t.割増残業分;
                hr._20時以降勤務時 = t._20時以降勤務時;
                hr._20時以降勤務分 = t._20時以降勤務分;
                hr._22時以降勤務時 = t._22時以降勤務時;
                hr._22時以降勤務分 = t._22時以降勤務分;
                hr.土日祝日労働時間時 = t.土日祝日労働時間時;
                hr.土日祝日労働時間分 = t.土日祝日労働時間分;
                hr.交通費 = t.交通費;
                hr.その他支給 = t.その他支給;
                hr.画像名 = t.画像名;
                hr.確認 = t.確認;
                hr.備考 = t.備考;
                hr.編集アカウント = t.編集アカウント;
                hr.更新年月日 = DateTime.Now;
                hr.基本就業時間帯1開始時 = t.基本就業時間帯1開始時;
                hr.基本就業時間帯1開始分 = t.基本就業時間帯1開始分;
                hr.基本就業時間帯1終了時 = t.基本就業時間帯1終了時;
                hr.基本就業時間帯1終了分 = t.基本就業時間帯1終了分;
                hr.基本就業時間帯2開始時 = t.基本就業時間帯2開始時;
                hr.基本就業時間帯2開始分 = t.基本就業時間帯2開始分;
                hr.基本就業時間帯2終了時 = t.基本就業時間帯2終了時;
                hr.基本就業時間帯2終了分 = t.基本就業時間帯2終了分;
                hr.基本就業時間帯3開始時 = t.基本就業時間帯3開始時;
                hr.基本就業時間帯3開始分 = t.基本就業時間帯3開始分;
                hr.基本就業時間帯3終了時 = t.基本就業時間帯3終了時;
                hr.基本就業時間帯3終了分 = t.基本就業時間帯3終了分;
                hr.訂正1 = t.訂正1;
                hr.訂正2 = t.訂正2;
                hr.訂正3 = t.訂正3;
                hr.訂正4 = t.訂正4;
                hr.訂正5 = t.訂正5;
                hr.訂正6 = t.訂正6;
                hr.訂正7 = t.訂正7;
                hr.訂正8 = t.訂正8;
                hr.訂正9 = t.訂正9;
                hr.訂正10 = t.訂正10;
                hr.訂正11 = t.訂正11;
                hr.訂正12 = t.訂正12;
                hr.訂正13 = t.訂正13;
                hr.訂正14 = t.訂正14;
                hr.訂正15 = t.訂正15;
                hr.訂正16 = t.訂正16;
                hr.訂正17 = t.訂正17;
                hr.訂正18 = t.訂正18;
                hr.訂正19 = t.訂正19;
                hr.訂正20 = t.訂正20;
                hr.訂正21 = t.訂正21;
                hr.訂正22 = t.訂正22;
                hr.訂正23 = t.訂正23;
                hr.訂正24 = t.訂正24;
                hr.訂正25 = t.訂正25;
                hr.訂正26 = t.訂正26;
                hr.訂正27 = t.訂正27;
                hr.訂正28 = t.訂正28;
                hr.訂正29 = t.訂正29;
                hr.訂正30 = t.訂正30;
                hr.訂正31 = t.訂正31;
                hr.基本実労働時 = t.基本実労働時;
                hr.基本実労働分 = t.基本実労働分;
                hr.特休日数 = t.特休日数;   // 2020/05/12
                
                // 保留データ追加処理
                dts.保留勤務票ヘッダ.Add保留勤務票ヘッダRow(hr);

                // 保留勤務票明細
                dAdpMn.保留勤務票明細TableAdapter.Fill(dts.保留勤務票明細);

                var mm = dts.勤務票明細.Where(a => a.ヘッダID == iX).OrderBy(a => a.ID);
                foreach (var m in mm)
                {
                    DataSet1.保留勤務票明細Row mr = dts.保留勤務票明細.New保留勤務票明細Row();

                    mr.ヘッダID = m.ヘッダID;
                    mr.日 = m.日;
                    mr.出勤状況 = m.出勤状況;
                    mr.出勤時 = m.出勤時;
                    mr.出勤分 = m.出勤分;
                    mr.退勤時 = m.退勤時;
                    mr.退勤分 = m.退勤分;
                    mr.休憩 = m.休憩;
                    mr.有給申請 = m.有給申請;
                    mr.店舗コード = m.店舗コード;
                    mr.編集アカウント = m.編集アカウント;
                    mr.更新年月日 = DateTime.Now;
                    mr.暦年 = m.暦年;
                    mr.暦月 = m.暦月;

                    // 保留勤務票明細データ追加処理
                    dts.保留勤務票明細.Add保留勤務票明細Row(mr);
                }
                
                // データベース更新
                dAdpMn.UpdateAll(dts);
                
                // 出勤簿データ削除
                t.Delete();                 // 出勤簿ヘッダ
                foreach (var item in mm)    // 出勤簿明細
                {
                    item.Delete();
                }
                adpMn.UpdateAll(dts);

                // 配列キー再構築
                keyArrayCreate();

                // 終了メッセージ
                MessageBox.Show("出勤簿が保留されました", "出勤簿保留", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 件数カウント
                if (dts.勤務票ヘッダ.Count() > 0)
                {
                    // カレントレコードインデックスを再設定
                    if (dts.勤務票ヘッダ.Count() - 1 < cI)
                    {
                        cI = dts.勤務票ヘッダ.Count() - 1;
                    }

                    // データ画面表示
                    showOcrData(cI);
                }
                else
                {
                    // ゼロならばプログラム終了
                    MessageBox.Show("全ての出勤簿データが削除されました。処理を終了します。", "出勤簿削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //終了処理
                    this.Tag = END_NODATA;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Image img;

            img = Image.FromFile(b_img);
            e.Graphics.DrawImage(img, 0, 0);
            e.HasMorePages = false;

            img.Dispose();  // 後片付け 2019/01/15
        }
    }
}
