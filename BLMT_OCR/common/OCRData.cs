using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace BLMT_OCR.common
{
    class OCRData
    {
        public OCRData(string dbName)
        {
            _dbName = dbName;
        }

        // 奉行シリーズデータ領域データベース名
        string _dbName = string.Empty;

        common.xlsData bs;

        #region エラー項目番号プロパティ
        //---------------------------------------------------
        //          エラー情報
        //---------------------------------------------------

        enum errCode
        {
            eNothing, eYearMonth, eMonth, eDay, eKinmuTaikeiCode
        }

        /// <summary>
        ///     エラーヘッダ行RowIndex</summary>
        public int _errHeaderIndex { get; set; }

        /// <summary>
        ///     エラー項目番号</summary>
        public int _errNumber { get; set; }

        /// <summary>
        ///     エラー明細行RowIndex </summary>
        public int _errRow { get; set; }

        /// <summary> 
        ///     エラーメッセージ </summary>
        public string _errMsg { get; set; }

        /// <summary> 
        ///     エラーなし </summary>
        public int eNothing = 0;

        /// <summary>
        ///     エラー項目 = 確認チェック </summary>
        public int eDataCheck = 35;

        /// <summary> 
        ///     エラー項目 = 対象年月日 </summary>
        public int eYearMonth = 1;

        /// <summary> 
        ///     エラー項目 = 対象月 </summary>
        public int eMonth = 2;

        /// <summary> 
        ///     エラー項目 = 日 </summary>
        public int eDay = 3;

        /// <summary> 
        ///     エラー項目 = 出勤状況 </summary>
        public int eShukkinStatus = 4;

        /// <summary> 
        ///     エラー項目 = 個人番号 </summary>
        public int eShainNo = 5;
        public int eShainNo2 = 27;

        /// <summary> 
        ///     エラー項目 = 勤務記号 </summary>
        public int eKintaiKigou = 6;

        /// <summary> 
        ///     エラー項目 = 部署コード </summary>
        public int eBushoCode = 31;

        /// <summary> 
        ///     エラー項目 = 応援チェック </summary>
        public int eChkOuen = 7;

        /// <summary> 
        ///     エラー項目 = シフト変更 </summary>
        public int eChksft = 8;

        /// <summary> 
        ///     エラー項目 = 事由 </summary>
        public int eJiyu1 = 9;
        public int eJiyu2 = 10;
        public int eJiyu3 = 11;

        /// <summary> 
        ///     エラー項目 = シフトコード </summary>
        public int eSftCode = 12;
        
        /// <summary> 
        ///     エラー項目 = 開始時 </summary>
        public int eSH = 13;

        /// <summary> 
        ///     エラー項目 = 開始分 </summary>
        public int eSM = 14;

        /// <summary> 
        ///     エラー項目 = 終了時 </summary>
        public int eEH = 15;

        /// <summary> 
        ///     エラー項目 = 終了分 </summary>
        public int eEM = 16;

        /// <summary> 
        ///     エラー項目 = 休憩 </summary>
        public int eRest = 17;

        /// <summary> 
        ///     エラー項目 = 有給申請 </summary>
        public int eYukyuCheck = 18;

        /// <summary> 
        ///     エラー項目 = 日別店舗コード </summary>
        public int eDailyShopCode = 19;

        /// <summary> 
        ///     エラー項目 = 実労日数 </summary>
        public int eWorkDays = 20;

        /// <summary> 
        ///     エラー項目 = 有給日数 </summary>
        public int eYukyuDays = 21;

        /// <summary> 
        ///     エラー項目 = 公休日数 </summary>
        public int eKoukyuDays = 22;

        /// <summary> 
        ///     エラー項目 = 遅刻早退 </summary>
        public int eChisou = 23;
        public int eLine2 = 28;

        /// <summary> 
        ///     エラー項目 = 部門 </summary>
        public int eBmn = 24;
        public int eBmn2 = 29;

        /// <summary> 
        ///     エラー項目 = 製品群 </summary>
        public int eHin = 25;
        public int eHin2 = 30;

        /// <summary> 
        ///     エラー項目 = 応援分 </summary>
        public int eOuenM = 26;

        /// <summary> 
        ///     エラー項目 = 応援分 </summary>
        public int eOuenIP = 32;
        public int eOuenIP2 = 33;

        /// <summary> 
        ///     エラー項目 = 応援移動票と勤怠データＩ／Ｐ票 </summary>
        public int eIpOuen = 34;

        #endregion
        
        #region 警告項目
        ///     <!--警告項目配列 -->
        public int[] warArray = new int[6];

        /// <summary>
        ///     警告項目番号</summary>
        public int _warNumber { get; set; }

        /// <summary>
        ///     警告明細行RowIndex </summary>
        public int _warRow { get; set; }

        /// <summary> 
        ///     警告項目 = 勤怠記号1&2 </summary>
        public int wKintaiKigou = 0;

        /// <summary> 
        ///     警告項目 = 開始終了時分 </summary>
        public int wSEHM = 1;

        /// <summary> 
        ///     警告項目 = 時間外時分 </summary>
        public int wZHM = 2;

        /// <summary> 
        ///     警告項目 = 深夜勤務時分 </summary>
        public int wSIHM = 3;

        /// <summary> 
        ///     警告項目 = 休日出勤時分 </summary>
        public int wKSHM = 4;

        /// <summary> 
        ///     警告項目 = 出勤形態 </summary>
        public int wShukeitai = 5;

        #endregion

        #region フィールド定義
        /// <summary> 
        ///     警告項目 = 時間外1.25時 </summary>
        public int [] wZ125HM = new int[global.MAX_GYO];

        /// <summary> 
        ///     実働時間 </summary>
        public double _workTime;

        /// <summary> 
        ///     深夜稼働時間 </summary>
        public double _workShinyaTime;
        #endregion

        #region 単位時間フィールド
        /// <summary> 
        ///     ３０分単位 </summary>
        private int tanMin30 = 30;

        /// <summary> 
        ///     １５分単位 </summary> 
        private int tanMin15 = 15;

        /// <summary> 
        ///     １０分単位 </summary> 
        private int tanMin10 = 10;

        /// <summary> 
        ///     １分単位 </summary>
        private int tanMin1 = 1;

        /// <summary> 
        ///     残業分単位０ </summary>
        private string zanMinTANI0 = "0";

        /// <summary> 
        ///     残業分単位５ </summary>
        private string zanMinTANI5 = "5";

        #endregion

        #region 時間チェック記号定数
        private const string cHOUR = "H";           // 時間をチェック
        private const string cMINUTE = "M";         // 分をチェック
        private const string cTIME = "HM";          // 時間・分をチェック
        #endregion

        private const string WKSPAN0750 = "7時間50分";
        private const string WKSPAN0755 = "7時間55分";
        private const string WKSPAN0800 = "8時間";
        private const string WKSPAN_KYUJITSU = "休日出勤";

        // 休憩時間
        private const Int64 RESTTIME0750 = 60;      // 7時間50分
        private const Int64 RESTTIME0755 = 65;      // 7時間55分
        private const Int64 RESTTIME0800 = 60;      // 8時間

        // テーブルアダプターマネージャーインスタンス
        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();
        
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _inPath, frmPrg frmP, string dbName)
        {
            string headerKey = string.Empty;    // ヘッダキー
            int shopCode = 0;     // 店舗コード

            // テーブルセットオブジェクト
            DataSet1 dts = new DataSet1();

            try
            {
                // 勤務表ヘッダデータセット読み込み
                DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
                adpMn.勤務票ヘッダTableAdapter = hAdp;
                adpMn.勤務票ヘッダTableAdapter.Fill(dts.勤務票ヘッダ);

                // 勤務表明細データセット読み込み
                DataSet1TableAdapters.勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();
                adpMn.勤務票明細TableAdapter = iAdp;
                adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);

                // 対象CSVファイル数を取得
                string [] t = System.IO.Directory.GetFiles(_inPath, "*.csv");
                int cLen = t.Length;

                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // ヘッダ行
                        if (stCSV[0] == "*")
                        {
                            // ヘッダーキー取得
                            headerKey = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
                            
                            // 店舗コード取得
                            shopCode = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 5));

                            // データセットに勤務票ヘッダデータを追加する
                            dts.勤務票ヘッダ.Add勤務票ヘッダRow(setNewHeadRecRow(dts, stCSV));
                        }
                        else　// 明細行
                        {
                            // データセットに勤務表明細データを追加する
                            dts.勤務票明細.Add勤務票明細Row(setNewItemRecRow(dts, headerKey, stCSV, shopCode));
                        }
                    }
                }

                // ローカルのデータベースを更新
                adpMn.UpdateAll(dts);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票ヘッダRowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセット</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <returns>
        ///     追加する勤務票ヘッダRowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private DataSet1.勤務票ヘッダRow setNewHeadRecRow(DataSet1 tblSt, string[] stCSV)
        {
            DataSet1.勤務票ヘッダRow r = tblSt.勤務票ヘッダ.New勤務票ヘッダRow();
            r.ID = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
            r.年 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 4));
            r.月 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2));
            r.スタッフコード = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 5));
            r.店舗コード = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 5));
            r.基本就業時間帯1開始時 = Utility.GetStringSubMax(stCSV[8].Trim().Replace("-", ""), 2);
            r.基本就業時間帯1開始分 = Utility.GetStringSubMax(stCSV[9].Trim().Replace("-", ""), 2);
            r.基本就業時間帯1終了時 = Utility.GetStringSubMax(stCSV[10].Trim().Replace("-", ""), 2);
            r.基本就業時間帯1終了分 = Utility.GetStringSubMax(stCSV[11].Trim().Replace("-", ""), 2);
            r.基本就業時間帯2開始時 = Utility.GetStringSubMax(stCSV[12].Trim().Replace("-", ""), 2);
            r.基本就業時間帯2開始分 = Utility.GetStringSubMax(stCSV[13].Trim().Replace("-", ""), 2);
            r.基本就業時間帯2終了時 = Utility.GetStringSubMax(stCSV[14].Trim().Replace("-", ""), 2);
            r.基本就業時間帯2終了分 = Utility.GetStringSubMax(stCSV[15].Trim().Replace("-", ""), 2);
            r.基本就業時間帯3開始時 = Utility.GetStringSubMax(stCSV[16].Trim().Replace("-", ""), 2);
            r.基本就業時間帯3開始分 = Utility.GetStringSubMax(stCSV[17].Trim().Replace("-", ""), 2);
            r.基本就業時間帯3終了時 = Utility.GetStringSubMax(stCSV[18].Trim().Replace("-", ""), 2);
            r.基本就業時間帯3終了分 = Utility.GetStringSubMax(stCSV[19].Trim().Replace("-", ""), 2);

            r.実労日数 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[20].Trim().Replace("-", ""), 2));
            r.公休日数 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[21].Trim().Replace("-", ""), 2));
            r.有休日数 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[22].Trim().Replace("-", ""), 2));
            r.要出勤日数 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[23].Trim().Replace("-", ""), 2));

            r.訂正1 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[24].Trim().Replace("-", ""), 1));
            r.訂正2 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[25].Trim().Replace("-", ""), 1));
            r.訂正3 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[26].Trim().Replace("-", ""), 1));
            r.訂正4 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[27].Trim().Replace("-", ""), 1));
            r.訂正5 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[28].Trim().Replace("-", ""), 1));
            r.訂正6 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[29].Trim().Replace("-", ""), 1));
            r.訂正7 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[30].Trim().Replace("-", ""), 1));
            r.訂正8 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[31].Trim().Replace("-", ""), 1));
            r.訂正9 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[32].Trim().Replace("-", ""), 1));
            r.訂正10 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[33].Trim().Replace("-", ""), 1));
            r.訂正11 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[34].Trim().Replace("-", ""), 1));
            r.訂正12 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[35].Trim().Replace("-", ""), 1));
            r.訂正13 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[36].Trim().Replace("-", ""), 1));
            r.訂正14 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[37].Trim().Replace("-", ""), 1));
            r.訂正15 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[38].Trim().Replace("-", ""), 1));
            r.訂正16 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[39].Trim().Replace("-", ""), 1));
            r.訂正17 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[40].Trim().Replace("-", ""), 1));
            r.訂正18 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[41].Trim().Replace("-", ""), 1));
            r.訂正19 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[42].Trim().Replace("-", ""), 1));
            r.訂正20 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[43].Trim().Replace("-", ""), 1));
            r.訂正21 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[44].Trim().Replace("-", ""), 1));
            r.訂正22 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[45].Trim().Replace("-", ""), 1));
            r.訂正23 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[46].Trim().Replace("-", ""), 1));
            r.訂正24 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[47].Trim().Replace("-", ""), 1));
            r.訂正25 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[48].Trim().Replace("-", ""), 1));
            r.訂正26 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[49].Trim().Replace("-", ""), 1));
            r.訂正27 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[50].Trim().Replace("-", ""), 1));
            r.訂正28 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[51].Trim().Replace("-", ""), 1));
            r.訂正29 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[52].Trim().Replace("-", ""), 1));
            r.訂正30 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[53].Trim().Replace("-", ""), 1));
            r.訂正31 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[54].Trim().Replace("-", ""), 1));
            
            r.担当エリアマネージャー名 = string.Empty;
            r.エリアコード = global.flgOff;
            r.エリア名 = string.Empty;
            r.店舗名 = string.Empty;
            r.氏名 = string.Empty;
            r.給与形態 = global.flgOff;
            r.遅早時間時 = global.flgOff;
            r.遅早時間分 = global.flgOff;
            r.実労働時間時 = global.flgOff;
            r.実労働時間分 = global.flgOff;
            r.基本時間内残業時 = global.flgOff;
            r.基本時間内残業分 = global.flgOff;
            r.割増残業時 = global.flgOff;
            r.割増残業分 = global.flgOff;
            r._20時以降勤務時 = global.flgOff;
            r._20時以降勤務分 = global.flgOff;
            r._22時以降勤務時 = global.flgOff;
            r._22時以降勤務分 = global.flgOff;
            r.土日祝日労働時間時 = global.flgOff;
            r.土日祝日労働時間分 = global.flgOff;
            r.交通費 = global.flgOff;
            r.その他支給 = global.flgOff;
            r.画像名 = Utility.GetStringSubMax(stCSV[1].Trim(), 21);
            r.確認 = global.flgOff;
            r.備考 = string.Empty;
            r.編集アカウント = global.loginUserID;
            r.更新年月日 = DateTime.Now;

            return r;
        }
        
        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票明細Rowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセットオブジェクト</param>
        /// <param name="headerKey">
        ///     ヘッダキー</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="sShopCode">
        ///     店舗コード</param>
        /// <returns>
        ///     追加する勤務票明細Rowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private DataSet1.勤務票明細Row setNewItemRecRow(DataSet1 tblSt, string headerKey, string[] stCSV, int sShopCode)
        {
            DataSet1.勤務票明細Row r = tblSt.勤務票明細.New勤務票明細Row();

            r.ヘッダID = headerKey;
            r.日 = 0;
            r.出勤状況 = Utility.GetStringSubMax(stCSV[0].Trim().Replace("-", ""), 1);
            r.出勤時 = Utility.GetStringSubMax(stCSV[1].Trim().Replace("-", ""), 2);
            r.出勤分 = Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 2);
            r.退勤時= Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2);
            r.退勤分 = Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2);
            r.休憩 = Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2);
            r.有給申請 = global.flgOff;

            if (r.出勤状況 != string.Empty || r.出勤時 != string.Empty || r.出勤分 != string.Empty ||
                 r.退勤時 != string.Empty || r.退勤分 != string.Empty)
            {
                r.店舗コード = sShopCode;
            }
            else
            {
                r.店舗コード = global.flgOff;
            }

            r.編集アカウント = global.loginUserID;
            r.更新年月日 = DateTime.Now;
            r.暦年 = global.flgOff;
            r.暦月 = global.flgOff;
            
            return r;
        }

        ///----------------------------------------------------------------------------------------
        /// <summary>
        ///     値1がemptyで値2がNot string.Empty のとき "0"を返す。そうではないとき値1をそのまま返す</summary>
        /// <param name="str1">
        ///     値1：文字列</param>
        /// <param name="str2">
        ///     値2：文字列</param>
        /// <returns>
        ///     文字列</returns>
        ///----------------------------------------------------------------------------------------
        private string hmStrToZero(string str1, string str2)
        {
            string rVal = str1;
            if (str1 == string.Empty && str2 != string.Empty)
                rVal = "0";

            return rVal;
        }


        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠データエラーチェックメイン処理。
        ///     エラーのときOCRDataクラスのヘッダ行インデックス、フィールド番号、明細行インデックス、
        ///     エラーメッセージが記録される </summary>
        /// <param name="sIx">
        ///     開始ヘッダ行インデックス</param>
        /// <param name="eIx">
        ///     終了ヘッダ行インデックス</param>
        /// <param name="frm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///-----------------------------------------------------------------------------------------------
        public Boolean errCheckMain(int sIx, int eIx, Form frm, DataSet1 dts, string[] cID, clsStaff sStf, clsStaff[] stf)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // レコード件数取得
            int cTotal = dts.勤務票ヘッダ.Rows.Count;

            // 出勤簿データ読み出し
            Boolean eCheck = true;

            for (int i = 0; i < cTotal; i++)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                //指定範囲ならエラーチェックを実施する：（i:行index）
                if (i >= sIx && i <= eIx)
                {
                    // 勤務票ヘッダ行のコレクションを取得します
                    DataSet1.勤務票ヘッダRow r = dts.勤務票ヘッダ.Single(a => a.ID == cID[i]);

                    // エラーチェック実施
                    eCheck = errCheckData(dts, r, stf);

                    if (!eCheck)　//エラーがあったとき
                    {
                        _errHeaderIndex = i;     // エラーとなったヘッダRowIndex
                        break;
                    }
                }
            }

            // いったんオーナーをアクティブにする
            frm.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            frm.Enabled = true;

            return eCheck;
        }
        
        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     エラー情報を取得します </summary>
        /// <param name="eID">
        ///     エラーデータのID</param>
        /// <param name="eNo">
        ///     エラー項目番号</param>
        /// <param name="eRow">
        ///     エラー明細行</param>
        /// <param name="eMsg">
        ///     表示メッセージ</param>
        ///---------------------------------------------------------------------------------
        private void setErrStatus(int eNo, int eRow, string eMsg)
        {
            //errHeaderIndex = eHRow;
            _errNumber = eNo;
            _errRow = eRow;
            _errMsg = eMsg;
        }


        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     勤務票ヘッダ行コレクション</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        /// 
        public Boolean errCheckData(DataSet1 dts, DataSet1.勤務票ヘッダRow r, clsStaff[] stf)
        {
            // 確認チェック
            if (r.確認 == global.flgOff)
            {
                setErrStatus(eDataCheck, 0, "未確認の出勤簿です");
                return false;
            }

            // 対象年月
            if ((2000 + r.年) != global.cnfYear)
            {
                setErrStatus(eYearMonth, 0, "設定された処理年月（" + global.cnfYear + "年" + global.cnfMonth + "月）と異なっています");
                return false;
            }

            if (r.月 != global.cnfMonth)
            {
                setErrStatus(eMonth, 0, "設定された処理年月（" + global.cnfYear + "年" + global.cnfMonth + "月）と異なっています");
                return false;
            }

            // スタッフコード
            if (!errCheckSNum(r, stf))
            {
                setErrStatus(eShainNo, 0, "マスター未登録のスタッフコードです");
                return false;
            }
            
            // 同じスタッフ番号の出勤簿が複数存在するときエラー
            if (!getSameNumber(dts, r.スタッフコード))
            {
                setErrStatus(eShainNo, 0, "同じスタッフ番号の出勤簿が複数あります");
                return false;
            }

            // 勤務実績チェック
            if (!errCheckNoWork(r)) return false;

            int iX = 0;

            // 勤務票明細データ行を取得
            List<DataSet1.勤務票明細Row> mList = dts.勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID).ToList();

            foreach (var m in mList)
            {
                // 行数
                iX++;

                // 対象外の行はチェック対象外とする
                if (m.日 == global.flgOff)
                {
                    continue;
                }

                // 無記入の行はチェック対象外とする
                if (m.出勤状況 == string.Empty && 
                    m.出勤時 == string.Empty && m.出勤分 == string.Empty && 
                    m.退勤時 == string.Empty && m.退勤分 == string.Empty && 
                    m.休憩 == string.Empty && m.有給申請 == global.flgOff)
                {
                    continue;
                }

                // 給与区分取得
                int kKbn = 0;
                foreach (var im in stf)
                {
                    if (im.スタッフコード == r.スタッフコード)
                    {
                        kKbn = im.給与区分;
                        break;
                    }
                }
                
                // 出勤状況：時給者のとき
                if (kKbn.ToString() == global.JIKKYU)
                {
                    if (m.出勤状況 != string.Empty && m.出勤状況 != global.STATUS_YUKYU && m.出勤状況 != global.STATUS_KOUKYU)
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "出勤状況の記入内容が正しくありません");
                        return false;
                    }
                }
                
                // 出勤状況
                if (m.出勤状況 != string.Empty)
                {
                    if (m.出勤状況 != global.STATUS_KIHON_1 &&
                        m.出勤状況 != global.STATUS_KIHON_2 &&
                        m.出勤状況 != global.STATUS_KIHON_3 &&
                        m.出勤状況 != global.STATUS_YUKYU &&
                        m.出勤状況 != global.STATUS_KOUKYU)
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "出勤状況が正しくありません");
                        return false;
                    }
                }

                if (m.出勤状況 == global.STATUS_KIHON_1)
                {
                    if (!isKihonTime(r.基本就業時間帯1開始時, r.基本就業時間帯1開始分))
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "基本時間帯１は有効な時間帯ではありません");
                        return false;
                    }

                    if (!isKihonTime(r.基本就業時間帯1終了時, r.基本就業時間帯1終了分))
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "基本時間帯１は有効な時間帯ではありません");
                        return false;
                    }
                }

                if (m.出勤状況 == global.STATUS_KIHON_2)
                {
                    if (!isKihonTime(r.基本就業時間帯2開始時, r.基本就業時間帯2開始分))
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "基本時間帯２は有効な時間帯ではありません");
                        return false;
                    }

                    if (!isKihonTime(r.基本就業時間帯2終了時, r.基本就業時間帯2終了分))
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "基本時間帯２は有効な時間帯ではありません");
                        return false;
                    }
                }

                if (m.出勤状況 == global.STATUS_KIHON_3)
                {
                    if (!isKihonTime(r.基本就業時間帯3開始時, r.基本就業時間帯3開始分))
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "基本時間帯３は有効な時間帯ではありません");
                        return false;
                    }

                    if (!isKihonTime(r.基本就業時間帯3終了時, r.基本就業時間帯3終了分))
                    {
                        setErrStatus(eShukkinStatus, iX - 1, "基本時間帯３は有効な時間帯ではありません");
                        return false;
                    }
                }
                
                // 明細記入チェック
                if (!errCheckRow(m, "出勤簿内容", iX)) return false;

                // 始業時刻・終業時刻チェック
                if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;

                // 休憩
                if (!errCheckRestTime(m, "休憩時間", iX)) return false;

                // 有給申請
                if (!errCheckYukyuCheck(m, "有給申請", iX)) return false;

                // 日別店舗コード
                if (!errCheckDailyShopCode(m, "日毎店舗コード", iX, stf)) return false;
            }
            
            // 実労日数チェック
            if (!errCheckWorkDaysTotal(r, mList)) return false;

            // 公休日数チェック
            if (!errCheckKoukyuDaysTotal(r, mList)) return false;

            // 有給日数チェック
            if (!errCheckYukyuDaysTotal(r, mList)) return false;

            // 要出勤日数チェック
            if (!errCheckTotalDays(r)) return false;

            // 遅刻早退チェック
            if (!errCheckChisou(r, stf)) return false;

            return true;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     基本就業時間帯が時間表記されているか </summary>
        /// <param name="sHH">
        ///     時</param>
        /// <param name="sMM">
        ///     分</param>
        /// <returns>
        ///     true:時間表記, false:時間表記でない</returns>
        ///-------------------------------------------------------------
        private bool isKihonTime(string sHH, string sMM)
        {
            DateTime sHHMM;
            bool rtn = false;

            if (DateTime.TryParse(sHH + ":" + sMM, out sHHMM))
            {
                rtn = true;
            }

            return rtn;
        }



        //private bool chkJiyuHas(DataSet1.勤務票明細Row m, string mJiyu)
        //{
        //    OCR.clsJiyuHas jiyu = new OCR.clsJiyuHas(mJiyu, _dbName);

        //    // マスター登録チェック
        //    if (!jiyu.isHasRows())
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}

        ///------------------------------------------------------------
        /// <summary>
        ///     スタッフコード存在チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="stf">
        ///     スタッフ情報クラス</param>
        /// <returns>
        ///     true:登録あり、false:登録なし</returns>
        ///------------------------------------------------------------
        private bool errCheckSNum(DataSet1.勤務票ヘッダRow r, clsStaff[] stf)
        {
            bool rtn = false;

            for (int i = 0; i < stf.Length; i++)
            {
                if (r.スタッフコード == stf[i].スタッフコード)
                {
                    rtn = true;
                    break;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="stKbn">
        ///     勤怠記号の出勤怠区分</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTime(DataSet1.勤務票明細Row m, string tittle, int Tani, int iX)
        {
            // 出勤時間と退勤時間
            string sTimeW = m.出勤時.Trim() + m.出勤分.Trim();
            string eTimeW = m.退勤時.Trim() + m.退勤分.Trim();

            if (sTimeW != string.Empty && eTimeW == string.Empty)
            {
                setErrStatus(eEH, iX - 1, tittle + "退勤時刻が未入力です");
                return false;
            }

            if (sTimeW == string.Empty && eTimeW != string.Empty)
            {
                setErrStatus(eSH, iX - 1, tittle + "出勤時刻が未入力です");
                return false;
            }

            // 記入のとき
            if (m.出勤時 != string.Empty || m.出勤分 != string.Empty ||
                m.退勤時 != string.Empty || m.退勤分 != string.Empty)
            {
                // 数字範囲、単位チェック
                if (!Utility.checkHourSpan(m.出勤時))
                {
                    setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.出勤分, Tani))
                {
                    setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                //if (!Utility.checkHourSpan(m.退勤時))
                //{
                //    setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
                //    return false;
                //}
                
                // 終了時刻範囲
                if (Utility.StrtoInt(Utility.NulltoStr(m.退勤時)) > 47 || 
                    (Utility.StrtoInt(Utility.NulltoStr(m.退勤時)) == 47 &&
                    Utility.StrtoInt(Utility.NulltoStr(m.退勤分)) > 0))
                {
                    setErrStatus(eEH, iX - 1, tittle + "終了時刻範囲を超えています（～４７：００）");
                    return false;
                }

                if (!Utility.checkMinSpan(m.退勤分, Tani))
                {
                    setErrStatus(eEM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if ((Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分)) >=
                    (Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分)))
                {
                    setErrStatus(eSH, iX - 1, "出勤時刻が退勤時刻以後になっています");
                    return false;
                }
            }

            return true;
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     休憩時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRestTime(DataSet1.勤務票明細Row m, string tittle, int iX)
        {
            // 出退勤時間未記入
            string sTimeW = m.出勤時.Trim() + m.出勤分.Trim();
            string eTimeW = m.退勤時.Trim() + m.退勤分.Trim();

            if (sTimeW == string.Empty && eTimeW == string.Empty)
            {
                if (m.休憩 != string.Empty)
                {
                    setErrStatus(eRest, iX - 1, "出退勤時刻が未入力で休憩が入力されています");
                    return false;
                }
            }

            // 出勤～退勤時間
            DateTime stm;
            DateTime etm;

            bool sb = DateTime.TryParse(m.出勤時 + ":" + m.出勤分, out stm);
            bool ed = DateTime.TryParse(m.退勤時 + ":" + m.退勤分, out etm);

            if (sb && ed)
            {
                double w = Utility.GetTimeSpan(stm, etm).TotalMinutes;
                if (Utility.StrtoDouble(m.休憩) >= w)
                {
                    setErrStatus(eRest, iX - 1, "休憩時間が出勤～退勤時間以上になっています");
                    return false;
                }
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     有給申請チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckYukyuCheck(DataSet1.勤務票明細Row m, string tittle, int iX)
        {
            if (m.出勤状況 == global.STATUS_YUKYU)
            {
                if (m.有給申請 == global.flgOff)
                {
                    setErrStatus(eYukyuCheck, iX - 1, "有給申請が未チェックです");
                    return false;
                }
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     日別店舗コードチェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション </param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="stf">
        ///     スタッフ情報配列</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckDailyShopCode(DataSet1.勤務票明細Row m, string tittle, int iX, clsStaff[] stf)
        {
            bool rtn = false;

            // 日付対象外行はチェックしない
            if (m.日 == global.flgOff)
            {
                return true;
            }

            // 非稼働日はチェックしない
            if (m.出勤状況 == string.Empty && m.出勤時 == string.Empty)
            {
                return true;
            }

            // 公休日と有休日はチェックしない
            if (m.出勤状況 == global.STATUS_KOUKYU || m.出勤状況 == global.STATUS_YUKYU)
            {
                return true;
            }

            // 無記入のとき
            if (m.店舗コード == global.flgOff)
            {
                setErrStatus(eDailyShopCode, iX - 1, "就労店舗コードが無記入です");
                return false;
            }

            // 記入された就労店舗コードのマスター参照
            for (int i = 0; i < stf.Length; i++)
            {
                if (m.店舗コード == stf[i].店舗コード)
                {
                    rtn = true;
                    break;
                }
            }

            if (!rtn)
            {
                setErrStatus(eDailyShopCode, iX - 1, "登録されていない店舗コードです");
            }

            return rtn;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実労日数チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="mList">
        ///     List<DataSet1.勤務票明細Row> </param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckWorkDaysTotal(DataSet1.勤務票ヘッダRow r, List<DataSet1.勤務票明細Row> mList)
        {
            int wdays = mList.Count(a => a.出勤状況 == global.STATUS_KIHON_1 || a.出勤状況 == global.STATUS_KIHON_2 || 
                                    a.出勤状況 == global.STATUS_KIHON_3 || a.出勤時 != string.Empty);

            if (wdays != r.実労日数)
            {
                setErrStatus(eWorkDays, 0, "実労日数が正しくありません（" + wdays +"日）");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤務実績チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckNoWork(DataSet1.勤務票ヘッダRow r)
        {
            if (r.公休日数 == global.flgOff && r.有休日数 == global.flgOff && r.実労日数 == global.flgOff)
            {
                setErrStatus(eWorkDays, 0, "勤務実績のない出勤簿は登録できません");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     合計日数チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTotalDays(DataSet1.勤務票ヘッダRow r)
        {
            if ((r.有休日数 + r.実労日数) != r.要出勤日数)
            {
                setErrStatus(eWorkDays, 0, "合計日数が正しくありません（" + (r.有休日数 + r.実労日数) + "日）");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     合計日数チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckChisou(DataSet1.勤務票ヘッダRow r, clsStaff [] stf)
        {
            if (r.遅早時間時 == global.flgOff && r.遅早時間分 == global.flgOff)
            {
                return true;
            }

            if (r.遅早時間分 > 60)
            {
                setErrStatus(eChisou, 0, "遅刻早退が正しくありません");
                return false;
            }

            for (int i = 0; i < stf.Length; i++)
            {
                if (r.スタッフコード == stf[i].スタッフコード)
                {
                    if (stf[i].店舗区分.ToString() != global.TENPO_CHOKUEI)
                    {
                        setErrStatus(eChisou, 0, "直営店以外で遅刻早退の記入はできません");
                        return false;
                    }

                    if (stf[i].給与区分.ToString() == global.JIKKYU && stf[i].締日区分.ToString() == global.SHIME_20)
                    {
                        setErrStatus(eChisou, 0, "直営店の20日締め時給者は遅刻早退の記入はできません");
                        return false;
                    }

                    break;
                }
            }


            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     有給日数チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="mList">
        ///     List<DataSet1.勤務票明細Row> </param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckYukyuDaysTotal(DataSet1.勤務票ヘッダRow r, List<DataSet1.勤務票明細Row> mList)
        {
            int ydays = mList.Count(a => a.出勤状況 == global.STATUS_YUKYU);

            if (ydays != r.有休日数)
            {
                setErrStatus(eYukyuDays, 0, "有給日数が正しくありません（" + ydays + "日）");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     公休日数チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="mList">
        ///     List<DataSet1.勤務票明細Row> </param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckKoukyuDaysTotal(DataSet1.勤務票ヘッダRow r, List<DataSet1.勤務票明細Row> mList)
        {
            int kdays = mList.Count(a => a.出勤状況 == global.STATUS_KOUKYU);

            if (kdays != r.公休日数)
            {
                setErrStatus(eKoukyuDays, 0, "公休日数が正しくありません（" + kdays + "日）");
                return false;
            }

            return true;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     検索用DepartmentCodeを取得する </summary>
        /// <returns>
        ///     DepartmentCode</returns>
        ///----------------------------------------------------------
        private string getDepartmentCode(string bCode)
        {
            string strCode = "";

            // DepartmentCode（部署コード）
            if (Utility.NumericCheck(bCode))
            {
                strCode = bCode.PadLeft(15, '0');
            }
            else
            {
                strCode = bCode.PadRight(15, ' ');
            }

            return strCode;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     残業分単位 </summary>
        /// <param name="zM">
        ///     残業分</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------------
        private bool chkZangyoMin(string zM)
        {
            bool rtn = true;

            if (zM != string.Empty)
            {
                if (zM != zanMinTANI0 && zM != zanMinTANI5)
                {
                    rtn = false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     社員コードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     社員コード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkShainCode(string s)
        {
            bool dm = false;

            ////// 奉行SQLServer接続文字列取得
            ////string sc = sqlControl.obcConnectSting.get(_dbName);
            ////sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            //// 社員コード取得
            //StringBuilder sb = new StringBuilder();
            //sb.Clear();
            //sb.Append("select EmployeeNo,RetireCorpDate from tbEmployeeBase ");
            //sb.Append("where EmployeeNo = '" + s.PadLeft(10, '0') + "'");
            //sb.Append(" and BeOnTheRegisterDivisionID != 9");

            //SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            //while (dR.Read())
            //{
            //    dm = true;
            //    break;
            //}

            //dR.Close();

            //sdCon.Close();

            return dm;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     シフトコードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     シフトコード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkSftCode(string s)
        {
            bool dm = false;

            ////// 奉行SQLServer接続文字列取得
            ////string sc = sqlControl.obcConnectSting.get(_dbName);
            ////sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            //// 勤務体系（シフト）コード取得
            //StringBuilder sb = new StringBuilder();
            //sb.Clear();
            //sb.Append("select LaborSystemCode, LaborSystemName from tbLaborSystem ");
            //sb.Append("where LaborSystemCode = '" + s.ToString().PadLeft(4, '0') + "'");
            
            //SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            //while (dR.Read())
            //{
            //    dm = true;
            //    break;
            //}

            //dR.Close();

            //sdCon.Close();

            return dm;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///      勤務票ヘッダデータで指定した社員番号の件数を調べる </summary>
        /// <param name="dts">
        ///     勤務票ヘッダデータセット</param>
        /// <param name="sNum">
        ///     スタッフ番号</param>
        /// <returns>
        ///     件数</returns>
        ///------------------------------------------------------------------------------------
        private bool getSameNumber(DataSet1 dts, int sNum)
        {
            bool rtn = true;

            //if (sNum == string.Empty) return rtn;

            // 指定した社員番号の件数を調べる
            if (dts.勤務票ヘッダ.Count(a => a.スタッフコード == sNum) > 1)
            {
                rtn = false;
            }

            return rtn;
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     明細記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     行を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRow(DataSet1.勤務票明細Row m, string tittle, int iX)
        {
            //// 社員番号以外に記入項目なしのときエラーとする
            //if (m.社員番号 != string.Empty && m.出勤時 == string.Empty &&
            //    m.出勤分 == string.Empty && m.退勤時 == string.Empty &&
            //    m.退勤分 == string.Empty && m.残業理由1 == string.Empty &&
            //    m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
            //    m.残業理由2 == string.Empty && m.残業時2 == string.Empty && 
            //    m.残業分2 == string.Empty && m.事由1 == string.Empty && 
            //    m.事由2 == string.Empty && m.事由3 == string.Empty && 
            //    m.応援 == global.FLGOFF && m.シフト通り == global.FLGOFF &&
            //    m.シフトコード == string.Empty)
            //{
            //    setErrStatus(eSH, iX - 1, tittle + "が未入力です");
            //    return false;
            //}

            if (m.出勤状況 == global.STATUS_KIHON_1 || m.出勤状況 == global.STATUS_KIHON_2 ||
                m.出勤状況 == global.STATUS_KIHON_3)
            {
                // 出勤状況と出退勤時刻
                if (m.出勤時 != string.Empty && m.出勤分 != string.Empty &&
                    m.退勤時 != string.Empty && m.退勤分 != string.Empty)
                {
                    setErrStatus(eShukkinStatus, iX - 1, "出勤状況と出退勤時刻が両方記入されています");
                    return false;
                }

                // 出勤状況と休憩
                if (m.休憩 != string.Empty)
                {
                    setErrStatus(eRest, iX - 1, "出勤状況と休憩が両方記入されています");
                    return false;
                }
            }

            if (m.出勤状況 == global.STATUS_YUKYU || m.出勤状況 == global.STATUS_KOUKYU)
            {
                // 公休、有休と出退勤時刻
                if (m.出勤時 != string.Empty && m.出勤分 != string.Empty &&
                    m.退勤時 != string.Empty && m.退勤分 != string.Empty)
                {
                    setErrStatus(eShukkinStatus, iX - 1, "公休または有休で出退勤時刻が記入されています");
                    return false;
                }

                // 公休、有休と休憩
                if (m.休憩 != string.Empty)
                {
                    setErrStatus(eRest, iX - 1, "公休または有休で休憩が記入されています");
                    return false;
                }
            }
            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間外記入チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="zan">
        ///     算出残業時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckZanTm(DataSet1.勤務票明細Row m, string tittle, int iX, Int64 zan)
        {
            //Int64 mZan = 0;

            //mZan = (Utility.StrtoInt(m.時間外時) * 60) + Utility.StrtoInt(m.時間外分);

            //// 記入時間と計算された残業時間が不一致のとき
            //if (zan != mZan)
            //{
            //    Int64 hh = zan / 60;
            //    Int64 mm = zan % 60;

            //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません。（" + hh.ToString() + "時間" + mm.ToString() + "分）");
            //    return false;
            //}

            return true;
        }

        /// ----------------------------------------------------------------------------------
        /// <summary>
        ///     時間外算出 2015/09/16 </summary>
        /// <param name="m">
        ///     SCCSDataSet.勤務票明細Row </param>
        /// <param name="Tani">
        ///     丸め単位・分</param>
        /// <param name="ws">
        ///     1日の所定労働時間</param>
        /// <returns>
        ///     時間外・分</returns>
        /// ----------------------------------------------------------------------------------
        public Int64 getZangyoTime(DataSet1.勤務票明細Row m, Int64 Tani, Int64 ws, Int64 restTime, out Int64 s10Rest, int taikeiCode)
        {
            Int64 zan = 0;  // 計算後時間外勤務時間
            s10Rest = 0;    // 深夜勤務時間帯の10分休憩時間

            DateTime cTm;
            DateTime sTm;
            DateTime eTm;
            DateTime zsTm;
            DateTime pTm;

            if (!m.Is出勤時Null() && !m.Is出勤分Null() && !m.Is出勤時Null() && !m.Is出勤分Null())
            {
                int ss = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);
                int ee = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);
                DateTime dt = DateTime.Today;
                string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

                // 始業時刻
                if (DateTime.TryParse(sToday + " " + m.出勤時 + ":" + m.出勤分, out cTm))
                {
                    sTm = cTm;
                }
                else return 0;

                // 終業時刻
                if (ss > ee)
                {
                    // 翌日
                    dt = DateTime.Today.AddDays(1);
                    sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
                    if (DateTime.TryParse(sToday + " " + m.退勤時 + ":" + m.退勤分, out cTm))
                    {
                        eTm = cTm;
                    }
                    else return 0;
                }
                else
                {
                    // 同日
                    if (DateTime.TryParse(sToday + " " + m.退勤時 + ":" + m.退勤分, out cTm))
                    {
                        eTm = cTm;
                    }
                    else return 0;
                }


                //MessageBox.Show(sTm.ToShortDateString() + " " + sTm.ToShortTimeString() + "    " + eTm.ToShortDateString() + " " + eTm.ToShortTimeString());


                // 作業日報に記入されている始業から就業までの就業時間取得
                double w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes - restTime;

                // 所定労働時間内なら時間外なし
                if (w <= ws)
                {
                    return 0;
                }

                // 所定労働時間＋休憩時間＋10分または15分経過後の時刻を取得（時間外開始時刻）
                zsTm = sTm.AddMinutes(ws);          // 所定労働時間
                zsTm = zsTm.AddMinutes(restTime);   // 休憩時間
                int zSpan = 0;

                if (taikeiCode == 100)
                {
                    zsTm = zsTm.AddMinutes(10);         // 体系コード：100 所定労働時間後の10分休憩
                    zSpan = 130;
                }
                else if (taikeiCode == 200 || taikeiCode == 300)
                {
                    zsTm = zsTm.AddMinutes(15);         // 体系コード：200,300 所定労働時間後の15分休憩
                    zSpan = 135;
                }
                
                pTm = zsTm;                         // 時間外開始時刻

                // 該当時刻から終業時刻まで130分または135分以上あればループさせる
                while (Utility.GetTimeSpan(pTm, eTm).TotalMinutes > zSpan)
                {
                    // 終業時刻まで2時間につき10分休憩として時間外を算出
                    // 時間外として2時間加算
                    zan += 120;

                    // 130分、または135分後の時刻を取得（2時間＋10分、または15分）
                    pTm = pTm.AddMinutes(zSpan);

                    // 深夜勤務時間中の10分または15分休憩時間を取得する
                    s10Rest += getShinya10Rest(pTm, eTm, zSpan - 120);
                }

                // 130分（135分）以下の時間外を加算
                zan += (Int64)Utility.GetTimeSpan(pTm, eTm).TotalMinutes;

                // 単位で丸める
                zan -= (zan % Tani);

                //MessageBox.Show(pTm.ToShortDateString() + "    " + eTm.ToShortDateString());
            }
                        
            return zan;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間中の10分休憩時間を取得する </summary>
        /// <param name="pTm">
        ///     時刻</param>
        /// <param name="eTm">
        ///     終業時刻</param>
        /// <param name="taikeiRest">
        ///     勤務体系別の休憩時間(10分または15分）</param>
        /// <returns>
        ///     休憩時間</returns>
        /// --------------------------------------------------------------------
        private int getShinya10Rest(DateTime pTm, DateTime eTm, int taikeiRest)
        {
            int restTime = 0;

            // 130(135)分後の時刻が終業時刻以内か
            TimeSpan ts = eTm.TimeOfDay;
            
            if (pTm <= eTm)
            {
                // 時刻が深夜時間帯か？
                if (pTm.Hour >= 22 || pTm.Hour <= 5)
                {
                    if (pTm.Hour == 22)
                    {
                        // 22時帯は22時以降の経過分を対象とします。
                        // 例）21:57～22:07のとき22時台の7分が休憩時間
                        if (pTm.Minute >= taikeiRest)
                        {
                            restTime = taikeiRest;
                        }
                        else
                        {
                            restTime = pTm.Minute;
                        }
                    }
                    else if (pTm.Hour == 5)
                    {
                        // 4時帯の経過分を対象とするので5時帯は減算します。
                        // 例）4:57～5:07のとき5時台の7分は差し引いて3分が休憩時間
                        if (pTm.Minute < taikeiRest)
                        {
                            restTime = (taikeiRest - pTm.Minute);
                        }
                    }
                    else
                    {
                        restTime = taikeiRest;
                    }
                }
            }

            return restTime;
        }


        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務記入チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShinya(DataSet1.勤務票明細Row m, string tittle, int Tani, int iX)
        {
        //    // 無記入なら終了
        //    if (m.深夜時 == string.Empty && m.深夜分 == string.Empty) return true;

        //    //  始業、終業時刻が無記入で深夜が記入されているときエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.深夜時 != string.Empty)
        //        {
        //            setErrStatus(eSIH, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.深夜分 != string.Empty)
        //        {
        //            setErrStatus(eSIM, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入のとき
        //    if (m.深夜時 != string.Empty || m.深夜分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        //if (!checkHourSpan(m.時間外時))
        //        //{
        //        //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません");
        //        //    return false;
        //        //}

        //        if (!checkMinSpan(m.深夜分, Tani))
        //        {
        //            setErrStatus(eSIM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //            return false;
        //        }
        //    }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="shinya">
        ///     算出された深夜k勤務時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShinyaTm(DataSet1.勤務票明細Row m, string tittle, int iX, Int64 shinya)
        {
            Int64 mShinya = 0;

            //mShinya = (Utility.StrtoInt(m.深夜時) * 60) + Utility.StrtoInt(m.深夜分);

            //// 記入時間と計算された深夜時間が不一致のとき
            //if (shinya != mShinya)
            //{
            //    Int64 hh = shinya / 60;
            //    Int64 mm = shinya % 60;

            //    setErrStatus(eSIH, iX - 1, tittle + "が正しくありません。（" + hh.ToString() + "時間" + mm.ToString() + "分）");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実働時間を取得する</summary>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="rH">
        ///     休憩時間・分</param>
        /// <returns>
        ///     実働時間</returns>
        ///------------------------------------------------------------------------------------
        public double getWorkTime(string sH, string sM, string eH, string eM, int rH)
        {
            DateTime sTm;
            DateTime eTm;
            DateTime cTm;
            double w = 0;   // 稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) || 
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            int ss = Utility.StrtoInt(sH) * 100 + Utility.StrtoInt(sM);
            int ee = Utility.StrtoInt(eH) * 100 + Utility.StrtoInt(eM);
            DateTime dt = DateTime.Today;
            string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

            // 開始時刻取得
            if (Utility.StrtoInt(sH) == 24)
            {
                if (DateTime.TryParse(sToday + " 0:" + Utility.StrtoInt(sM).ToString(), out cTm))
                {
                    sTm = cTm;
                }
                else return 0;
            }
            else
            {
                if (DateTime.TryParse(sToday + " " + Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
                {
                    sTm = cTm;
                }
                else return 0;
            }
            
            // 終業時刻
            if (ss > ee)
            {
                // 翌日
                dt = DateTime.Today.AddDays(1);
                sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
            }

            // 終了時刻取得
            if (Utility.StrtoInt(eH) == 24)
                eTm = DateTime.Parse(sToday + " 23:59");
            else
            {
                if (DateTime.TryParse(sToday + " " + Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
                {
                    eTm = cTm;
                }
                else return 0;
            }

            // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes + 1;
            }
            else if (sTm == eTm)    // 同時刻の場合は翌日の同時刻とみなす 2014/10/10
            {
                w = Utility.GetTimeSpan(sTm, eTm.AddDays(1)).TotalMinutes;  // 稼働時間
            }
            else
            {
                w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes;  // 稼働時間
            }

            // 休憩時間を差し引く
            if (w >= rH) w = w - rH;
            else w = 0;

            // 値を返す
            return w;
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間を取得する</summary>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="tani">
        ///     丸め単位</param>
        /// <param name="s10">
        ///     深夜勤務時間中の10分休憩</param>
        /// <returns>
        ///     深夜勤務時間・分</returns>
        /// ------------------------------------------------------------
        public double getShinyaWorkTime(string sH, string sM, string eH, string eM, int tani, Int64 s10)
        {
            DateTime sTime;
            DateTime eTime;
            DateTime cTm;

            double wkShinya = 0;    // 深夜稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) ||
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            // 開始時間を取得
            if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
            {
                sTime = cTm;
            }
            else return 0;

            // 終了時間を取得
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                eTime = global.dt2359;
            }
            else if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
            {
                eTime = cTm;
            }
            else return 0;


            // 当日内の勤務のとき
            if (sTime.TimeOfDay < eTime.TimeOfDay)
            {
                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    // 早朝時間帯稼働時間
                    if (eTime >= global.dt0500)
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                    }
                    else
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }
                }

                // 終了時刻が22:00以降のとき
                if (eTime >= global.dt2200)
                {
                    // 当日分の深夜帯稼働時間を求める
                    if (sTime <= global.dt2200)
                    {
                        // 出勤時刻が22:00以前のとき深夜開始時刻は22:00とする
                        wkShinya += Utility.GetTimeSpan(global.dt2200, eTime).TotalMinutes;
                    }
                    else
                    {
                        // 出勤時刻が22:00以降のとき深夜開始時刻は出勤時刻とする
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }

                    // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
                    if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
                        wkShinya += 1;
                }
            }
            else
            {
                // 日付を超えて終了したとき（開始時刻 >= 終了時刻）※2014/10/10 同時刻は翌日の同時刻とみなす

                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                }

                // 当日分の深夜勤務時間（～０：００まで）
                if (sTime <= global.dt2200)
                {
                    // 出勤時刻が22:00以前のとき無条件に120分
                    wkShinya += global.TOUJITSU_SINYATIME;
                }
                else
                {
                    // 出勤時刻が22:00以降のとき出勤時刻から24:00までを求める
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt2359).TotalMinutes + 1;
                }

                // 0:00以降の深夜勤務時間を加算（０：００～終了時刻）
                if (eTime.TimeOfDay > global.dt0500.TimeOfDay)
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, global.dt0500).TotalMinutes;
                }
                else
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, eTime).TotalMinutes;
                }
            }

            // 深夜勤務時間中の10分または15分休憩時間を差し引く
            wkShinya -= s10;

            // 単位分で丸め
            wkShinya -= (wkShinya % tani);

            return wkShinya;
        }
    }
}
