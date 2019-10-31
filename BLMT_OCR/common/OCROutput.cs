using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using BLMT_OCR.common;

namespace BLMT_OCR.common
{
    ///------------------------------------------------------------------
    /// <summary>
    ///     給与計算受け渡しデータ作成クラス </summary>     
    ///------------------------------------------------------------------
    class OCROutput
    {
        // 親フォーム
        Form _preForm;

        #region データテーブルインスタンス
        DataSet1.勤務票ヘッダDataTable _hTbl;
        DataSet1.勤務票明細DataTable _mTbl;
        #endregion

        private const string TXTFILENAME = "給与勤怠データ";

        DataSet1 _dts = new DataSet1();

        // ＰＣＡ給与分散データ（勤怠汎用データ）レイアウト
        public string c1 = string.Empty;    // 社員コード
        public string c2 = string.Empty;    // 要勤務日数
        public string c3 = ",";             // 要勤務時間
        public string c4 = string.Empty;    // 出勤日数
        public string c5 = string.Empty;    // 出勤時間
        public string c6 = "0,";            // 事故欠勤日数
        public string c7 = "0,";            // 病気欠勤日数
        public string c8 = ",";             // 代休特休時間
        public string c9 = ",";             // 休日出勤日数
        public string c10 = ",";            // 有休消化日数
        public string c11 = ",";            // 有休残日数
        public string c12 = string.Empty;   // 残業平日普通 ※割増残業
        public string c13 = string.Empty;   // 残業平日深夜 ※22時以降勤務時間
        public string c14 = ",";            // 残業休日普通
        public string c15 = ",";            // 残業休日深夜
        public string c16 = ",";            // 残業法定普通
        public string c17 = ",";            // 残業法定深夜
        public string c18 = ",";            // 遅刻早退回数
        public string c19 = string.Empty;   // 遅刻早退時間
        public string c20 = string.Empty;   // 有休日数消化
        public string c21 = ",";            // 有休時間消化
        public string c22 = ",";            // 有休日数残
        public string c23 = ",";            // 有休時間残
        public string c24 = ",";            // 有休可能時間
        public string c25 = ",";            // 残業平日普通４５下
        public string c26 = ",";            // 残業平日普通４５超
        public string c27 = ",";            // 残業平日普通６０超
        public string c28 = ",";            // 残業平日普通代休
        public string c29 = ",";            // 残業平日深夜４５下
        public string c30 = ",";            // 残業平日深夜４５超
        public string c31 = ",";            // 残業平日深夜６０超
        public string c32 = ",";            // 残業平日深夜代休
        public string c33 = ",";            // 残業平日普通４５下
        public string c34 = ",";            // 残業平日普通４５超
        public string c35 = ",";            // 残業平日普通６０超
        public string c36 = ",";            // 残業平日普通代休
        public string c37 = ",";            // 残業平日深夜４５下
        public string c38 = ",";            // 残業平日深夜４５超
        public string c39 = ",";            // 残業平日深夜６０超
        public string c40 = ",";            // 残業平日深夜代休

        public string[] c41 = new string[50];   // c41[0]:20時以降の勤務時間、c41[1]:土日祝の勤務時間

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     給与計算用計算用受入データ作成クラスコンストラクタ</summary>
        /// <param name="preFrm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     DataSet1</param>
        /// <param name="hTbl">
        ///     勤務票ヘッダDataTable</param>
        /// <param name="mTbl">
        ///     勤務票明細DataTable</param>
        ///--------------------------------------------------------------------------
        public OCROutput(Form preFrm, DataSet1 dts)
        {
            _preForm = preFrm;
            _dts = dts;
            _hTbl = dts.勤務票ヘッダ;
            _mTbl = dts.勤務票明細;

            for (int i = 0; i < c41.Length; i++)
            {
                if (i != c41.Length - 1) c41[i] = ",";
            }
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     給与受入データ作成</summary>
        ///--------------------------------------------------------------------------------------     
        public void SaveData()
        {
            string[] arrayCsv = null;     // 出力配列

            int sCnt = 0;   // 社員出力件数

            StringBuilder sb = new StringBuilder();
            Boolean pblFirstGyouFlg = true;
            string wID = string.Empty;
            string hKinmutaikei = string.Empty;

            global gl = new global();

            // 出力先フォルダがあるか？なければ作成する
            string cPath = Properties.Settings.Default.okPath;
            if (!System.IO.Directory.Exists(cPath)) System.IO.Directory.CreateDirectory(cPath);

            try
            {
                //オーナーフォームを無効にする
                _preForm.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = _preForm;
                frmP.Show();

                int rCnt = 1;

                // 伝票最初行フラグ
                pblFirstGyouFlg = true;

                // 勤務票データ取得
                foreach (var r in _hTbl.OrderBy(a => a.ID))
                {
                    // プログレスバー表示
                    frmP.Text = "PCA給与X受入データ作成中です・・・" + rCnt + "/" + _hTbl.Count;
                    frmP.progressValue = rCnt * 100 / _hTbl.Count;
                    frmP.ProgressStep();

                    //// ヘッダファイル出力
                    //if (pblFirstGyouFlg)
                    //{
                    //    // 配列にデータを出力
                    //    sCnt++;
                    //    Array.Resize(ref arrayCsv, sCnt);
                    //    arrayCsv[sCnt - 1] = sb.ToString();
                    //}

                    // 勤務票明細から受入れデータを作成する
                    string uCSV = getUkeireCsv(r);

                    // 配列にデータを格納します
                    sCnt++;
                    Array.Resize(ref arrayCsv, sCnt);
                    arrayCsv[sCnt - 1] = uCSV;

                    // データ件数加算
                    rCnt++;

                    pblFirstGyouFlg = false;
                }

                // 勤怠CSVファイル出力
                if (arrayCsv != null) txtFileWrite(cPath, arrayCsv);

                // いったんオーナーをアクティブにする
                _preForm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _preForm.Enabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("PCA給与X受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {

            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     勤務票明細から受入れデータ文字列を作成する </summary>
        /// <param name="r">
        ///     DataSet1.勤務票明細Row </param>
        /// <param name="dbName">
        ///     データベース名</param>
        /// <param name="hKinmutaikei">
        ///     勤務体系コード </param>
        /// <returns>
        ///     受入れデータ文字列</returns>
        ///-----------------------------------------------------------
        private string getUkeireCsv(DataSet1.勤務票ヘッダRow r)
        {
            string hDate = string.Empty;

            StringBuilder sb = new StringBuilder();

            sb.Clear();

            c1 = r.スタッフコード + ",";

            // 要勤務日数
            c2 = r.要出勤日数 + ",";

            // 出勤日数
            c4 = r.実労日数 + ",";
            
            // 総労働時間：10進数　2017/08/07
            //c5 = r.実労働時間時 + ":" + r.実労働時間分.ToString("D2") + ",";
            c5 = r.実労働時間時 + "." + get10shinsu(r.実労働時間分) + ",";
            
            // 割増残業時間：10進数　2017/08/07
            //c12 = r.割増残業時 + ":" + r.割増残業分.ToString("D2") + ",";
            c12 = r.割増残業時 + "." + get10shinsu(r.割増残業分) + ",";
            
            // 深夜勤務時間（２２時以降の勤務時間）：10進数　2017/08/07
            //c13 = r._22時以降勤務時 + ":" + r._22時以降勤務分.ToString("D2") + ",";
            c13 = r._22時以降勤務時 + "." + get10shinsu(r._22時以降勤務分) + ",";

            // 遅刻早退：10進数　2017/08/07
            //c19 = r.遅早時間時 + ":" + r.遅早時間分.ToString("D2") + ",";
            c19 = r.遅早時間時 + "." + get10shinsu(r.遅早時間分) + ",";

            // 有給日数
            c20 = r.有休日数 + ",";
            
            // 20時以降の勤務時間：10進数　2017/08/07
            //c41[0] = r._20時以降勤務時 + ":" + r._20時以降勤務分.ToString("D2") + ",";
            c41[0] = r._20時以降勤務時 + "." + get10shinsu(r._20時以降勤務分) + ",";

            // 土日・祝日勤務時間計：10進数　2017/08/07
            //c41[1] = r.土日祝日労働時間時 + ":" + r.土日祝日労働時間分.ToString("D2") + ",";
            c41[1] = r.土日祝日労働時間時 + "." + get10shinsu(r.土日祝日労働時間分) + ",";

            // 文字列を作成して返す
            sb.Append(c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9 + c10);
            sb.Append(c11 + c12 + c13 + c14 + c15 + c16 + c17 + c18 + c19 + c20);
            sb.Append(c21 + c22 + c23 + c24 + c25 + c26 + c27 + c28 + c29 + c30);
            sb.Append(c31 + c32 + c33 + c34 + c35 + c36 + c37 + c38 + c39 + c40);

            for (int i = 0; i < c41.Length; i++)
            {
                sb.Append(c41[i]);
            }

            return sb.ToString();
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     分を10進数に変換する </summary>
        /// <param name="m">
        ///     分</param>
        /// <returns>
        ///     10進数変換値</returns>
        ///-------------------------------------------------------
        private double get10shinsu(int m)
        {
            double mm = Utility.StrtoDouble(m.ToString());
            double cal = System.Math.Floor(mm / 60 * 100);
            return cal;
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     配列にテキストデータをセットする </summary>
        /// <param name="array">
        ///     社員、パート、出向社員の各配列</param>
        /// <param name="cnt">
        ///     拡張する配列サイズ</param>
        /// <param name="txtData">
        ///     セットする文字列</param>
        ///----------------------------------------------------------------------------
        private void txtArraySet(string [] array, int cnt, string txtData)
        {
            Array.Resize(ref array, cnt);   // 配列のサイズ拡張
            array[cnt - 1] = txtData;       // 文字列のセット
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string [] arrayData)
        {
            // 同名ファイルがあった場合、タイムスタンプを付加して保存
            if (System.IO.File.Exists(sPath + TXTFILENAME + ".dat"))
            {
                // 付加文字列（タイムスタンプ）
                string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                        DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                        DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                // 新ファイル名
                string outFileName = sPath + TXTFILENAME + newFileName + ".dat";
                System.IO.File.Move(sPath + TXTFILENAME + ".dat", outFileName);             
            }
            
            // テキストファイル出力
            File.WriteAllLines(sPath + TXTFILENAME + ".dat", arrayData, System.Text.Encoding.GetEncoding(932));
        }
    }
}
