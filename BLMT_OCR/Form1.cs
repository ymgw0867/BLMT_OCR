using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using Leadtools.ImageProcessing.Core;
using BLMT_OCR.common;
using System.Data.OleDb;

namespace BLMT_OCR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // 環境設定項目よみこみ
            Config.getConfig cnf = new Config.getConfig();

            // ログインユーザー登録有るか？
            Config.clsLogin clgin = new Config.clsLogin();
            clgin.addAdminUser();

            // ログイン
            frmLogin fl = new frmLogin();
            fl.ShowDialog();

            // ログイン未完了の場合は終了する
            if (!global.loginStatus) Environment.Exit(0);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            /* 過去勤務票ヘッダ、勤務票ヘッダ、保留勤務票ヘッダ各テーブルに
            「特休日数」フィールドを追加 2020/05/12 */
            mdbAlter();


            // キャプションにバージョンを追加
            this.Text += "   ver " + Application.ProductVersion;
        }
        
        ///---------------------------------------------------------------------
        /// <summary>
        ///     2020/05/12 ローカルMDB 
        ///     過去勤務票ヘッダ、勤務票ヘッダ、保留勤務票ヘッダ各テーブルに
        ///     「特休日数」フィールドを追加 </summary>
        ///--------------------------------------------------------------------- 
        private void mdbAlter()
        {
            try
            {
                // ローカルデータベース接続
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = Utility.dbConnect();

                string sqlSTRING = string.Empty;

                // 過去勤務票ヘッダテーブルに「特休日数」フィールドを追加する : 2020/05/12
                sqlSTRING = "ALTER TABLE 過去勤務票ヘッダ ADD COLUMN 特休日数 int";
                sCom.CommandText = sqlSTRING;
                sCom.ExecuteNonQuery();

                // 勤務票ヘッダテーブルに「特休日数」フィールドを追加する : 2020/05/12
                sqlSTRING = "ALTER TABLE 勤務票ヘッダ ADD COLUMN 特休日数 int";
                sCom.CommandText = sqlSTRING;
                sCom.ExecuteNonQuery();

                // 保留勤務票ヘッダテーブルに「特休日数」フィールドを追加する : 2020/05/12
                sqlSTRING = "ALTER TABLE 保留勤務票ヘッダ ADD COLUMN 特休日数 int";
                sCom.CommandText = sqlSTRING;
                sCom.ExecuteNonQuery();

                sCom.Connection.Close();
            }
            catch (Exception)
            {
                
                //throw;
            }
        }

        private void btnPrePrint_Click(object sender, EventArgs e)
        {
            Hide();
            prePrint.frmPrePrint frm = new prePrint.frmPrePrint();
            frm.ShowDialog();
            Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Hide();
            config.frmConfig frm = new config.frmConfig();
            frm.ShowDialog();
            Show();

            // 環境設定項目よみこみ
            Config.getConfig cnf = new Config.getConfig();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // 閉じる
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Hide();
            config.frmCalendar frm = new config.frmCalendar();
            frm.ShowDialog();
            Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 環境設定年月の確認
            string msg = "処理対象年月は " + global.cnfYear.ToString() + "年 " + global.cnfMonth.ToString() + "月です。よろしいですか？";
            if (MessageBox.Show(msg, "勤怠データ作成", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            bool hol = true;

            Hide();

            if (getHoldData() > 0)
            {
                // 保留データ選択
                OCR.frmFaxSelect frmH = new OCR.frmFaxSelect();
                frmH.ShowDialog();
                hol = frmH.myBool;  // 中止のとき：false
                frmH.Dispose();
            }

            if (hol)
            {
                /* 出勤簿フォームを開く前に処理可能な出勤簿データがあるか確認してなければ終了する
                // 2017/10/24
                // CSVファイル数をカウント */
                int n = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv").Count();

                DataSet1 dts = new DataSet1();
                DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
                hAdp.Fill(dts.勤務票ヘッダ);

                // 勤務票件数カウント
                if (n == 0 && dts.勤務票ヘッダ.Count == 0)
                {
                    MessageBox.Show("出勤簿がありません", "出勤簿登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Show(); // メニュー画面再表示
                    return; // 戻る
                }

                // 勤怠データ作成
                OCR.frmCorrect frm = new OCR.frmCorrect(string.Empty);
                frm.ShowDialog();
            }

            Show();
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     保留データ件数取得 </summary>
        /// <returns>
        ///     保留データ件数</returns>
        ///------------------------------------------------------------
        private int getHoldData()
        {
            DataSet1 dts = new DataSet1();
            DataSet1TableAdapters.保留勤務票ヘッダTableAdapter adp = new DataSet1TableAdapters.保留勤務票ヘッダTableAdapter();

            adp.Fill(dts.保留勤務票ヘッダ);

            int cnt = dts.保留勤務票ヘッダ.Count();
            return cnt;
        }

        private void btnOCR_Click(object sender, EventArgs e)
        {
            Hide();
            doFaxOCR();
            Show();
        }

        private void doFaxOCR()
        {
            int cnt = System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif").Count();
            if (cnt == 0)
            {
                MessageBox.Show("現在、受信済みのＦＡＸ出勤簿はありません", "受信出勤簿", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                if (MessageBox.Show("現在、" + cnt + "件のＦＡＸ出勤簿が受信済みです。ＯＣＲ認識を行いますか？", "出勤簿ＯＣＲ認識確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                // ファイル名（日付時間部分）
                string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                        string.Format("{0:00}", DateTime.Today.Month) +
                        string.Format("{0:00}", DateTime.Today.Day) +
                        string.Format("{0:00}", DateTime.Now.Hour) +
                        string.Format("{0:00}", DateTime.Now.Minute) +
                        string.Format("{0:00}", DateTime.Now.Second);

                int dNum = 0;                       // ファイル名末尾連番

                // 2019/03/27 コメント化
                /* マルチTiff画像をシングルtifに分解後に解像度200×200、A4サイズ変換
                 * (SCANフォルダ → TRAYフォルダ) */
                //if (MultiTif(Properties.Settings.Default.scanPath, Properties.Settings.Default.trayPath, fName))
                //{
                //    // WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する
                //    WinReaderOCR();

                //    /* OCR認識結果ＣＳＶデータを出勤簿ごとに分割して
                //     * 画像ファイルと共にDATAフォルダへ移動する */
                //    LoadCsvDivide(fName, ref dNum);
                //}

                /* マルチTiff画像をシングルtifに分解：LZW圧縮形式に対応 2019/03/27
                 * (SCANフォルダ → TRAYフォルダ) */
                if (MultiTif_New(Properties.Settings.Default.scanPath, Properties.Settings.Default.trayPath))
                {
                    // 画像の解像度、サイズを変更する：2019/03/28
                    imageResize(Properties.Settings.Default.trayPath);

                    // WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する
                    WinReaderOCR();

                    /* OCR認識結果ＣＳＶデータを出勤簿ごとに分割して
                     * 画像ファイルと共にDATAフォルダへ移動する */
                    LoadCsvDivide(fName, ref dNum);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                MessageBox.Show("終了しました");
            }
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する </summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <param name="outPath">
        ///     分割後出力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool MultiTif(string InPath, string outPath, string fName)
        {
            //スキャン出力画像を確認
            if (System.IO.Directory.GetFiles(InPath, "*.tif").Count() == 0)
            {
                return false;
            }

            // 出力先フォルダがなければ作成する
            if (System.IO.Directory.Exists(outPath) == false)
            {
                System.IO.Directory.CreateDirectory(outPath);
            }

            // 出力先フォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(outPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();

            int _pageCount = 0;
            string fnm = string.Empty;

            //コマンドを準備します。(傾き・ノイズ除去・リサイズ)
            DeskewCommand Dcommand = new DeskewCommand();
            DespeckleCommand Dkcommand = new DespeckleCommand();
            SizeCommand Rcommand = new SizeCommand();

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            int cImg = System.IO.Directory.GetFiles(InPath, "*.tif").Count();
            int cCnt = 0;

            // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                cCnt++;

                //プログレスバー表示
                frmP.Text = "OCR変換画像データロード中　" + cCnt.ToString() + "/" + cImg;
                frmP.progressValue = cCnt * 100 / cImg;
                frmP.ProgressStep();

                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
                    //ページを移動する
                    leadImg.Page = i;

                    // ファイル名設定
                    _pageCount++;
                    fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                    //画像補正処理　開始 ↓ ****************************
                    try
                    {
                        //画像の傾きを補正します。
                        Dcommand.Flags = DeskewCommandFlags.DeskewImage | DeskewCommandFlags.DoNotFillExposedArea;
                        Dcommand.Run(leadImg);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(i + "画像の傾き補正エラー：" + ex.Message);
                    } 
                    
                    //ノイズ除去
                    try
                    {
                        Dkcommand.Run(leadImg);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(i + "ノイズ除去エラー：" + ex.Message);
                    }

                    //解像度調整(200*200dpi)
                    leadImg.XResolution = 200;
                    leadImg.YResolution = 200;

                    //A4縦サイズに変換(ピクセル単位)
                    Rcommand.Width = 1637;
                    Rcommand.Height = 2322;
                    try
                    {
                        Rcommand.Run(leadImg);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(i + "解像度調整エラー：" + ex.Message);
                    }

                    //画像補正処理　終了↑ ****************************

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.Tif, 0, i, i, 1, CodecsSavePageMode.Insert);
                }
            }

            //LEADTOOLS入出力ライブラリを終了します。
            RasterCodecs.Shutdown();

            // InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return true;
        }


        ///------------------------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する：OpenCVバージョン</summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <param name="outPath">
        ///     分割後出力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool MultiTif_New(string InPath, string outPath)
        {
            //スキャン出力画像を確認
            if (System.IO.Directory.GetFiles(InPath, "*.tif").Count() == 0)
            {
                MessageBox.Show("ＯＣＲ変換処理対象の画像ファイルが指定フォルダ " + InPath + " に存在しません", "スキャン画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            // 出力先フォルダがなければ作成する
            if (System.IO.Directory.Exists(outPath) == false)
            {
                System.IO.Directory.CreateDirectory(outPath);
            }

            // 出力先フォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(outPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            int _pageCount = 0;
            string fnm = string.Empty;

            // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                //TIFFのImageCodecInfoを取得する
                ImageCodecInfo ici = GetEncoderInfo("image/tiff");

                if (ici == null)
                {
                    return false;
                }

                using (FileStream tifFS = new FileStream(files, FileMode.Open, FileAccess.Read))
                {
                    Image gim = Image.FromStream(tifFS);

                    FrameDimension gfd = new FrameDimension(gim.FrameDimensionsList[0]);

                    //全体のページ数を得る
                    int pageCount = gim.GetFrameCount(gfd);

                    for (int i = 0; i < pageCount; i++)
                    {
                        gim.SelectActiveFrame(gfd, i);

                        // ファイル名（日付時間部分）
                        string fName = string.Format("{0:0000}", DateTime.Today.Year) + string.Format("{0:00}", DateTime.Today.Month) +
                                string.Format("{0:00}", DateTime.Today.Day) + string.Format("{0:00}", DateTime.Now.Hour) +
                                string.Format("{0:00}", DateTime.Now.Minute) + string.Format("{0:00}", DateTime.Now.Second);

                        _pageCount++;

                        // ファイル名設定
                        fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                        EncoderParameters ep = null;

                        // 圧縮方法を指定する
                        ep = new EncoderParameters(1);
                        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

                        // 画像保存
                        gim.Save(fnm, ici, ep);
                        ep.Dispose();
                    }
                }
            }

            // InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            return true;
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     画像の解像度とサイズを変更する </summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool imageResize(string InPath)
        {
            // 画像を確認
            if (System.IO.Directory.GetFiles(InPath, "*.tif").Count() == 0)
            {
                return false;
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();

            string fnm = string.Empty;

            //コマンドを準備します。(傾き・ノイズ除去・リサイズ)
            DeskewCommand Dcommand = new DeskewCommand();
            DespeckleCommand Dkcommand = new DespeckleCommand();
            SizeCommand Rcommand = new SizeCommand();

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            int cImg = System.IO.Directory.GetFiles(InPath, "*.tif").Count();
            int cCnt = 0;

            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                cCnt++;

                //プログレスバー表示
                frmP.Text = "OCR変換画像データロード中　" + cCnt.ToString() + "/" + cImg;
                frmP.progressValue = cCnt * 100 / cImg;
                frmP.ProgressStep();

                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                //画像補正処理　開始 ↓ ****************************
                try
                {
                    //画像の傾きを補正します。
                    Dcommand.Flags = DeskewCommandFlags.DeskewImage | DeskewCommandFlags.DoNotFillExposedArea;
                    Dcommand.Run(leadImg);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(i + "画像の傾き補正エラー：" + ex.Message);
                }

                //ノイズ除去
                try
                {
                    Dkcommand.Run(leadImg);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(i + "ノイズ除去エラー：" + ex.Message);
                }

                //解像度調整(200*200dpi)
                leadImg.XResolution = 200;
                leadImg.YResolution = 200;

                //A4縦サイズに変換(ピクセル単位)
                Rcommand.Width = 1637;
                Rcommand.Height = 2322;
                try
                {
                    Rcommand.Run(leadImg);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(i + "解像度調整エラー：" + ex.Message);
                }

                //画像補正処理　終了↑ ****************************

                // 画像保存
                cs.Save(leadImg, files, RasterImageFormat.Tif, 0, 1, 1, 1, CodecsSavePageMode.Overwrite);
            }

            //LEADTOOLS入出力ライブラリを終了します。
            RasterCodecs.Shutdown();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return true;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     MimeTypeで指定されたImageCodecInfoを探して返す </summary>
        /// <param name="mineType">
        ///     </param>
        /// <returns>
        ///     </returns>
        ///---------------------------------------------------------------------
        private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo(string mineType)
        {
            //GDI+ に組み込まれたイメージ エンコーダに関する情報をすべて取得
            System.Drawing.Imaging.ImageCodecInfo[] encs = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();
            //指定されたMimeTypeを探して見つかれば返す
            foreach (System.Drawing.Imaging.ImageCodecInfo enc in encs)
            {
                if (enc.MimeType == mineType)
                {
                    return enc;
                }
            }
            return null;
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     CSVファイルと画像ファイルの名前を日付スタンプに変更する </summary>
        /// <param name="readFilePath">
        ///     入力CSVファイル名(フルパス）</param>
        /// <param name="newFnm">
        ///     新ファイル名（フルパス・但し拡張子なし）</param>
        /// -----------------------------------------------------------------
        private void wrhOutFileRename(string readFilePath, string newFnm)
        {
            string imgName = string.Empty;      // 画像ファイル名
            string[] stArrayData;               // CSVファイルを１行単位で格納する配列
            string inFilePath = string.Empty;   // ＯＣＲ認識モードごとの入力ファイル名

            // CSVデータの存在を確認します
            if (!System.IO.File.Exists(readFilePath)) return;

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            System.IO.StreamReader inFile = new System.IO.StreamReader(readFilePath, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」があったらヘッダー情報
                if ((stArrayData[0] == "*"))
                {
                    //文字列バッファをクリア
                    stResult = string.Empty;

                    // 文字列再構成（画像ファイル名を変更する）
                    stBuffer = string.Empty;
                    for (int i = 0; i < stArrayData.Length; i++)
                    {
                        if (stBuffer != string.Empty)
                        {
                            stBuffer += ",";
                        }

                        // 画像ファイル名を変更する
                        if (i == 1)
                        {
                            stArrayData[i] = System.IO.Path.GetFileName(newFnm) + ".tif"; // 画像ファイル名を変更
                        }

                        // フィールド結合
                        string sta = stArrayData[i].Trim();
                        stBuffer += sta;
                    }

                    //// 再ＦＡＸか確認
                    //if (stArrayData[4] == global.FLGON)
                    //{
                    //    // 書き出しファイル名を再FAXフォルダにする
                    //    newFnm = Properties.Settings.Default.reFaxPath + System.IO.Path.GetFileName(newFnm);
                    //}
                }

                // 読み込んだものを追加で格納する
                stResult += (stBuffer + Environment.NewLine);
            }

            // CSVファイル書き出し
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(newFnm + ".csv",
                                                    false, System.Text.Encoding.GetEncoding(932));
            outFile.Write(stResult);

            // 出力ファイルを閉じる
            outFile.Close();

            // 入力ファイルを閉じる
            inFile.Close();

            // 入力ファイル削除 : "txtout.csv"
            string inPath = System.IO.Path.GetDirectoryName(readFilePath);
            Utility.FileDelete(inPath, Properties.Settings.Default.wrReaderOutFile);

            // 画像ファイルをリネーム
            System.IO.File.Move(Properties.Settings.Default.readPath + "WRH00001.tif", newFnm + ".tif");
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する </summary>
        ///----------------------------------------------------------------------------------
        private void WinReaderOCR()
        {
            // WinReaderJOB起動文字列
            string JobName = @"""" + Properties.Settings.Default.wrHands_Job + @"""" + " /H2";
            string winReader_exe = Properties.Settings.Default.wrHands_Path +
                Properties.Settings.Default.wrHands_Prg;

            // ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();

            // 起動するアプリケーションを設定する
            p.FileName = winReader_exe;

            // コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
            p.Arguments = JobName;

            // WinReaderを起動します
            System.Diagnostics.Process hProcess = System.Diagnostics.Process.Start(p);

            // taskが終了するまで待機する
            hProcess.WaitForExit();
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     伝票ＣＳＶデータを一枚ごとに分割する </summary>
        ///-----------------------------------------------------------------
        private void LoadCsvDivide(string fnm, ref int dNum)
        {
            string imgName = string.Empty;      // 画像ファイル名
            string firstFlg = global.FLGON;
            string[] stArrayData;               // CSVファイルを１行単位で格納する配列
            string newFnm = string.Empty;
            int dCnt = 0;   // 処理件数

            // 対象ファイルの存在を確認します
            if (!System.IO.File.Exists(Properties.Settings.Default.readPath + Properties.Settings.Default.wrReaderOutFile))
            {
                return;
            }

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            System.IO.StreamReader inFile = new System.IO.StreamReader(Properties.Settings.Default.readPath + Properties.Settings.Default.wrReaderOutFile, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;

            // 行番号
            int sRow = 0;

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」があったら新たな伝票なのでCSVファイル作成
                if ((stArrayData[0] == "*"))
                {
                    //最初の伝票以外のとき
                    if (firstFlg != global.FLGON)
                    {
                        //ファイル書き出し
                        outFileWrite(stResult, Properties.Settings.Default.readPath + imgName, newFnm);
                    }

                    firstFlg = global.FLGOFF;

                    // 伝票連番
                    dNum++;

                    // 処理件数
                    dCnt++;

                    // ファイル名
                    newFnm = fnm + dNum.ToString().PadLeft(3, '0');

                    //画像ファイル名を取得
                    imgName = stArrayData[1];

                    //文字列バッファをクリア
                    stResult = string.Empty;

                    // 文字列再校正（画像ファイル名を変更する）
                    stBuffer = string.Empty;
                    for (int i = 0; i < stArrayData.Length; i++)
                    {
                        if (stBuffer != string.Empty)
                        {
                            stBuffer += ",";
                        }

                        // 画像ファイル名を変更する
                        if (i == 1)
                        {
                            stArrayData[i] = newFnm + ".tif"; // 画像ファイル名を変更
                        }

                        //// 日付（６桁）を年月日（２桁毎）に分割する
                        //if (i == 3)
                        //{
                        //    string dt = stArrayData[i].PadLeft(6, '0');
                        //    stArrayData[i] = dt.Substring(0, 2) + "," + dt.Substring(2, 2) + "," + dt.Substring(4, 2);
                        //}

                        // フィールド結合
                        stBuffer += stArrayData[i];
                    }

                    sRow = 0;
                }
                else
                {
                    sRow++;
                }

                // 読み込んだものを追加で格納する
                stResult += (stBuffer + Environment.NewLine);

                //// 最終行は追加しない（伝票区別記号(*)のため）
                //if (sRow <= global.MAXGYOU_PRN)
                //{
                //    // 読み込んだものを追加で格納する
                //    stResult += (stBuffer + Environment.NewLine);
                //}
            }

            // 後処理
            if (dNum > 0)
            {
                //ファイル書き出し
                outFileWrite(stResult, Properties.Settings.Default.readPath + imgName, newFnm);

                // 入力ファイルを閉じる
                inFile.Close();

                //入力ファイル削除 : "txtout.csv"
                Utility.FileDelete(Properties.Settings.Default.readPath, Properties.Settings.Default.wrReaderOutFile);

                //画像ファイル削除 : "WRH***.tif"
                Utility.FileDelete(Properties.Settings.Default.readPath, "WRH*.tif");
            }
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     分割ファイルを書き出す </summary>
        /// <param name="tempResult">
        ///     書き出す文字列</param>
        /// <param name="tempImgName">
        ///     元画像ファイルパス</param>
        /// <param name="outFileName">
        ///     新ファイル名</param>
        ///----------------------------------------------------------------------------
        private void outFileWrite(string tempResult, string tempImgName, string outFileName)
        {
            //出力ファイル
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(Properties.Settings.Default.dataPath + outFileName + ".csv",
                                                    false, System.Text.Encoding.GetEncoding(932));
            // ファイル書き出し
            outFile.Write(tempResult);

            //ファイルクローズ
            outFile.Close();

            //画像ファイルをコピー
            System.IO.File.Copy(tempImgName, Properties.Settings.Default.dataPath + outFileName + ".tif");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Hide();
            OCR.frmXlsLoad frm = new OCR.frmXlsLoad();
            frm.ShowDialog();
            Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Hide();
            OCR.frmRecovery frm = new OCR.frmRecovery();
            frm.ShowDialog();
            Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Hide();
            sumData.frmAreaByRep frm = new sumData.frmAreaByRep();
            frm.ShowDialog();
            Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Hide();
            sumData.frmPastByMonthRep frm = new sumData.frmPastByMonthRep();
            frm.ShowDialog();
            Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Hide();
            sumData.frmPastByStuffRep frm = new sumData.frmPastByStuffRep(global.flgOff, global.flgOff, global.flgOff, global.flgOff, global.flgOff, string.Empty);
            frm.ShowDialog();
            Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Hide();
            sumData.frmEditLogRep frm = new sumData.frmEditLogRep();
            frm.ShowDialog();
            Show();
        }

    }
}
