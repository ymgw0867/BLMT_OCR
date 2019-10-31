﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Leadtools.Codecs;
using Leadtools;
using Leadtools.ImageProcessing;
using Leadtools.ImageProcessing.Core;
using BLMT_SCAN.common;


namespace BLMT_SCAN
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            timer1.Tick += new EventHandler(timer1_Tick);

            cAdp.Fill(dts.環境設定);
        }

        Timer timer1 = new Timer();
        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.環境設定TableAdapter cAdp = new DataSet1TableAdapters.環境設定TableAdapter();

        private void Form1_Load(object sender, EventArgs e)
        {
            // バルーン表示
            notifyIcon1.ShowBalloonTip(500);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dataGridView1);

            dataGridView1.Rows.Clear();

            // データ受信～OCR認識
            doOCrTimer();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            // データ受信～OCR認識
            logOutput(getTimeNow() + " タイマー開始");
            timer1.Enabled = false;
            doOCrTimer();
            timer1.Enabled = true;
            logOutput(getTimeNow() + " タイマー終了");
        }


        #region カラム定義
        string cDate = "col0";
        string cFrom = "col1";
        string cSubject = "col2";
        string cFileName = "col3";
        string cTlcnt = "col8";
        string cIdxcnt = "col9";
        string cScnt = "col4";
        string cEcnt = "col5";
        string cNGcnt = "col6";
        string cNonOcrcnt = "col7";
        string cMessageID = "col10";
        string cMemo = "col11";
        #endregion

        ///-------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-------------------------------------------------------------------
        private void GridViewSetting(DataGridView g)
        {
            try
            {
                g.EnableHeadersVisualStyles = false;
                g.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                g.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                g.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                g.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                g.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 10, FontStyle.Regular);

                // 行の高さ
                g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                g.ColumnHeadersHeight = 22;
                g.RowTemplate.Height = 22;

                // 全体の高さ
                g.Height = 507;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                g.Columns.Add(cDate, "受信日時");
                g.Columns.Add(cTlcnt, "OCR認識件数");
                g.Columns.Add(cEcnt, "エラー件数");
                g.Columns.Add(cMemo, "備考");


                g.Columns[cDate].Width = 160;
                g.Columns[cTlcnt].Width = 120;
                g.Columns[cEcnt].Width = 120;
                g.Columns[cMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                g.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                g.Columns[cTlcnt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                g.Columns[cEcnt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 行ヘッダを表示しない
                g.RowHeadersVisible = false;

                // 選択モード
                g.SelectionMode = DataGridViewSelectionMode.CellSelect;
                g.MultiSelect = false;

                // 編集不可とする
                g.ReadOnly = true;

                // 追加行表示しない
                g.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                g.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                g.AllowUserToOrderColumns = false;

                // 列サイズ変更可
                g.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                g.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                g.StandardTab = true;

                // 罫線
                g.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                g.CellBorderStyle = DataGridViewCellBorderStyle.None;
                //g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     タイマー処理 </summary>
        ///------------------------------------------------------------
        private void doOCrTimer()
        {
            notifyIcon1.Visible = false;

            // ファイル名（日付時間部分）
            string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                    string.Format("{0:00}", DateTime.Today.Month) +
                    string.Format("{0:00}", DateTime.Today.Day) +
                    string.Format("{0:00}", DateTime.Now.Hour) +
                    string.Format("{0:00}", DateTime.Now.Minute) +
                    string.Format("{0:00}", DateTime.Now.Second);

            int dNum = 0;                       // ファイル名末尾連番

            /* 前回分の処理残ファイルをOCR認識結果ＣＳＶデータを出勤簿ごとに分割して
             * 画像ファイルと共にDATAフォルダへ移動する 
             * */
            //LoadCsvDivide(fName, ref dNum);

            /* マルチTiff画像をシングルtifに分解後に解像度200×200、A4サイズ変換
             * (SCANフォルダ → TRAYフォルダ) */
            if (MultiTif(Properties.Settings.Default.scanPath, Properties.Settings.Default.trayPath, fName))
            {
                ////// 帳票ライブラリV8.0.3によるOCR認識実行
                ////string jobname = Properties.Settings.Default.wrHands_Job;
                ////wrhs803LibOCR(jobname);

                // WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する
                WinReaderOCR();

                // スリープ
                logOutput(getTimeNow() + " スリープ開始");
                System.Threading.Thread.Sleep(60000);
                logOutput(getTimeNow() + " スリープ終了");

                /* OCR認識結果ＣＳＶデータを出勤簿ごとに分割して
                 * 画像ファイルと共にDATAフォルダへ移動する */ 
                LoadCsvDivide(fName, ref dNum);
            }

            notifyIcon1.Visible = true;
        }

        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;

            // 終了する
            Application.Exit();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //インターバルセット
            if (dts.環境設定.Any(a => a.ID == global.configKEY))
            {
                var s = dts.環境設定.Single(a => a.ID == global.configKEY);
                timer1.Interval = s.受信監視スパン * 1000;    // 秒単位
                timer1.Enabled = true;
            }
            else
            {
                MessageBox.Show("環境設定データが登録されていません", "BLMT_SCAN", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
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

            // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
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
                    catch// (Exception e)
                    {
                    }                       //ノイズ除去

                    try
                    {
                        Dkcommand.Run(leadImg);
                    }
                    catch// (Exception e)
                    {
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
                    catch// (Exception e)
                    {
                    }

                    //画像補正処理　終了↑ ****************************

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.Tif, 0, i, i, 1, CodecsSavePageMode.Insert);
                }
            }

            // InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            return true;
        }
        
        ///----------------------------------------------------------------
        /// <summary>
        ///     帳票認識ライブラリ V8.0.3 による認識処理実行 </summary>
        ///----------------------------------------------------------------
        //private void wrhs803LibOCR(string jobName)
        //{
        //    // ファイル名のタイムスタンプを設定
        //    string fnm = string.Format("{0:0000}", DateTime.Today.Year) +
        //                 string.Format("{0:00}", DateTime.Today.Month) +
        //                 string.Format("{0:00}", DateTime.Today.Day) +
        //                 string.Format("{0:00}", DateTime.Now.Hour) +
        //                 string.Format("{0:00}", DateTime.Now.Minute) +
        //                 string.Format("{0:00}", DateTime.Now.Second);

        //    int sNum = 0;
        //    int ret = 0;

        //    int okCnt = 0;
        //    int errCnt = 0;
        //    string msg = "";

        //    try
        //    {
        //        //// オーナーフォームを無効にする
        //        //this.Enabled = false;

        //        //// プログレスバーを表示する
        //        //frmPrg frmP = new frmPrg();
        //        //frmP.Owner = this;
        //        //frmP.Show();

        //        // 処理する画像数を取得
        //        int t = System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif").Count();

        //        // 順番に認識処理を実行
        //        foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif"))
        //        {
        //            // 画像数カウント
        //            sNum++;

        //            //// プログレス表示
        //            //frmP.Text = "OCR認識中です ... " + sNum.ToString() + "/" + t.ToString();
        //            //frmP.progressValue = sNum * 100 / t;
        //            //frmP.ProgressStep();

        //            // 標準パターンの読み込み
        //            ret = FormRecog.OcrPatternLoad(Properties.Settings.Default.ocrPatternLoadPath);

        //            // パターン読み込みに成功したとき
        //            if (ret > 0)
        //            {
        //                // 帳票認識ライブラリの制御内容を設定
        //                //FormRecog.OcrSetStatus(5, 1);   // 強制終了制御
        //                FormRecog.OcrSetStatus(0, 0);

        //                // 認識結果出力イメージファイル
        //                StringBuilder outimage = new StringBuilder(256);
        //                outimage.Append(Properties.Settings.Default.wrOutPath + System.IO.Path.GetFileName(files));

        //                // 認識結果出力テキストファイル
        //                StringBuilder outtext = new StringBuilder(256);
        //                outtext.Append(Properties.Settings.Default.wrOutPath + System.IO.Path.GetFileNameWithoutExtension(files) + ".csv");

        //                // 認識結果 構造体
        //                FormRecog.FORM_RECOG_DATA dt = new FormRecog.FORM_RECOG_DATA();

        //                // 認識処理を開始
        //                ret = FormRecog.OcrFormRecogStart(jobName, files, outimage, outtext, ref dt, false, false);

        //                // 認識成功のとき
        //                if (ret > 0)
        //                {
        //                    // 認識結果のメモリ解放
        //                    ret = FormRecog.OcrFormStructFree(ref dt);

        //                    // 認識終了
        //                    ret = FormRecog.OcrFormRecogEnd();
                            
        //                    // 出力先フォルダ
        //                    string rPath = Properties.Settings.Default.dataPath;

        //                    // 出力されたイメージファイルとテキストファイルのリネーム処理を行います
        //                    // READフォルダ → DATAフォルダ
        //                    string inCsvFile = Properties.Settings.Default.wrOutPath +
        //                                       Properties.Settings.Default.wrReaderOutFile;
        //                    string newFileName = rPath + fnm + sNum.ToString().PadLeft(3, '0');
        //                    wrhOutFileRename(inCsvFile, newFileName);

        //                    okCnt++;
        //                }
        //                else
        //                {
        //                    //MessageBox.Show("OCR認識開始に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                    msg = "OCR認識開始に失敗しました";

        //                    //エラー画像はerrorフォルダに画像をCOPY
        //                    System.IO.File.Copy(files, Properties.Settings.Default.ocrErrPath + System.IO.Path.GetFileName(files));

        //                    errCnt++;
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("OCR標準パターンの読み込みに失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                msg = "OCR標準パターンの読み込みに失敗しました";

        //                //エラー画像はerrorフォルダに画像をCOPY
        //                System.IO.File.Copy(files, Properties.Settings.Default.ocrErrPath + System.IO.Path.GetFileName(files));
        //                errCnt++;
        //            }
        //        }

        //        //// いったんオーナーをアクティブにする
        //        //this.Activate();

        //        //// 進行状況ダイアログを閉じる
        //        //frmP.Close();

        //        //// オーナーのフォームを有効に戻す
        //        //this.Enabled = true;

        //        //// 終了表示
        //        //MessageBox.Show(sNum.ToString() + "件のOCR認識処理を行いました", "終了", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //        // TRAYフォルダの全てのtifファイルを削除します
        //        foreach (var files in System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif"))
        //        {
        //            System.IO.File.Delete(files);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        throw;
        //    }
        //    finally
        //    {
        //        // ログ表示
        //        logGridView(dataGridView1, okCnt, errCnt, msg);
        //    }
        //}
        
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

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Visible = true;               // フォームの表示
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal; // 最小化をやめる
            }
            //this.notifyIcon1.Visible = false;  // Notifyアイコン非表示
            this.Activate();                   // フォームをアクティブにする
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason != CloseReason.ApplicationExitCall)
            {
                e.Cancel = true; // フォームが閉じるのをキャンセル
                this.Visible = false; // フォームの非表示
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     データグリッドに受信監視ログを表示する </summary>
        /// <param name="md">
        ///     メールデータ</param>
        ///------------------------------------------------------------
        private void logGridView(DataGridView dg, int cCnt, int eCnt, string msg)
        {
            dg.Rows.Add();
            dg[cDate, dg.RowCount - 1].Value = DateTime.Now.ToString();
            dg[cTlcnt, dg.RowCount - 1].Value = cCnt.ToString();
            dg[cEcnt, dg.RowCount - 1].Value = eCnt.ToString();
            dg[cMemo, dg.RowCount - 1].Value = msg;

            dg.CurrentCell = null;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     WinReaderを起動して出勤簿をスキャンしてOCR処理を実施する </summary>
        ///----------------------------------------------------------------------------------
        private void WinReaderOCR()
        {
            //// WinReaderJOB起動文字列
            //string JobName = @"""" + Properties.Settings.Default.wrHands_Job + @"""" + " /H2";
            //string winReader_exe = Properties.Settings.Default.wrHands_Path +
            //    Properties.Settings.Default.wrHands_Prg;

            // ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();

            // 起動するアプリケーションを設定する
            //p.FileName = winReader_exe;
            p.FileName = @"C:\Windows\System32\schtasks.exe"; 

            // コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
            //p.Arguments = JobName;
            p.Arguments = "/run /tn blmt_scan";

            // ログ
            logOutput(getTimeNow() + " task起動");

            // WinReaderを起動します
            System.Diagnostics.Process hProcess = System.Diagnostics.Process.Start(p);
            logOutput(getTimeNow() + " task起動後");
            
            // taskが終了するまで待機する
            hProcess.WaitForExit();
            logOutput(getTimeNow() + " task終了待機後");
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

            logOutput(getTimeNow() + " OCR認識結果処理開始");

            // 対象ファイルの存在を確認します
            if (!System.IO.File.Exists(Properties.Settings.Default.readPath + Properties.Settings.Default.wrReaderOutFile))
            {
                logOutput(getTimeNow() + " OCR認識データなし");
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

                // ログ表示
                logGridView(dataGridView1, dCnt, 0, "");
            }


            logOutput(getTimeNow() + " OCR認識結果処理終了");
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

        private void logOutput(string tempResult)
        {
            //出力ファイル
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(@"C:\BLMT_OCR\ocrLog.csv",true,
                                                    System.Text.Encoding.GetEncoding(932));
            // ファイル書き出し
            outFile.Write(tempResult + Environment.NewLine);

            //ファイルクローズ
            outFile.Close();
        }

        private string getTimeNow()
        {
            DateTime dt = DateTime.Now;

            string str = dt.Year.ToString() + "/" + dt.Month.ToString().PadLeft(2, '0') + "/" + dt.Day.ToString().PadLeft(2, '0') + " " +
                         dt.Hour.ToString().PadLeft(2, '0') + ":" + dt.Minute.ToString().PadLeft(2, '0') + ":" + dt.Second.ToString().PadLeft(2, '0');

            return str; 
        }
    }
}
