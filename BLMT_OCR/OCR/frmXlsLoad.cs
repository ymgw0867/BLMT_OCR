using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;
using Excel = Microsoft.Office.Interop.Excel;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using Leadtools.ImageProcessing.Core;

namespace BLMT_OCR.OCR
{
    public partial class frmXlsLoad : Form
    {
        public frmXlsLoad()
        {
            InitializeComponent();
        }

        private void frmXlsLoad_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // エクセル出勤簿があるか？
            int s = System.IO.Directory.GetFiles(Properties.Settings.Default.exPath, "*.xlsx").Count();
            
            // ボタンEnable調整
            if (s > 0)
            {
                btnData.Enabled = true;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Maximum = s;
            }
            else
            {
                btnData.Enabled = false;
            }
        }

        private void getExcelData()
        {
            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excel起動
            //string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = null;

            Excel.Worksheet oxlsSheet = null;
            Excel.Worksheet oxlsMsSheet = null; // テンプレートシート

            Excel.Range rng = null;
            Excel.Range[] aRng = new Excel.Range[3];

            object[,] rtnArray = null;

            try
            {
                // ファイル名（日付時間部分）
                string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                        string.Format("{0:00}", DateTime.Today.Month) +
                        string.Format("{0:00}", DateTime.Today.Day) +
                        string.Format("{0:00}", DateTime.Now.Hour) +
                        string.Format("{0:00}", DateTime.Now.Minute) +
                        string.Format("{0:00}", DateTime.Now.Second);

                // ファイル名末尾連番
                int dNum = 0;
                int ngCnt = 0;

                // エクセルファイルを順次読み込む
                foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.exPath, "*.xlsx"))
                {
                    oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(file, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
                    oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
                    oxlsSheet.Select(Type.Missing);

                    dNum++;
                    toolStripProgressBar1.Value = dNum;

                    // シートのセルを一括して2次元配列に取得します
                    rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[1, 1], oxlsMsSheet.Cells[60, 204]];
                    rtnArray = (object[,])rng.Value2;

                    // 識別記号
                    string mk = Utility.NulltoStr(rtnArray[1, 1]);
                    if (mk != Properties.Settings.Default.exMark && mk != Properties.Settings.Default.exMark2)
                    {
                        ngCnt++;
                        // エクセル出勤簿以外のときネグる
                        continue;
                    }

                    StringBuilder sb = new StringBuilder();

                    // 2次元配列のデータを読み取る
                    sb.Append("*").Append(",");
                    sb.Append(fName + dNum.ToString("D3") + ".tif").Append(",");          // 画像名
                    sb.Append(Utility.NulltoStr(rtnArray[2, 14])).Append(",");   // 年
                    sb.Append(Utility.NulltoStr(rtnArray[2, 32])).Append(",");  // 月
                    sb.Append(Utility.NulltoStr(rtnArray[4, 105])).Append(",");  // スタッフ№
                    sb.Append(Utility.NulltoStr(rtnArray[9, 162])).Append(",");  // 基本実労働時間・時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 176])).Append(",");  // 基本実労働時間・分
                    sb.Append(Utility.NulltoStr(rtnArray[7, 173])).Append(",");  // 店舗コード
                    sb.Append(Utility.NulltoStr(rtnArray[9, 29])).Append(",");  // 基本就労時間帯１・開始時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 43])).Append(",");  // 基本就労時間帯１・開始分
                    sb.Append(Utility.NulltoStr(rtnArray[9, 57])).Append(",");  // 基本就労時間帯１・終了時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 71])).Append(",");  // 基本就労時間帯１・終了分
                    sb.Append(Utility.NulltoStr(rtnArray[9, 89])).Append(",");  // 基本就労時間帯２・開始時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 103])).Append(",");  // 基本就労時間帯２・開始分
                    sb.Append(Utility.NulltoStr(rtnArray[9, 117])).Append(",");  // 基本就労時間帯２・終了時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 131])).Append(",");  // 基本就労時間帯２・終了分
                    sb.Append(Utility.NulltoStr(rtnArray[9, 149])).Append(",");  // 基本就労時間帯３・開始時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 163])).Append(",");  // 基本就労時間帯３・開始分
                    sb.Append(Utility.NulltoStr(rtnArray[9, 177])).Append(",");  // 基本就労時間帯３・終了時
                    sb.Append(Utility.NulltoStr(rtnArray[9, 191])).Append(",");  // 基本就労時間帯３・終了分
                    sb.Append(Utility.NulltoStr(rtnArray[48, 113])).Append(",");     // 実労日数
                    sb.Append(Utility.NulltoStr(rtnArray[48, 137])).Append(",");     // 公休日数
                    sb.Append(Utility.NulltoStr(rtnArray[48, 161])).Append(",");     // 有休日数
                    sb.Append(Utility.NulltoStr(rtnArray[48, 185])).Append(",");    // 合計日数

                    // 訂正欄
                    for (int i = 15; i <= 195; i += 6)
                    {
                        sb.Append(Utility.getTeiseiCheck(Utility.NulltoStr(rtnArray[53, i])));

                        if (i < 195)
                        {
                            sb.Append(",");
                        }
                        else
                        {
                            sb.Append(Environment.NewLine);
                        }
                    }
                    
                    // 出勤簿・前半
                    for (int i = 17; i < 49; i += 2)
                    {
                        sb.Append(Utility.NulltoStr(rtnArray[i, 19])).Append(",");  // 出勤状況
                        sb.Append(Utility.NulltoStr(rtnArray[i, 29])).Append(",");  // 出勤時刻・時
                        sb.Append(Utility.NulltoStr(rtnArray[i, 43])).Append(",");  // 出勤時刻・分
                        sb.Append(Utility.NulltoStr(rtnArray[i, 53])).Append(",");  // 退勤時刻・時
                        sb.Append(Utility.NulltoStr(rtnArray[i, 67])).Append(",");  // 退勤時刻・分
                        sb.Append(Utility.NulltoStr(rtnArray[i, 77])).Append(Environment.NewLine);  // 休憩
                    }

                    // 出勤簿・後半
                    for (int i = 17; i < 47; i += 2)
                    {
                        sb.Append(Utility.NulltoStr(rtnArray[i, 117])).Append(",");  // 出勤状況
                        sb.Append(Utility.NulltoStr(rtnArray[i, 127])).Append(",");  // 出勤時刻・時
                        sb.Append(Utility.NulltoStr(rtnArray[i, 141])).Append(",");  // 出勤時刻・分
                        sb.Append(Utility.NulltoStr(rtnArray[i, 151])).Append(",");  // 退勤時刻・時
                        sb.Append(Utility.NulltoStr(rtnArray[i, 165])).Append(",");  // 退勤時刻・分
                        sb.Append(Utility.NulltoStr(rtnArray[i, 175])).Append(Environment.NewLine);  // 休憩
                    }
                    
                    // 出力ファイル
                    string outPath = Properties.Settings.Default.dataPath + fName + dNum.ToString("D3") + ".csv"; 
                    System.IO.StreamWriter outFile = new System.IO.StreamWriter(outPath, false, System.Text.Encoding.GetEncoding(932));

                    // ファイル書き出し
                    outFile.Write(sb.ToString());

                    // コピー
                    rng.Copy(Type.Missing);

                    //クリップボードからデータ取得
                    Image iData = Clipboard.GetImage();
                    iData.Save(Properties.Settings.Default.dataPath + fName + dNum.ToString("D3") + ".tif", System.Drawing.Imaging.ImageFormat.Tiff);
                    iData.Dispose();

                    imgResize(Properties.Settings.Default.dataPath + fName + dNum.ToString("D3") + ".tif");

                    // ファイルクローズ
                    outFile.Close();
                }

                // 終了処理
                string msg = "エクセル出勤簿：" + (dNum - ngCnt) + "件" + Environment.NewLine;
                msg += "出勤簿以外のエクセルファイル：" + ngCnt + "件";

                MessageBox.Show(msg, "処理終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void btnData_Click(object sender, EventArgs e)
        {
            getExcelData();

            moveXlsToBox();
            
            Close();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            // 閉じる
            Close();
        }

        private void frmXlsLoad_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void imgResize(string files)
        {
            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();

            int _pageCount = 0;
            string fnm = string.Empty;

            //コマンドを準備します。(傾き・ノイズ除去・リサイズ)
            DeskewCommand Dcommand = new DeskewCommand();
            DespeckleCommand Dkcommand = new DespeckleCommand();
            SizeCommand Rcommand = new SizeCommand();

            // 画像読み出す
            RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

            // 頁数を取得
            int _fd_count = leadImg.PageCount;

            //// ファイル名設定
            //_pageCount++;
            //fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";
            fnm = files;

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
            //cs.Save(leadImg, fnm, RasterImageFormat.Tif, 1, 1, 1, 1, CodecsSavePageMode.Insert);
            cs.Save(leadImg, fnm, RasterImageFormat.Ccitt, 1, 1, 1, 1, CodecsSavePageMode.Insert);  

            //LEADTOOLS入出力ライブラリを終了します。
            RasterCodecs.Shutdown();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     処理済みExcelファイルをボックスへ移動 </summary>
        ///----------------------------------------------------------------
        private void moveXlsToBox()
        {
            // ボックスフォルダがなければ作成する
            if (!System.IO.Directory.Exists(Properties.Settings.Default.exRePath))
            {
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.exRePath);
            }
            
            // 処理済みExcelファイルをボックスへ移動
            foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.exPath, "*.xlsx"))
            {
                // 移動先ファイル名
                string nfm = Properties.Settings.Default.exRePath + System.IO.Path.GetFileName(file);

                // 同名ファイルがあるときは削除する
                if (System.IO.File.Exists(nfm))
                {
                    System.IO.File.Delete(nfm);
                }

                // 移動する
                System.IO.File.Move(file, nfm);
            }                
        }
    }
}
