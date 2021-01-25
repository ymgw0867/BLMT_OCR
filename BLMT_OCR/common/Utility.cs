using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace BLMT_OCR.common
{
    class Utility
    {
        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     ログインアカウントコンボボックスクラス </summary>
        ///--------------------------------------------------------
        public class comboAccount
        {
            public int code { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     ログインアカウントコンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox, DataSet1TableAdapters.ログインユーザーTableAdapter uAdp)
            {
                DataSet1.ログインユーザーDataTable dt = uAdp.GetData();

                try
                {
                    comboAccount cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "code";

                    foreach (var a in dt.OrderBy(a => a.ID))
                    {
                        cmb1 = new comboAccount();
                        cmb1.code = a.ID;
                        cmb1.Name = a.名前;
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "アカウントコンボボックスロード");
                }
            }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     ログインアカウントコンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, int code)
            {
                comboAccount cmbS = new comboAccount();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboAccount)tempBox.SelectedItem;

                    if (cmbS.code == code)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     休日コンボボックスクラス </summary>
        ///--------------------------------------------------------
        public class comboHoliday
        {
            public string Date { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox)
            {

                // 休日配列
                string[] sDay = {"01/01元旦", "     成人の日", "02/11建国記念の日", "     春分の日", "04/29昭和の日",
                            "05/03憲法記念日","05/04みどりの日","05/05こどもの日","08/11山の日","     海の日","     敬老の日",
                            "     秋分の日","     体育の日","11/03文化の日","11/23勤労感謝の日","12/23天皇誕生日",
                            "     振替休日","     国民の休日","     土曜日","     年末年始休暇","     夏季休暇"};

                try
                {
                    comboHoliday cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Date";

                    foreach (var a in sDay)
                    {
                        cmb1 = new comboHoliday();
                        cmb1.Date = a.Substring(0, 5);
                        int s = a.Length;
                        cmb1.Name = a.Substring(5, s - 5);
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "休日コンボボックスロード");
                }
            }
            
            ///------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, string dt)
            {
                comboHoliday cmbS = new comboHoliday();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboHoliday)tempBox.SelectedItem;

                    if (cmbS.Date == dt)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        ///--------------------------------------------------
        /// <summary>
        ///     エリアコンボボックスクラス </summary>
        ///--------------------------------------------------
        public class comboMana
        {
            public string Code { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     エリアコンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            /// <param name="shp">
            ///     エリアマスター配列</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox, clsMana[] ara)
            {
                try
                {
                    comboMana cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Code";

                    foreach (var a in ara.OrderBy(a => a.エリアコード))
                    {
                        cmb1 = new comboMana();
                        cmb1.Code = a.エリアマネージャー名;
                        cmb1.Name = a.エリアマネージャー名;
                        tempBox.Items.Add(cmb1);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "エリアマネージャーコンボボックスロード");
                }
            }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="sCode">
            ///     エリアマネージャー</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, string sCode)
            {
                comboMana cmbS = new comboMana();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboMana)tempBox.SelectedItem;

                    if (cmbS.Code == sCode)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        ///--------------------------------------------------
        /// <summary>
        ///     エリアコンボボックスクラス </summary>
        ///--------------------------------------------------
        public class comboArea
        {
            public int Code { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     エリアコンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            /// <param name="shp">
            ///     エリアマスター配列</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox, clsArea[] ara)
            {
                try
                {
                    comboArea cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Code";

                    foreach (var a in ara.OrderBy(a => a.エリアコード))
                    {
                        cmb1 = new comboArea();
                        cmb1.Code = a.エリアコード;
                        cmb1.Name = a.エリアコード.ToString("D3") + ":" + a.エリア名;
                        tempBox.Items.Add(cmb1);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "エリアマスターコンボボックスロード");
                }
            }
            
            ///------------------------------------------------------------------------
            /// <summary>
            ///     コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="sCode">
            ///     エリアコード</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, int sCode)
            {
                comboArea cmbS = new comboArea();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboArea)tempBox.SelectedItem;

                    if (cmbS.Code == sCode)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }


        ///--------------------------------------------------
        /// <summary>
        ///     店舗コンボボックスクラス </summary>
        ///--------------------------------------------------
        public class comboShop
        {
            public int Code { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     店舗コンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            /// <param name="shp">
            ///     店舗マスター配列</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox, clsShop [] shp)
            {
                try
                {
                    comboShop cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Code";

                    foreach (var a in shp.OrderBy(a => a.エリアコード).ThenBy(a => a.店舗コード))
                    {
                        cmb1 = new comboShop();
                        cmb1.Code = a.店舗コード;
                        cmb1.Name = a.エリアコード.ToString("D3") + ":" + a.エリア名 + " " + a.店舗コード.ToString("D5") + ":" + a.店舗名;
                        tempBox.Items.Add(cmb1);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "店舗マスターコンボボックスロード");
                }
            }


            ///------------------------------------------------------------------------
            /// <summary>
            ///     コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, int sCode)
            {
                comboShop cmbS = new comboShop();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboShop)tempBox.SelectedItem;

                    if (cmbS.Code == sCode)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
            ///------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        ///------------------------------------------------------------------------
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// <summary>
        /// 文字型をIntへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>Int型の値</returns>
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// <summary>
        /// 文字型をDoubleへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>double型の値</returns>
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        ///-----------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てします。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     ファイル選択ダイアログボックスの表示 </summary>
        /// <param name="sTitle">
        ///     タイトル文字列</param>
        /// <param name="sFilter">
        ///     ファイルのフィルター</param>
        /// <returns>
        ///     選択したファイル名</returns>
        ///------------------------------------------------------------------
        public static string userFileSelect(string sTitle, string sFilter)
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = sTitle;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = sFilter;
            //openFileDialog1.Filter = "CSVファイル(*.CSV)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        public class frmMode
        {
            public int ID { get; set; }

            public int Mode { get; set; }

            public int rowIndex { get; set; }
        }

        public class xlsShain
        {
            public int sCode { get; set; }
            public string sName { get; set; }
            public int bCode { get; set; }
            public string bName { get; set; }
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを追加モードで出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            //// ファイル名（タイムスタンプ）
            //string timeStamp = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            //// ファイル名
            //string outFileName = sPath + timeStamp + ".csv";

            //// 出力ファイルが存在するとき
            //if (System.IO.File.Exists(outFileName))
            //{
            //    // 既存のファイルを削除
            //    System.IO.File.Delete(outFileName);
            //}

            // CSVファイル出力
            //System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("shift-jis"));
            System.IO.File.AppendAllLines(sPath, arrayData, Encoding.GetEncoding("shift-Jis"));
        }
        
        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;

            // 文字間のスペースを除去 2015/03/10
            s = s.Replace(" ", "");

            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     訂正欄にチェックされているとき「１」されていないとき「０」を返す</summary>
        /// <param name="s">
        ///     文字列</param>
        /// --------------------------------------------------------------------
        public static string getTeiseiCheck(string s)
        {
            string val = string.Empty;

            if (s != string.Empty)
            {
                val = global.FLGON;
            }
            else
            {
                val = global.FLGOFF;
            }

            return val;
        }
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     分記入範囲チェック：0～59の数値及び記入単位 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <param name="tani">
        ///     記入単位分</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        public static bool checkMinSpan(string m, int tani)
        {
            if (!Utility.NumericCheck(m)) return false;
            else if (int.Parse(m) < 0 || int.Parse(m) > 59) return false;
            else if (int.Parse(m) % tani != 0) return false;
            else return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入範囲チェック 0～23の数値 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        public static bool checkHourSpan(string h)
        {
            if (!Utility.NumericCheck(h)) return false;
            else if (int.Parse(h) < 0 || int.Parse(h) > 23) return false;
            else return true;
        }


        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     エリア別社員マスター（Excel)シート情報配列に取り込む </summary>
        ///-----------------------------------------------------------------------------
        public bool getXlsToArray(string xlsPath, clsXlsmst[] xlsArray, ref System.Collections.ArrayList al)
        {
            int iR = 0;
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = null;

            try
            {
                for (int i = 1; i <= oXlsBook.Sheets.Count; i++)
                {
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[i];
                    oxlsSheet.Select(Type.Missing);

                    Excel.Range rng = null;

                    Array.Resize(ref xlsArray, i);
                    xlsArray[i - 1] = new clsXlsmst();

                    rng = oxlsSheet.Range[oxlsSheet.Cells[4, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                    xlsArray[i - 1].xlsMst = (object[,])rng.Value2;
                    xlsArray[i - 1].xlsMstID = i;

                    //cmbArea.Items.Add(oXlsBook.Sheets[i].Name);
                }

                //cmbArea.SelectedIndex = 0;

                // 2次元配列から1次元配列に取り込み
                al = array2ToArrayList(xlsArray);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "エクセル社員マスター読み取りエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     エリア別社員マスター（Excel)シート情報をスタッフ配列クラスに取り込む </summary>
        ///-----------------------------------------------------------------------------
        public bool getXlsToArray(string xlsPath, clsXlsmst[] xlsArray, ref clsStaff [] stf)
        {
            int iR = 0;
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = null;

            try
            {
                for (int i = 1; i <= oXlsBook.Sheets.Count; i++)
                {
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[i];
                    oxlsSheet.Select(Type.Missing);

                    Excel.Range rng = null;

                    Array.Resize(ref xlsArray, i);
                    xlsArray[i - 1] = new clsXlsmst();

                    rng = oxlsSheet.Range[oxlsSheet.Cells[4, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, oxlsSheet.UsedRange.Columns.Count]];
                    xlsArray[i - 1].xlsMst = (object[,])rng.Value2;
                    xlsArray[i - 1].xlsMstID = i;

                    //cmbArea.Items.Add(oXlsBook.Sheets[i].Name);
                }

                //cmbArea.SelectedIndex = 0;

                // 2次元配列からスタッフ配列クラス作成
                array2ToArrayList(xlsArray, ref stf);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "エクセル社員マスター読み取りエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     2次元配列からArrayListにデータをセットする </summary>
        ///----------------------------------------------------------------------
        private System.Collections.ArrayList array2ToArrayList(clsXlsmst[] xlsArray)
        {
            System.Collections.ArrayList al = new System.Collections.ArrayList();

            for (int i = 0; i < xlsArray.Length; i++)
            {
                for (int iX = 1; iX <= xlsArray[i].xlsMst.GetLength(0); iX++)
                {
                    // 26列ないときはエリア別社員マスターシートとみなさない
                    if (xlsArray[i].xlsMst.GetLength(1) < 26)
                    {
                        continue;
                    }

                    // 使用区分が１以外はネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 26]).Trim() != global.FLGON)
                    {
                        continue;
                    }

                    // エリア№が空白のときネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).Trim() == string.Empty)
                    {
                        continue;
                    }

                    StringBuilder sb = new StringBuilder();
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).PadLeft(3, '0')).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 4]).PadLeft(5, '0')).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 20]).PadLeft(5, '0')).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 2])).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 5])).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 21])).Append(",");
                    sb.Append(getTimeString(xlsArray, i, iX, 8)).Append(",");
                    sb.Append(getTimeString(xlsArray, i, iX, 12)).Append(",");
                    sb.Append(getTimeString(xlsArray, i, iX, 16)).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 23])).Append(",");
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 22])).Append(",");

                    // 実労働時間
                    string wH = string.Empty;
                    string wM = string.Empty;
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]) != string.Empty)
                    {
                        wH = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]).PadLeft(2, ' ');
                    }

                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]) != string.Empty)
                    {
                        wM = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]).PadLeft(2, '0');
                    }

                    if (wH != string.Empty || wM != string.Empty)
                    {
                        sb.Append(wH).Append(":").Append(wM).Append(",");
                    }
                    else
                    {
                        sb.Append(string.Empty).Append(",");
                    }

                    // エリアマネージャー名
                    sb.Append(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 3]));
                    
                    //Array.Resize(ref csvArray, iR + 1);
                    //csvArray[iR] = sb.ToString();

                    al.Add(sb.ToString());

                    //iR++;
                }
            }

            al.Sort();

            return al;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     2次元配列から配列スタッフクラスにデータをセットする </summary>
        ///----------------------------------------------------------------------
        private void array2ToArrayList(clsXlsmst[] xlsArray, ref clsStaff [] stf)
        {
            int r = 0;

            for (int i = 0; i < xlsArray.Length; i++)
            {
                for (int iX = 1; iX <= xlsArray[i].xlsMst.GetLength(0); iX++)
                {
                    // 26列ないときはエリア別社員マスターシートとみなさない
                    if (xlsArray[i].xlsMst.GetLength(1) < 26)
                    {
                        continue;
                    }

                    // 使用区分が１以外はネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 26]).Trim() != global.FLGON)
                    {
                        continue;
                    }

                    // エリア№が空白のときネグる
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).Trim() == string.Empty)
                    {
                        continue;
                    }

                    Array.Resize(ref stf, r + 1);
                                            
                    stf[r] = new clsStaff();
                    stf[r].エリアコード = Utility.StrtoInt(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 1]).PadLeft(3, '0'));
                    stf[r].エリア名 = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 2]);
                    stf[r].店舗コード = Utility.StrtoInt(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 4]));
                    stf[r].店舗名 = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 5]);
                    stf[r].スタッフコード = Utility.StrtoInt(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 20]));
                    stf[r].スタッフ名 = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 21]);
                    stf[r].基本時間帯1 = getTimeString(xlsArray, i, iX, 8);
                    stf[r].基本時間帯2 = getTimeString(xlsArray, i, iX, 12);
                    stf[r].基本時間帯3 = getTimeString(xlsArray, i, iX, 16);
                    stf[r].給与区分 = Utility.StrtoInt(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 22]));
                    stf[r].締日区分 = Utility.StrtoInt(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 23]));
                    stf[r].店舗区分 = Utility.StrtoInt(Utility.NulltoStr(xlsArray[i].xlsMst[iX, 6]));
                    
                    // 実労働時間
                    string wH = string.Empty;
                    string wM = string.Empty;
                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]) != string.Empty)
                    {
                        wH = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 24]).PadLeft(2, ' ');
                    }

                    if (Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]) != string.Empty)
                    {
                        wM = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 25]).PadLeft(2, '0');
                    }

                    if (wH != string.Empty || wM != string.Empty)
                    {
                        stf[r].実労時間 = wH + ":" + wM;
                    }
                    else
                    {
                        stf[r].実労時間 = string.Empty;
                    }

                    stf[r].エリアマネージャー名 = Utility.NulltoStr(xlsArray[i].xlsMst[iX, 3]);

                    r++;
                }
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     基本就業時間帯文字列取得 </summary>
        /// <param name="x">
        ///     Excelシート配列クラス</param>
        /// <param name="i">
        ///     Excelシート配列クラスインデックス</param>
        /// <param name="iX">
        ///     シート配列１次元インデックス</param>
        /// <param name="r">
        ///     シート配列２次元インデックス</param>
        /// <returns>
        ///     基本就業時間帯文字列</returns>
        ///-----------------------------------------------------------
        private string getTimeString(clsXlsmst[] x, int i, int iX, int r)
        {
            string rtn = string.Empty;

            string sHH = Utility.NulltoStr(x[i].xlsMst[iX, r]);
            string sMM = Utility.NulltoStr(x[i].xlsMst[iX, r + 1]);
            string eHH = Utility.NulltoStr(x[i].xlsMst[iX, r + 2]);
            string eMM = Utility.NulltoStr(x[i].xlsMst[iX, r + 3]);

            if (sHH != string.Empty || sMM != string.Empty || eHH != string.Empty || eMM != string.Empty)
            {
                rtn = sHH.PadLeft(2, ' ') + ":" + sMM.PadLeft(2, '0') + "～" + eHH.PadLeft(2, ' ') + ":" + eMM.PadLeft(2, '0');
            }

            return rtn;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     エリアマネージャー配列クラス作成  </summary>
        /// <param name="stf">
        ///     スタッフ配列</param>
        /// <param name="shp">
        ///     エリアマスター配列</param>
        ///-------------------------------------------------------------
        public void getManaMaster(clsStaff[] stf, ref clsMana[] ara)
        {
            var s = stf.OrderBy(a => a.エリアマネージャー名).Select(a => a.エリアマネージャー名).Distinct();

            int i = 0;

            foreach (var t in s)
            {
                Array.Resize(ref ara, i + 1);

                ara[i] = new clsMana();
                ara[i].エリアコード = 0;
                ara[i].エリア名 = string.Empty;
                ara[i].エリアマネージャー名 = t;

                for (int iX = 0; iX < stf.Length; iX++)
                {
                    if (stf[iX].エリアマネージャー名 == t)
                    {
                        ara[i].エリアコード = stf[iX].エリアコード;
                        ara[i].エリア名 = stf[iX].エリア名;
                        break;
                    }
                }

                i++;
            }
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     エリアマスター配列クラス作成  </summary>
        /// <param name="stf">
        ///     スタッフ配列</param>
        /// <param name="shp">
        ///     エリアマスター配列</param>
        ///-------------------------------------------------------------
        public void getAreaMaster(clsStaff[] stf, ref clsArea[] ara)
        {
            var s = stf.OrderBy(a => a.エリアコード).Select(a => a.エリアコード).Distinct();

            int i = 0;

            foreach (var t in s)
            {
                Array.Resize(ref ara, i + 1);

                ara[i] = new clsArea();
                ara[i].エリアコード = t;
                ara[i].エリア名 = string.Empty;
                ara[i].エリアマネージャー名 = string.Empty;

                for (int iX = 0; iX < stf.Length; iX++)
                {
                    if (stf[iX].エリアコード == t)
                    {
                        ara[i].エリア名 = stf[iX].エリア名;
                        ara[i].エリアマネージャー名 = stf[iX].エリアマネージャー名;
                        break;
                    }
                }

                i++;
            }
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     店舗マスター配列クラス作成  </summary>
        /// <param name="stf">
        ///     スタッフ配列</param>
        /// <param name="shp">
        ///     店舗マスター配列</param>
        ///-------------------------------------------------------------
        public void getShopMaster(clsStaff[] stf, ref clsShop[] shp)
        {
            var s = stf.OrderBy(a => a.店舗コード).Select(a => a.店舗コード).Distinct();

            int i = 0;

            foreach (var t in s)
            {
                Array.Resize(ref shp, i + 1);

                shp[i] = new clsShop();
                shp[i].店舗コード = t;
                shp[i].エリアコード = global.flgOff;
                shp[i].エリア名 = string.Empty;
                shp[i].エリアマネージャー名 = string.Empty;
                shp[i].店舗名 = string.Empty;

                for (int iX = 0; iX < stf.Length; iX++)
                {
                    if (stf[iX].店舗コード == t)
                    {
                        shp[i].エリアコード = stf[iX].エリアコード;
                        shp[i].エリア名 = stf[iX].エリア名;
                        shp[i].エリアマネージャー名 = stf[iX].エリアマネージャー名;
                        shp[i].店舗名 = stf[iX].店舗名;
                        break;
                    }
                }

                i++;
            }
        }
        
        public static OleDbConnection dbConnect()
        {
            // データベース接続文字列
            OleDbConnection Cn = new OleDbConnection();
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            sb.Append(Properties.Settings.Default.mdbPath + global.MDBFILE);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            return Cn;
        }
    }
}
