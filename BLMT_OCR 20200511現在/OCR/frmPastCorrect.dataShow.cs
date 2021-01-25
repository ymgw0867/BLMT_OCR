using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Data.OleDb;
using BLMT_OCR.common;
using GrapeCity.Win.MultiRow;

namespace BLMT_OCR.OCR
{
    partial class frmPastCorrect
    {
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
        #endregion

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダと勤務票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getDataSet()
        {
            pAdpMn.過去勤務票ヘッダTableAdapter.Fill(dts.過去勤務票ヘッダ);
            pAdpMn.過去勤務票明細TableAdapter.Fill(dts.過去勤務票明細);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(string iX)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            WorkTotalSumStatus = false;

            // 過去勤務票ヘッダテーブル行を取得
            DataSet1.過去勤務票ヘッダRow r = dts.過去勤務票ヘッダ.Single(a => a.ID == iX);

            // フォーム初期化
            formInitialize(dID, iX);

            // ヘッダ情報表示
            gcMultiRow2[0, "txtYear"].Value = r.年.ToString(); ;
            gcMultiRow2[0, "txtMonth"].Value = Utility.EmptytoZero(r.月.ToString());
            gcMultiRow2[0, "txtSNum"].Value = r.スタッフコード.ToString();

            gcMultiRow2.SetValue(0, "lblName", r.氏名);
            gcMultiRow2.SetValue(0, "lblShopCode", r.店舗コード);
            gcMultiRow2.SetValue(0, "lblShopName", r.店舗名);

            if (r.Is基本実労働時Null())
            {
                gcMultiRow2[0, "txtWh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2.SetValue(0, "txtWh", r.基本実労働時);
            }

            if (r.Is基本実労働分Null())
            {
                gcMultiRow2[0, "txtWm"].Value = string.Empty;
            }
            else
            {
                if (r.基本実労働分 != string.Empty)
                {
                    gcMultiRow2[0, "txtWm"].Value = r.基本実労働分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtWm"].Value = string.Empty;
                }
            }

            // 基本就業時間帯1
            if (r.Is基本就業時間帯1開始時Null())
            {
                gcMultiRow2[0, "txtTm1Sh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2[0, "txtTm1Sh"].Value = r.基本就業時間帯1開始時;
            }

            if (r.Is基本就業時間帯1開始分Null())
            {
                gcMultiRow2[0, "txtTm1Sm"].Value = string.Empty;
            }
            else
            {
                if (r.基本就業時間帯1開始分 != string.Empty)
                {
                    gcMultiRow2[0, "txtTm1Sm"].Value = r.基本就業時間帯1開始分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtTm1Sm"].Value = string.Empty;
                }
            }

            if (r.Is基本就業時間帯1終了時Null())
            {
                gcMultiRow2[0, "txtTm1Eh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2[0, "txtTm1Eh"].Value = r.基本就業時間帯1終了時;
            }

            if (r.Is基本就業時間帯1終了分Null())
            {
                gcMultiRow2[0, "txtTm1Em"].Value = string.Empty;
            }
            else
            {
                if (r.基本就業時間帯1終了分 != string.Empty)
                {
                    gcMultiRow2[0, "txtTm1Em"].Value = r.基本就業時間帯1終了分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtTm1Em"].Value = string.Empty;
                }
            }

            // 基本就業時間帯2
            if (r.Is基本就業時間帯2開始時Null())
            {
                gcMultiRow2[0, "txtTm2Sh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2[0, "txtTm2Sh"].Value = r.基本就業時間帯2開始時;
            }

            if (r.Is基本就業時間帯2開始分Null())
            {
                gcMultiRow2[0, "txtTm2Sm"].Value = string.Empty;
            }
            else
            {
                if (r.基本就業時間帯2開始分 != string.Empty)
                {
                    gcMultiRow2[0, "txtTm2Sm"].Value = r.基本就業時間帯2開始分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtTm2Sm"].Value = string.Empty;
                }
            }

            if (r.Is基本就業時間帯2終了時Null())
            {
                gcMultiRow2[0, "txtTm2Eh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2[0, "txtTm2Eh"].Value = r.基本就業時間帯2終了時;
            }

            if (r.Is基本就業時間帯2終了分Null())
            {
                gcMultiRow2[0, "txtTm2Em"].Value = string.Empty;
            }
            else
            {
                if (r.基本就業時間帯2終了分 != string.Empty)
                {
                    gcMultiRow2[0, "txtTm2Em"].Value = r.基本就業時間帯2終了分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtTm2Em"].Value = string.Empty;
                }
            }

            // 基本就業時間帯3
            if (r.Is基本就業時間帯3開始時Null())
            {
                gcMultiRow2[0, "txtTm3Sh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2[0, "txtTm3Sh"].Value = r.基本就業時間帯3開始時;
            }

            if (r.Is基本就業時間帯3開始分Null())
            {
                gcMultiRow2[0, "txtTm3Sm"].Value = string.Empty;
            }
            else
            {
                if (r.基本就業時間帯3開始分 != string.Empty)
                {
                    gcMultiRow2[0, "txtTm3Sm"].Value = r.基本就業時間帯3開始分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtTm3Sm"].Value = string.Empty;
                }
            }

            if (r.Is基本就業時間帯3終了時Null())
            {
                gcMultiRow2[0, "txtTm3Eh"].Value = string.Empty;
            }
            else
            {
                gcMultiRow2[0, "txtTm3Eh"].Value = r.基本就業時間帯3終了時;
            }

            if (r.Is基本就業時間帯3終了分Null())
            {
                gcMultiRow2[0, "txtTm3Em"].Value = string.Empty;
            }
            else
            {
                if (r.基本就業時間帯3終了分 != string.Empty)
                {
                    gcMultiRow2[0, "txtTm3Em"].Value = r.基本就業時間帯3終了分.PadLeft(2, '0');
                }
                else
                {
                    gcMultiRow2[0, "txtTm3Em"].Value = string.Empty;
                }
            }
            
            gcMultiRow2.CurrentCell = null;

            // 集計欄
            gcMultiRow3[0, "txtKdays"].Style.BackColor = Color.Empty;
            gcMultiRow3[0, "txtYuDays"].Style.BackColor = Color.Empty;
            gcMultiRow3[0, "txtZdays"].Style.BackColor = Color.Empty;

            gl.ChangeValueStatus = false;   // チェンジバリューステータス

            gcMultiRow3.SetValue(0, "txtKdays", r.公休日数);
            gcMultiRow3.SetValue(0, "txtYuDays", r.有休日数);
            gcMultiRow3.SetValue(0, "txtZdays", r.実労日数);
            gcMultiRow3.SetValue(0, "txtTlDays", r.実労日数 + r.有休日数);
            gcMultiRow3.SetValue(0, "txtChisou", r.遅早時間時);
            gcMultiRow3.SetValue(0, "txtChisou2", r.遅早時間分.ToString("D2"));
            gcMultiRow3.SetValue(0, "txtKotuhi", r.交通費);
            gcMultiRow3.SetValue(0, "txtSonota", r.その他支給);

            txtMemo.Text = r.備考;

            gl.ChangeValueStatus = true;    // チェンジバリューステータス

            //// 日付配列クラスインスタンス作成
            //clsDayItems dItm = new clsDayItems();
            //clsDayItems[] ddd = new clsDayItems[31];

            // 明細表示
            showItem(gcMultiRow1, r);
     
            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示
            _imgFile = Properties.Settings.Default.tifPath + r.画像名.ToString();
            ShowImage(_imgFile);

            //// 確認チェック
            //if (r.確認 == global.flgOff)
            //{
            //    checkBox1.Checked = false;
            //}
            //else
            //{
            //    checkBox1.Checked = true;
            //}

            // ログ書き込み状態とする
            editLogStatus = true;

            // 労働時間集計
            getWorkTimeSection();
            
            WorkTotalSumStatus = true;
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     労働時間集計 </summary>
        ///---------------------------------------------------------
        private void getWorkTimeSection()
        {
            double kihon = global.cnfKihonWh * 60 + global.cnfKihonWm;

            double _w20Time = 0;
            double _w22Time = 0;
            double _naiZan = 0;
            double _mashiZan = 0;
            double _doniShuku = 0;

            // 月間集計値取得
            double wt = getTotalWorkTime(out _w20Time, out _w22Time, out _naiZan, out _mashiZan, out _doniShuku, kihon);

            // 総労働時間表示
            gcMultiRow3.SetValue(0, "txtWorkTime", (int)(wt / 60));
            gcMultiRow3.SetValue(0, "txtWorkTime2", ((int)(wt % 60)).ToString("D2"));

            // 基本時間内残業時間
            gcMultiRow3.SetValue(0, "txtNaiZan", (int)(_naiZan / 60));
            gcMultiRow3.SetValue(0, "txtNaiZan2", ((int)(_naiZan % 60)).ToString("D2"));

            // 割増残業時間
            gcMultiRow3.SetValue(0, "txtMashiZan", (int)(_mashiZan / 60));
            gcMultiRow3.SetValue(0, "txtMashiZan2", ((int)(_mashiZan % 60)).ToString("D2"));

            // 20時以降労働時間表示
            gcMultiRow3.SetValue(0, "txt20Zan", (int)(_w20Time / 60));
            gcMultiRow3.SetValue(0, "txt20Zan2", ((int)(_w20Time % 60)).ToString("D2"));

            // 22時以降労働時間表示
            gcMultiRow3.SetValue(0, "txt22Zan", (int)(_w22Time / 60));
            gcMultiRow3.SetValue(0, "txt22Zan2", ((int)(_w22Time % 60)).ToString("D2"));

            // 土日・祝日労働時間表示
            gcMultiRow3.SetValue(0, "txtHolZan", (int)(_doniShuku / 60));
            gcMultiRow3.SetValue(0, "txtHolZan2", ((int)(_doniShuku % 60)).ToString("D2"));
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠明細表示 </summary>
        /// <param name="hID">
        ///     ヘッダID</param>
        ///------------------------------------------------------------------------------------
        private void showItem(GcMultiRow mr, DataSet1.過去勤務票ヘッダRow r)
        {
            // 勤務実績表示
            int mC = dts.過去勤務票明細.Count(a => a.ヘッダID == r.ID);
                
            // 行数を設定して表示色を初期化
            mr.Rows.Clear();
            mr.RowCount = mC;

            for (int i = 0; i < mC; i++)
            {
                mr.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
                mr.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }
            
            //
            //DateTime nDt = DateTime.Today;
            //DateTime zDt = DateTime.Today;             
            
            //if (DateTime.TryParse((2000 + r.年) + "/" + r.月 + "/01", out nDt))
            //{
            //    // スタッフ情報が取得済みのとき
            //    if (sStf != null)
            //    {
            //        // 前月
            //        zDt = nDt.AddMonths(-1);

            //        dItm.setDayArray(ref ddd, sStf.締日区分.ToString(), zDt, nDt);

            //        // 日にちと曜日を表示する
            //        for (int i = 0; i < ddd.Length; i++)
            //        {
            //            if (ddd[i].sDay != global.flgOff)
            //            {
            //                mr.SetValue(i, "lblDay", ddd[i].sDay);

            //                mr.SetValue(i, "lblWeek", ddd[i].sDayWeek);
            //                mr.SetValue(i, "txtYear", ddd[i].sYear);
            //                mr.SetValue(i, "txtMonth", ddd[i].sMonth);

            //                // 訂正欄チェック
            //                if (isTeisei(i, r))
            //                {
            //                    mr[i, "chkTeisei"].Value = true;
            //                }
            //                else
            //                {
            //                    mr[i, "chkTeisei"].Value = false;
            //                }

            //                // 時刻記号
            //                mr.SetValue(i, "labelCell3", ":");
            //                mr.SetValue(i, "labelCell4", ":");

            //                // 背景色
            //                mr[i, "lblDay"].Style.BackColor = Color.FromArgb(224, 224, 224);
            //                mr[i, "lblWeek"].Style.BackColor = Color.FromArgb(224, 224, 224);

            //                // 選択可能
            //                mr.Rows[i].Selectable = true;
            //            }
            //            else
            //            {
            //                mr.SetValue(i, "lblDay", "");
            //                mr.SetValue(i, "lblWeek", "");
            //                mr.SetValue(i, "txtYear", "");
            //                mr.SetValue(i, "txtMonth", "");
            //                mr[i, "chkTeisei"].Value = false;

            //                // 表示色を初期化
            //                //mr[i, "lblDay"].Style.BackColor = Color.FromName("Control");
            //                //mr[i, "lblWeek"].Style.BackColor = Color.FromName("Control");

            //                // 時刻記号
            //                mr.SetValue(i, "labelCell3", "");
            //                mr.SetValue(i, "labelCell4", "");

            //                // 選択不可
            //                mr.Rows[i].Selectable = false;
            //                mr.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
            //            }
            //        }
            //    }
            //}
            
            // 行インデックス初期化
            int mRow = 0;

            foreach (var t in dts.過去勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID))
            {
                //gl.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                mr[mRow, "txtID"].Value = t.ID;

                DateTime gDt = DateTime.Today;

                string tYYMMDD = t.暦年 + "/" + t.暦月 + "/" + t.日;

                // 無効日付
                if (!DateTime.TryParse(tYYMMDD, out gDt))
                {
                    mr.SetValue(mRow, "lblDay", "");
                    mr.SetValue(mRow, "lblWeek", "");
                    mr.SetValue(mRow, "txtYear", "");
                    mr.SetValue(mRow, "txtMonth", "");
                    mr[mRow, "chkTeisei"].Value = false;
                    mr[mRow, "chkZan"].Value = true;
                    mr[mRow, "txtShopCode"].Value = string.Empty;

                    // 表示色を初期化
                    //mr[i, "lblDay"].Style.BackColor = Color.FromName("Control");
                    //mr[i, "lblWeek"].Style.BackColor = Color.FromName("Control");

                    // 時刻記号
                    mr.SetValue(mRow, "labelCell3", "");
                    mr.SetValue(mRow, "labelCell4", "");

                    // 選択不可
                    mr.Rows[mRow].Selectable = false;
                    mr.Rows[mRow].DefaultCellStyle.BackColor = Color.FromName("Control");
                }
                else
                {
                    // 時刻記号
                    mr.SetValue(mRow, "labelCell3", ":");
                    mr.SetValue(mRow, "labelCell4", ":");

                    // 背景色
                    mr[mRow, "lblDay"].Style.BackColor = Color.FromArgb(224, 224, 224);
                    mr[mRow, "lblWeek"].Style.BackColor = Color.FromArgb(224, 224, 224);

                    // 選択可能
                    mr.Rows[mRow].Selectable = true;

                    // 編集を可能とする
                    mr.Rows[mRow].ReadOnly = false;

                    mr.SetValue(mRow, "txtYear", t.暦年);
                    mr.SetValue(mRow, "txtMonth", t.暦月);
                    mr.SetValue(mRow, "lblDay", t.日);
                    mr[mRow, "lblWeek"].Value = ("日月火水木金土").Substring(int.Parse(gDt.DayOfWeek.ToString("d")), 1);               
                    mr[mRow, "txtStatus"].Value = t.出勤状況;
                    mr[mRow, "txtSh"].Value = t.出勤時;

                    if (t.出勤分 != string.Empty)
                    {
                        mr[mRow, "txtSm"].Value = Utility.StrtoInt(t.出勤分).ToString("D2");
                    }
                    else
                    {
                        mr[mRow, "txtSm"].Value = t.出勤分;
                    }

                    mr[mRow, "txtEh"].Value = t.退勤時;

                    if (t.退勤分 != string.Empty)
                    {
                        mr[mRow, "txtEm"].Value = Utility.StrtoInt(t.退勤分).ToString("D2");
                    }
                    else
                    {
                        mr[mRow, "txtEm"].Value = t.退勤分;
                    }

                    mr[mRow, "txtRest"].Value = t.休憩;

                    // 訂正欄チェック
                    if (isTeisei(mRow, r))
                    {
                        mr[mRow, "chkTeisei"].Value = true;
                    }
                    else
                    {
                        mr[mRow, "chkTeisei"].Value = false;
                    }

                    // 有給申請書チェック
                    if (t.有給申請 == global.flgOff)
                    {
                        mr[mRow, "chkZan"].Value = false;
                    }
                    else
                    {
                        mr[mRow, "chkZan"].Value = true;
                    }

                    // 店舗コード
                    if (t.出勤状況 != string.Empty || t.出勤時 != string.Empty || t.出勤分 != string.Empty ||
                            t.退勤時 != string.Empty || t.退勤分 != string.Empty)
                    {
                        mr[mRow, "txtShopCode"].Value = t.店舗コード.ToString("#");
                    }
                    else
                    {
                        mr[mRow, "txtShopCode"].Value = string.Empty;
                    }
                }

                gl.ChangeValueStatus = true;            // changeValueイベントをtrueに戻す
                
                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            mr.CurrentCell = null;
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     訂正チェック取得 </summary>
        /// <param name="i">
        ///     index</param>
        /// <param name="r">
        ///     DataSet1.過去勤務票ヘッダRow</param>
        /// <returns>
        ///     true:チェックオン, false:チェックオフ</returns>
        ///--------------------------------------------------------------------------
        private bool isTeisei(int i, DataSet1.過去勤務票ヘッダRow r)
        {
            bool rtn = false;

            if (i == 0)
            {
                if (!r.Is訂正1Null() && r.訂正1 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 1)
            {
                if (!r.Is訂正2Null() && r.訂正2 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 2)
            {
                if (!r.Is訂正3Null() && r.訂正3 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 3)
            {
                if (!r.Is訂正4Null() && r.訂正4 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 4)
            {
                if (!r.Is訂正5Null() && r.訂正5 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 5)
            {
                if (!r.Is訂正6Null() && r.訂正6 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 6)
            {
                if (!r.Is訂正7Null() && r.訂正7 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 7)
            {
                if (!r.Is訂正8Null() && r.訂正8 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 8)
            {
                if (!r.Is訂正9Null() && r.訂正9 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 9)
            {
                if (!r.Is訂正10Null() && r.訂正10 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 10)
            {
                if (!r.Is訂正11Null() && r.訂正11 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 11)
            {
                if (!r.Is訂正12Null() && r.訂正12 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 12)
            {
                if (!r.Is訂正13Null() && r.訂正13 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 13)
            {
                if (!r.Is訂正14Null() && r.訂正14 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 14)
            {
                if (!r.Is訂正15Null() && r.訂正15 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 15)
            {
                if (!r.Is訂正16Null() && r.訂正16 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 16)
            {
                if (!r.Is訂正17Null() && r.訂正17 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 17)
            {
                if (!r.Is訂正18Null() && r.訂正18 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 18)
            {
                if (!r.Is訂正19Null() && r.訂正19 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 19)
            {
                if (!r.Is訂正20Null() && r.訂正20 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 20)
            {
                if (!r.Is訂正21Null() && r.訂正21 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 21)
            {
                if (!r.Is訂正22Null() && r.訂正22 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 22)
            {
                if (!r.Is訂正23Null() && r.訂正23 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 23)
            {
                if (!r.Is訂正24Null() && r.訂正24 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 24)
            {
                if (!r.Is訂正25Null() && r.訂正25 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 25)
            {
                if (!r.Is訂正26Null() && r.訂正26 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 26)
            {
                if (!r.Is訂正27Null() && r.訂正27 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 27)
            {
                if (!r.Is訂正28Null() && r.訂正28 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 28)
            {
                if (!r.Is訂正29Null() && r.訂正29 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 29)
            {
                if (!r.Is訂正30Null() && r.訂正30 == global.flgOn)
                {
                    rtn = true;
                }
            }

            if (i == 30)
            {
                if (!r.Is訂正31Null() && r.訂正31 == global.flgOn)
                {
                    rtn = true;
                }
            }

            return rtn;
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     画像を表示する </summary>
        /// <param name="pic">
        ///     pictureBoxオブジェクト</param>
        /// <param name="imgName">
        ///     イメージファイルパス</param>
        /// <param name="fX">
        ///     X方向のスケールファクター</param>
        /// <param name="fY">
        ///     Y方向のスケールファクター</param>
        ///------------------------------------------------------------------------------------
        private void ImageGraphicsPaint(PictureBox pic, string imgName, float fX, float fY, int RectDest, int RectSrc)
        {
            Image _img = Image.FromFile(imgName);
            Graphics g = Graphics.FromImage(pic.Image);

            // 各変換設定値のリセット
            g.ResetTransform();

            // X軸とY軸の拡大率の設定
            g.ScaleTransform(fX, fY);

            // 画像を表示する
            g.DrawImage(_img, RectDest, RectSrc);

            // 現在の倍率,座標を保持する
            gl.ZOOM_NOW = fX;
            gl.RECTD_NOW = RectDest;
            gl.RECTS_NOW = RectSrc;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     フォーム表示初期化 </summary>
        /// <param name="sID">
        ///     過去データ表示時のヘッダID</param>
        /// <param name="cIx">
        ///     勤務票ヘッダカレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize(string sID, string cIx)
        {
            // 表示色設定
            gcMultiRow2[0, "txtYear"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtMonth"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtSNum"].Style.BackColor = SystemColors.Window;
            
            lblNoImage.Visible = false;

            // 編集可否
            gcMultiRow1.ReadOnly = false;
            gcMultiRow2.ReadOnly = true;
            gcMultiRow3.ReadOnly = false;

            txtMemo.ReadOnly = true;
                
            
            //if (_eMode)
            //{
            //    // その他のボタンを有効とする
            //    btnErrCheck.Visible = true;
            //    btnDataMake.Visible = true;
            //    btnDelete.Visible = true;
            //}
            //else
            //{
            //    // 応援移動票画面から遷移のときその他のボタンを無効とする
            //    btnErrCheck.Visible = false;
            //    btnDataMake.Visible = false;
            //    btnDelete.Visible = false;
            //}

            ////データ数表示
            //lblCnt.Text = " (" + (cI + 1).ToString() + "/" + dts.勤務票ヘッダ.Rows.Count.ToString() + ")";
            
            //// 確認チェック欄
            //checkBox1.BackColor = SystemColors.Control;
            //checkBox1.Checked = false;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     エラー表示 </summary>
        /// <param name="ocr">
        ///     OCRDATAクラス</param>
        ///------------------------------------------------------------------------------------
        private void ErrShow(OCRPastData ocr)
        {
            // エラーなし
            if (ocr._errNumber == ocr.eNothing)
            {
                return;
            }

            // グリッドビューCellEnterイベント処理は実行しない
            gridViewCellEnterStatus = false;

            lblErrMsg.Visible = true;
            lblErrMsg.Text = ocr._errMsg;

            //// 確認チェック
            //if (ocr._errNumber == ocr.eDataCheck)
            //{
            //    checkBox1.BackColor = Color.Yellow;
            //    checkBox1.Focus();
            //}

            // 対象年月
            if (ocr._errNumber == ocr.eYearMonth)
            {
                gcMultiRow2[0, "txtYear"].Style.BackColor = Color.Yellow;
                gcMultiRow2.Focus();
                gcMultiRow2.CurrentCell = gcMultiRow2[0, "txtYear"];
                gcMultiRow2.BeginEdit(true);
            } 

            if (ocr._errNumber == ocr.eMonth)
            {
                gcMultiRow2[0, "txtMonth"].Style.BackColor = Color.Yellow;
                gcMultiRow2.Focus();
                gcMultiRow2.CurrentCell = gcMultiRow2[0, "txtMonth"];
                gcMultiRow2.BeginEdit(true);
            }

            // 部署コード
            if (ocr._errNumber == ocr.eBushoCode)
            {
                gcMultiRow2[0, "txtBushoCode"].Style.BackColor = Color.Yellow;
                gcMultiRow2.Focus();
                gcMultiRow2.CurrentCell = gcMultiRow2[0, "txtBushoCode"];
                gcMultiRow2.BeginEdit(true);
            }
                
            // 社員番号
            if (ocr._errNumber == ocr.eShainNo)
            {
                gcMultiRow2[ocr._errRow, "txtSNum"].Style.BackColor = Color.Yellow;
                gcMultiRow2.Focus();
                gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtSNum"];
                gcMultiRow2.BeginEdit(true);
            }
                
            // 出勤状況
            if (ocr._errNumber == ocr.eShukkinStatus)
            {
                gcMultiRow1[ocr._errRow, "txtStatus"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtStatus"];
                gcMultiRow1.BeginEdit(true);
            }

            // 開始時
            if (ocr._errNumber == ocr.eSH)
            {
                gcMultiRow1[ocr._errRow, "txtSh"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtSh"];
                gcMultiRow1.BeginEdit(true);
            }

            // 開始分
            if (ocr._errNumber == ocr.eSM)
            {
                gcMultiRow1[ocr._errRow, "txtSm"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtSm"];
                gcMultiRow1.BeginEdit(true);
            }

            // 終了時
            if (ocr._errNumber == ocr.eEH)
            {
                gcMultiRow1[ocr._errRow, "txtEh"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtEh"];
                gcMultiRow1.BeginEdit(true);
            }

            // 終了分
            if (ocr._errNumber == ocr.eEM)
            {
                gcMultiRow1[ocr._errRow, "txtEm"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtEm"];
                gcMultiRow1.BeginEdit(true);
            }

            // 休憩
            if (ocr._errNumber == ocr.eRest)
            {
                gcMultiRow1[ocr._errRow, "txtRest"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtRest"];
                gcMultiRow1.BeginEdit(true);
            }

            // 有給申請
            if (ocr._errNumber == ocr.eYukyuCheck)
            {
                gcMultiRow1[ocr._errRow, "chkZan"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "chkZan"];
                gcMultiRow1.BeginEdit(true);
            }

            // 日別店舗コード
            if (ocr._errNumber == ocr.eDailyShopCode)
            {
                gcMultiRow1[ocr._errRow, "txtShopCode"].Style.BackColor = Color.Yellow;
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtShopCode"];
                gcMultiRow1.BeginEdit(true);
            }

            // 実労日数
            if (ocr._errNumber == ocr.eWorkDays)
            {
                gcMultiRow3[ocr._errRow, "txtZdays"].Style.BackColor = Color.Yellow;
                gcMultiRow3.Focus();
                gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZdays"];
                gcMultiRow3.BeginEdit(true);
            }

            //  有給日数
            if (ocr._errNumber == ocr.eYukyuDays)
            {
                gcMultiRow3[ocr._errRow, "txtYuDays"].Style.BackColor = Color.Yellow;
                gcMultiRow3.Focus();
                gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtYuDays"];
                gcMultiRow3.BeginEdit(true);
            }

            //  公休日数
            if (ocr._errNumber == ocr.eKoukyuDays)
            {
                gcMultiRow3[ocr._errRow, "txtKdays"].Style.BackColor = Color.Yellow;
                gcMultiRow3.Focus();
                gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtKdays"];
                gcMultiRow3.BeginEdit(true);
            }

            // グリッドビューCellEnterイベントステータスを戻す
            gridViewCellEnterStatus = true;
        }


        ///----------------------------------------------------------
        /// <summary>
        ///     総労働時間取得 </summary>
        /// <returns>
        ///     総労働時間・分</returns>
        ///----------------------------------------------------------
        private double getTotalWorkTime(out double w20Total, out double w22Total, out double naiZan, out double mashiZan, out double doniShuku, double kihon)
        {
            DateTime sDt;
            DateTime eDt;
            DateTime tDt;

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
            
            // 個人基本労働時間
            double pKihon = (Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtWh"].Value))) * 60 + 
                Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtWm"].Value));

            for (int i = 0; i < gcMultiRow1.Rows.Count(); i++)
            {
                sStatus = Utility.NulltoStr(gcMultiRow1[i, "txtStatus"].Value);

                // 有給休暇取得日：2017/12/19
                if (sStatus == global.STATUS_YUKYU)
                {
                    // 時給者のとき基本労働時間を労働時間に加算
                    if (sStf.給与区分.ToString() == global.JIKKYU)
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
                        sSh = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Sh"].Value);
                        sSm = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Sm"].Value);
                        sEh = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Eh"].Value);
                        sEm = Utility.NulltoStr(gcMultiRow2[0, "txtTm1Em"].Value);
                        sRest = REST60;
                    }
                    else if (sStatus == global.STATUS_KIHON_2)
                    {
                        // 基本就業時間2
                        sSh = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Sh"].Value);
                        sSm = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Sm"].Value);
                        sEh = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Eh"].Value);
                        sEm = Utility.NulltoStr(gcMultiRow2[0, "txtTm2Em"].Value);
                        sRest = REST60;
                    }
                    else if (sStatus == global.STATUS_KIHON_3)
                    {
                        // 基本就業時間3
                        sSh = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Sh"].Value);
                        sSm = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Sm"].Value);
                        sEh = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Eh"].Value);
                        sEm = Utility.NulltoStr(gcMultiRow2[0, "txtTm3Em"].Value);
                        sRest = REST60;
                    }
                }
                else
                {
                    sSh = Utility.NulltoStr(gcMultiRow1[i, "txtSh"].Value);
                    sSm = Utility.NulltoStr(gcMultiRow1[i, "txtSm"].Value);
                    sEh = Utility.NulltoStr(gcMultiRow1[i, "txtEh"].Value);
                    sEm = Utility.NulltoStr(gcMultiRow1[i, "txtEm"].Value);
                    sRest = Utility.NulltoStr(gcMultiRow1[i, "txtRest"].Value);
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
                            // 2017/12/19
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
                    
                    // 20時以降勤務時間：直営店の日給者＆直営店の15日締時給者のみ対象 2017/12/19
                    if (sStf.店舗区分.ToString() == global.TENPO_CHOKUEI)
                    {
                        if (sStf.給与区分.ToString() == global.NIKKYU ||
                            (sStf.給与区分.ToString() == global.JIKKYU && sStf.締日区分.ToString() == global.SHIME_15))
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

                    // 土日・祝日の勤務時間：直営店の日給者＆直営店の15日締時給者のみ対象 2017/12/19
                    if (sStf.店舗区分.ToString() == global.TENPO_CHOKUEI)
                    {
                        if (sStf.給与区分.ToString() == global.NIKKYU ||
                            (sStf.給与区分.ToString() == global.JIKKYU && sStf.締日区分.ToString() == global.SHIME_15))
                        {
                            if (Utility.NulltoStr(gcMultiRow1[i, "lblWeek"].Value) == "土" || Utility.NulltoStr(gcMultiRow1[i, "lblWeek"].Value) == "日")
                            {
                                // 土日
                                doniShuku += wTime;
                            }
                            else
                            {
                                // 土日以外の休日
                                if (DateTime.TryParse(Utility.NulltoStr(gcMultiRow1[i, "txtYear"].Value) + "/" +
                                    Utility.NulltoStr(gcMultiRow1[i, "txtMonth"].Value) + "/" +
                                    Utility.NulltoStr(gcMultiRow1[i, "lblDay"].Value), out tDt))
                                {
                                    if (dts.休日.Any(a => a.年月日 == tDt))
                                    {
                                        doniShuku += wTime;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return wTotal;
        }
    }
}
