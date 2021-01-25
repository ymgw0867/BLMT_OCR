using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;

namespace BLMT_OCR.OCR
{
    public partial class frmFaxSelect : Form
    {
        public frmFaxSelect()
        {
            InitializeComponent();

            /* テーブルアダプターマネージャーに勤務票ヘッダ、明細テーブル、
             * 保留勤務票ヘッダ、保留明細テーブルアダプターを割り付ける */
            adpMn.勤務票ヘッダTableAdapter = hAdp;
            adpMn.勤務票明細TableAdapter = iAdp;

            dAdpMn.保留勤務票ヘッダTableAdapter = dhAdp;
            dAdpMn.保留勤務票明細TableAdapter = diAdp;

            // 勤務票データ読み込み
            adpMn.勤務票ヘッダTableAdapter.Fill(dts.勤務票ヘッダ);
            adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);

            // 保留データ読み込み
            dAdpMn.保留勤務票ヘッダTableAdapter.Fill(dts.保留勤務票ヘッダ);
            dAdpMn.保留勤務票明細TableAdapter.Fill(dts.保留勤務票明細);
        }
        
        int ALLCnt = 0;

        DataSet1 dts = new DataSet1();

        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();

        DataSet1TableAdapters.TableAdapterManager dAdpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.保留勤務票ヘッダTableAdapter dhAdp = new DataSet1TableAdapters.保留勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.保留勤務票明細TableAdapter diAdp = new DataSet1TableAdapters.保留勤務票明細TableAdapter();

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void frmFaxSelect_Load(object sender, EventArgs e)
        {
            //ALLCnt = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv").Count();

            foreach (var t in dts.保留勤務票ヘッダ.OrderBy(a => a.更新年月日))
            {
                checkedListBox1.Items.Add(t.更新年月日.ToShortDateString() + " " + t.更新年月日.Hour.ToString("D2") + ":" + t.更新年月日.Minute.ToString("D2") + ":" + t.更新年月日.Second.ToString("D2") + ", " + t.ID);                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            myCnt = 0;
            myBool = false;
            Close();
        }

        public int myCnt { get; set; }
        public bool myBool { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            //int n = getDataCount();

            //if (n == 0)
            //{
            //    MessageBox.Show("処理するデータがありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
            
            myBool = true;

            // 保留データをＦＡＸ発注書データに戻す
            holdToData();

            Close();
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     保留データをＦＡＸ発注書データに戻す </summary>
        ///-------------------------------------------------------
        private void holdToData()
        {
            if (checkedListBox1.SelectedItems.Count == 0)
            {
                return;
            }

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                string s = checkedListBox1.CheckedItems[i].ToString();
                string[] st = s.Split(',');

                setHoldToData(st[1].Trim());
            }
        }
        
        ///----------------------------------------------------------
        /// <summary>
        ///     保留データを戻す </summary>
        /// <param name="iX">
        ///     データインデックス</param>
        ///----------------------------------------------------------
        private void setHoldToData(string iX)
        {
            try
            {
                var t = dts.保留勤務票ヘッダ.Single(a => a.ID == iX);

                DataSet1.勤務票ヘッダRow hr = dts.勤務票ヘッダ.New勤務票ヘッダRow();

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

                if (t.Is基本実労働時Null())
                {
                    hr.基本実労働時 = string.Empty;
                }
                else
                {
                    hr.基本実労働時 = t.基本実労働時;
                }

                if (t.Is基本実労働分Null())
                {
                    hr.基本実労働分 = string.Empty;
                }
                else
                {
                    hr.基本実労働分 = t.基本実労働分;
                }

                // 特休日数：2020/05/12
                if (t.Is特休日数Null())
                {
                    hr.特休日数 = 0;
                }
                else
                {
                    hr.特休日数 = t.特休日数;
                }


                // 保留データ追加処理
                dts.勤務票ヘッダ.Add勤務票ヘッダRow(hr);

                // 保留勤務票明細
                var mm = dts.保留勤務票明細.Where(a => a.ヘッダID ==  iX).OrderBy(a => a.ID);
                foreach (var m in mm)
                {
                    DataSet1.勤務票明細Row mr = dts.勤務票明細.New勤務票明細Row();

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

                    // 勤務票明細データ追加処理
                    dts.勤務票明細.Add勤務票明細Row(mr);
                }

                // データベース更新
                adpMn.UpdateAll(dts);

                // 保留データ削除
                t.Delete();                 // 保留出勤簿ヘッダ
                
                foreach (var item in mm)    // 保留出勤簿明細
                {
                    item.Delete();
                }

                dAdpMn.UpdateAll(dts);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        ///-----------------------------------------------------------
        /// <summary>
        ///     処理データ件数を取得する </summary>
        /// <returns>
        ///     データ件数 </returns>
        ///-----------------------------------------------------------
        private int getDataCount()
        {
            int dCnt = 0;

            dCnt += ALLCnt;
            dCnt += checkedListBox1.CheckedItems.Count;

            return dCnt;
        }
    }
}
