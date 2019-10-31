using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BLMT_OCR.common
{
    class clsDayItems
    {
        public int sYear;          // 年
        public int sMonth;         // 月
        public int sDay;           // 日
        public string sDayWeek;    // 曜日
        public int sTeisei;        // 訂正フラグ

        ///-------------------------------------------------------------------
        /// <summary>
        ///     日付配列を作成 </summary>
        /// <param name="cld">
        ///     日付配列クラス</param>
        /// <param name="shime">
        ///     締日</param>
        /// <param name="zDt">
        ///     開始月の前月（勤怠記入開始日の月）</param>
        /// <param name="nDt">
        ///     開始月</param>
        ///-------------------------------------------------------------------
        public void setDayArray(ref clsDayItems [] clds, string shime, DateTime zDt, DateTime nDt)
        {
            int idl = 0;
            int iZ = 1;
            DateTime gDt;

            while (idl < 31)
            {
                int day = Utility.StrtoInt(shime) + iZ;

                clds[idl] = new clsDayItems();

                if (DateTime.TryParse(zDt.Year.ToString() + "/" + zDt.Month.ToString() + "/" + day.ToString(), out gDt))
                {
                    clds[idl].sYear = gDt.Year;
                    clds[idl].sMonth = gDt.Month;
                    clds[idl].sDay = gDt.Day;
                    clds[idl].sDayWeek = ("日月火水木金土").Substring(int.Parse(gDt.DayOfWeek.ToString("d")), 1);
                    clds[idl].sTeisei = global.flgOff;
                }
                else
                {
                    clds[idl].sYear = global.flgOff;
                    clds[idl].sMonth = global.flgOff;
                    clds[idl].sDay = global.flgOff;
                    clds[idl].sDayWeek = string.Empty;
                    clds[idl].sTeisei = global.flgOff;
                }

                if (day == 31)
                {
                    iZ = 1;
                    shime = "";
                    zDt = nDt;
                }
                else
                {
                    iZ++;
                }

                idl++;
            }
        }
    }
}
