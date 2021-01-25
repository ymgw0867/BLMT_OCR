using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLMT_OCR.common;

namespace BLMT_OCR.Config
{
    class clsLogin
    {
        public clsLogin()
        {
            adp.Fill(dts.ログインユーザー);
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.ログインユーザーTableAdapter adp = new DataSet1TableAdapters.ログインユーザーTableAdapter();

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     ユーザーが登録されていないとき adminユーザーを自動登録する
        /// </summary>
        /// --------------------------------------------------------------------------
        public void addAdminUser()
        {
            if (dts.ログインユーザー.Count() == 0)
            {
                DataSet1.ログインユーザーRow r = dts.ログインユーザー.NewログインユーザーRow();

                r.アカウント = global.ADMIN_USER;
                r.パスワード = global.ADMIN_PASS;
                r.名前 = "管理者";
                r.備考 = "";
                r.更新年月日 = DateTime.Now;

                dts.ログインユーザー.AddログインユーザーRow(r);

                // データベース登録
                adp.Update(dts.ログインユーザー);
            }

        }

    }
}
