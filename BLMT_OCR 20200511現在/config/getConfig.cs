using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLMT_OCR.common;

namespace BLMT_OCR.Config
{
    public class getConfig
    {
        DataSet1TableAdapters.環境設定TableAdapter adp = new DataSet1TableAdapters.環境設定TableAdapter();
        DataSet1.環境設定DataTable cTbl = new DataSet1.環境設定DataTable(); 

        public getConfig()
        {
            try
            {
                adp.Fill(cTbl);
                DataSet1.環境設定Row r = cTbl.FindByID(global.configKEY);

                global.cnfYear = r.年;
                global.cnfMonth = r.月;
                global.cnfPath = r.汎用データ出力先;
                global.cnfArchived = r.データ保存月数;
                global.cnfKihonWh = r.基本実労働時;
                global.cnfKihonWm = r.基本実労働分;
                global.cnfMsPath = r.社員マスターパス;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定年月取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
            }
        }
    }
}
