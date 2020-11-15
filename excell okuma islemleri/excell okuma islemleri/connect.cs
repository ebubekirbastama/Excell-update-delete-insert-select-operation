using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;


//Coder By& Ebubekir Bastama bizlere ücretli destek olabilirsiniz.
//Sadece ücretli destekler için: 05554128854

namespace excell_okuma_islemleri
{
    public class connect
    {
        #region Değişkenler
        OleDbConnection EBSCon = new OleDbConnection();//Bağlantı nesnesi.
        OleDbCommand EBSCommand = new OleDbCommand();// Command nesnesi.
        public string sayfaadi = "";//Globalde aldığımız sayfa adı.
        public string xlsxad;//Excel Yolu ve adı.
        #endregion
        #region Metodlar
        /// <summary>
        /// Bu metodun Görevi Otomatik olarak bağlantınızı açmaktır...
        /// </summary>
        /// <returns></returns>
        public OleDbConnection EBSconneciton()
        {
            this.EBSCon.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source ='" + xlsxad + "'; Extended Properties = Excel 12.0;";

            bool flag = this.EBSCon.State == ConnectionState.Closed;
            if (flag)
            {
                this.EBSCon.Open();
            }
            return this.EBSCon;
        }
        /// <summary>
        /// Bu Metod ile ilk sayfa adını alıp işleyebilirsiniz.
        /// </summary>
        /// <returns></returns>
        public OleDbConnection EBSconnecitonExcellsayfaadi()
        {
            this.EBSCon.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source ='" + xlsxad + "'; Extended Properties = Excel 12.0;";

            bool flag = EBSCon.State == ConnectionState.Closed;
            if (flag)
            {
                this.EBSCon.Open();
                sayfaadi = EBSCon.GetSchema("Tables").Rows[1].ItemArray[2].ToString();
                EBSdisconneciton();
            }
            return this.EBSCon;
        }
        /// <summary>
        /// Bu Metod ile Bütün tabloyu alabilir ve bütün sayfa adları ile işlemler yapabilirsiniz.
        /// </summary>
        /// <returns></returns>
        public DataTable EBSconnecitonExcellsayfaadiDatatable()
        {
            EBSCon.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source ='" + xlsxad + "'; Extended Properties = Excel 12.0;";
            DataTable dt = new DataTable();
            bool flag = this.EBSCon.State == ConnectionState.Closed;
            if (flag)
            {
                this.EBSCon.Open();
                dt = EBSCon.GetSchema("Tables");
                EBSdisconneciton();
            }
            return dt;
        }
        /// <summary>
        /// Bu metodun görevi ilgili bağlantıyı kapatmaktır...
        /// </summary>
        /// <returns></returns>
        public OleDbConnection EBSdisconneciton()
        {
            bool flag = EBSCon.State == ConnectionState.Open;
            if (flag)
            {
                EBSCon.Close();
            }
            return EBSCon;
        }
        /// <summary>
        /// Bu metodun görevi verdiğiniz inser,update ve ya delete komutlarını işlemektir.
        /// </summary>
        /// <param name="cmd"></param>
        public void xlscmd(string cmd)
        {
            // "update [Ürünler$] set Resim='" + cmd[0] + "' where [Barkod]=" + cmd[1] + "";
            OleDbCommand kmt = new OleDbCommand(cmd, EBSconneciton());
            kmt.ExecuteNonQuery();
            EBSdisconneciton();

        }
        /// <summary>
        /// Bu metodun görevi  OleDbDataReader ile kolon belirterek veri çekmektir.
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="lst"></param>
        public void Dataa(string cmd, ListBox lst)
        {
            OleDbCommand kmt = new OleDbCommand(cmd, EBSconneciton());
            OleDbDataReader rdr = kmt.ExecuteReader();
            while (rdr.Read())
            {
                if (rdr["Barkod"].ToString() != "")
                {
                    lst.Items.Add(rdr["Barkod"].ToString().Replace(".", "").Trim());
                }
            }
            EBSdisconneciton();
        }
        /// <summary>
        /// Bu metodun görevi Select ettiğiniz verileri verdiğiniz DataGridView nesnesine aktarır.
        /// </summary>
        /// <param name="cml"></param>
        /// <param name="dg"></param>
        public void excelldata(string cml, DataGridView dg)
        {
            DataTable dt = new DataTable();
            OleDbDataAdapter adtr = new OleDbDataAdapter(cml, EBSconneciton());
            adtr.Fill(dt);
            dg.DataSource = dt;
            EBSdisconneciton();
        }
        #endregion
    }
}
