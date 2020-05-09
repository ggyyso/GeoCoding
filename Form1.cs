using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Geocoding.Properties;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Data.SqlClient;
using System.Data.OleDb;
namespace Geocoding
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            refreshData();
        }
        /// <summary>
        /// 刷新表格
        /// </summary>
        private void refreshData()
        {
            //this.Refresh();
            //try
            //{
            //    string sql = "select * from LEGAL_PERSON";
            //    dataGridView1.Rows.Clear();
            //    DataSet ds = new DataSet();
            //    DB.getadaoter(sql).Fill(ds, "person");
            //    dataGridView1.DataSource = ds.Tables[0].DefaultView;
            //    this.Refresh();
            //}
            //catch (Exception ex)
            //{

            //}
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            refreshData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }

        private void addressGeocoding()
        {
            if (String.IsNullOrEmpty(this.tbx_key.Text))
            {
                MessageBox.Show("请输入API 密钥key!", "提示");
                return;
            }
            if (String.IsNullOrEmpty(this.tbx_field.Text))
            {
                MessageBox.Show("请输入地址字段!", "提示");
                return;
            }
            if (String.IsNullOrEmpty(this.tbx_mdb.Text))
            {
                MessageBox.Show("请选择数据库!", "提示");
                return;
            }
            if (String.IsNullOrEmpty(this.tbx_tableName.Text))
            {
                MessageBox.Show("请输入表名!", "提示");
                return;
            }
            string result;
            Dictionary<string, string> reDic = null;
            double x = 0, y = 0;

            DateTime Stm;
            DateTime Etm;
            string Selectsql;
            //StreamWriter sw = new StreamWriter("D:\\1.txt", true);
            string addr;
            string value = "";
            int num1 = 0;
            int num2 = 0;
            int once = int.Parse(this.textBox1.Text);//每次提交1000条
            Selectsql = "select count(*) from " + this.tbx_tableName.Text + " where NX is null";
            DataTable dtAll = DB.GetAccessData(Selectsql, this.tbx_mdb.Text);
            int endid = int.Parse(dtAll.Rows[0][0].ToString());//数据库记录总数
            if (dtAll.Rows.Count == 0)
            {
                MessageBox.Show("地址解析已经完成");
                return;
            }
            //標記更新一萬
            int w1 = 0;
            //需要执行的数据区间
            //try
            //{
            //    endid = int.Parse(this.tbx_end.Text);
            //    startid = int.Parse(this.tbx_start.Text);
            //    if(endid<startid){
            //        MessageBox.Show("输入ID范围不正确");
            //        return;
            //    }
            //}
            //catch (System.Exception ex)
            //{
            //    return;
            //}
            //int notInt=0;
            //if (((endid - startid)%once)!=0){
            //    notInt = 1;
            //}
            backgroundWorker1.ReportProgress(endid / once + 1, "5");
            backgroundWorker1.ReportProgress(0, "7");
            string source = "";
            for (int i = 0; i < endid / once + 1; i++)
            {
                //if ((endid -startid)<1000){
                //    Selectsql = "select top " + (endid-startid) + " * from (select top " + endid + " * from zb1000企业 order by UniqID) order by UniqID DESC";
                //}else{
                //    //取出1千条
                //    Selectsql = "select top " + once + " * from (select top " + (i + 1) * once + startid + " * from zb1000企业 order by UniqID) order by UniqID DESC";
                //}
                w1 += once;
                if (w1>10000)
                {
                    MessageBox.Show("已更新一万条");
                    backgroundWorker1.CancelAsync();
                    backgroundWorker1.Dispose();
                    return;
                }
                Selectsql = "select top " + once + " * from (select top " + (i + 1) * once 
                    + " * from  (select * from " + this.tbx_tableName.Text
                    + " where NX is null )A order by ZZJGDM) order by ZZJGDM DESC";
                DataTable dt = DB.GetAccessData(Selectsql, this.tbx_mdb.Text);
                //开始计时
                Stm = DateTime.Now;
                //执行查询
                backgroundWorker1.ReportProgress(dt.Rows.Count +1,"6");
                
                try
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (backgroundWorker1.CancellationPending)
                        {
                            backgroundWorker1.Dispose();
                            return;
                        }
                        addr = dr[this.tbx_field.Text].ToString();
                        if (string.IsNullOrEmpty(addr))
                        {
                            //addr = dr["FRMC"].ToString();
                            continue;
                        }
                        if (radioButton1.Checked)//百度
                        {
                            result = baiduGeocode(addr);
                            reDic = parsonBaidu(result);
                            if (reDic != null)
                            {
                                reDic.TryGetValue("x", out value);
                                dr["NX"] = Convert.ToDouble(value);
                                reDic.TryGetValue("y", out value);
                                dr["NY"] = Convert.ToDouble(value);
                                reDic.TryGetValue("confidence", out value);
                                dr["可信度"] = Convert.ToDouble(value);
                                dr["SOURCE"] = "baidu";
                                value = null;
                            }
                        }
                        else if (radioButton2.Checked)//arcgis
                        {
                            result = AGSGeocode(addr);
                            x = Double.Parse(parsonAGS(result, "x"));
                            y = Double.Parse(parsonAGS(result, "y"));
                            dr["NX"] = x;
                            dr["NY"] = y;
                            dr["SOURCE"] = "arcgis";
                        }
                        else if (this.radioButton3.Checked)//高德
                        {
                            result = AmapGeocode(addr);
                            x = Double.Parse(parsonAMap(result, "x"));
                            y = Double.Parse(parsonAMap(result, "y"));
                            dr["NX"] = x;
                            dr["NY"] = y;
                            dr["SOURCE"] = "高德";
                        }
                        else
                        {//QQ
                            result = QQGeocode(addr);
                            reDic = parsQQ(result);
                        if (reDic != null)
                        {
                            reDic.TryGetValue("x", out value);
                            dr["NX"] = Convert.ToDouble(value);
                            reDic.TryGetValue("y", out value);
                            dr["NY"] = Convert.ToDouble(value);
                            reDic.TryGetValue("similarity", out value);
                            dr["文本相似度"] = Convert.ToDouble(value);
                            reDic.TryGetValue("deviation", out value);
                            dr["误差距离"] = Convert.ToDouble(value);
                            reDic.TryGetValue("reliability", out value);
                            dr["可信度"] = Convert.ToDouble(value);
                            dr["SOURCE"] = "QQ" ;
                            value = null;
                        }
                        }
                      /*  if (this.radioButton4.Checked)
                        {
                            if (reDic != null)
                            {
                                reDic.TryGetValue("x", out value);
                                dr["NX"] = Convert.ToDouble(value);
                                reDic.TryGetValue("y", out value);
                                dr["NY"] = Convert.ToDouble(value);
                                reDic.TryGetValue("similarity", out value);
                                dr["文本相似度"] = Convert.ToDouble(value);
                                reDic.TryGetValue("deviation", out value);
                                dr["误差距离"] = Convert.ToDouble(value);
                                reDic.TryGetValue("reliability", out value);
                                dr["可信度"] = Convert.ToDouble(value);
                                value = null;
                            }
                            else
                            {//QQ获取不到  百度获取
                                result = baiduGeocode(addr);
                                reDic = parsonBaidu(result);
                                if (reDic != null)
                                {
                                    reDic.TryGetValue("x", out value);
                                    dr["NX"] = Convert.ToDouble(value);
                                    reDic.TryGetValue("y", out value);
                                    dr["NY"] = Convert.ToDouble(value);
                                    reDic.TryGetValue("confidence", out value);
                                    dr["可信度"] = Convert.ToDouble(value);
                                    value = null;
                                }
                            }
                            //num++;
                            //if (num / 9 == 1)
                            //{//每秒10次
                            //    Etm = DateTime.Now;
                            //    var dltaTm = Etm - Stm;
                            //    if (dltaTm.Seconds <= 1)//时间间隔小于1秒
                            //    {
                            //        System.Threading.Thread.Sleep(1000);
                            //    }
                            //    Stm = DateTime.Now;
                            //    num = 0;
                            //}
                        }*/
                        //else
                        //{
                        //    dr["NX"] = x;
                        //    dr["NY"] = y;
                        //}
                        System.Threading.Thread.Sleep(200);
                        num1++;
                        backgroundWorker1.ReportProgress(num1,"1");
                        backgroundWorker1.ReportProgress(dt.Rows.Count, "3");
                        //this.label3.Text = "当前 " + this.progressBar1.Value + "/" + dt.Rows.Count;
                    }//end foreach
                    DB.UpDateAccessAdapter(Selectsql, dt, this.tbx_mdb.Text);
                }//end try
                catch (System.Exception ex)
                {
                    //将错误结果记录在界面框中
                    MessageBox.Show(ex.Message);
                }
                ////结束计时
                //Etm = new DateTime();
                //var dltaTm = Etm - Stm;
                //if (dltaTm.Minutes < 10)//时间间隔小于10分钟
                //{
                //    //System.Threading.Thread.Sleep((10 - dltaTm.Minutes) * 60 * 1000);
                //}
                num1=0;
                backgroundWorker1.ReportProgress(num1, "1");
                num2++;
                backgroundWorker1.ReportProgress(num2, "2");
                backgroundWorker1.ReportProgress(endid, "4");
                
            }//end for
            //sw.Close();
            //refreshData();
            backgroundWorker1.ReportProgress(0, "2");
            backgroundWorker1.ReportProgress(0, "4");
        }
        /// <summary>
        /// 腾讯地址解析
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private string QQGeocode(string address)
        {
            //  Encoding myEncoding = Encoding.GetEncoding("UTF8");
            //string addStr=  myEncoding.GetString(Encoding.UTF8.GetBytes(address));
            //string region = address.Substring(0, 6);
            //if (region.Substring(5).Equals("市"))
            //{
            //    region = region.Substring(3, 2);
            //}
            if(address.Contains("#")){
              address= address.Replace("#","号");
            }
            string Url = "http://apis.map.qq.com/ws/geocoder/v1/";
            string param = "?address=" + address + "&key=" + this.tbx_key.Text;
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(Url + param);
            req.Method = "GET";
            using (WebResponse wr = req.GetResponse())
            {
                StreamReader strd = new StreamReader(wr.GetResponseStream(), Encoding.UTF8);
                return strd.ReadToEnd();
            }
        }
        /// <summary>
        /// 解析QQ结果
        /// </summary>
        /// <param name="result"></param>
        /// <param name="XorY"></param>
        /// <returns></returns>
        private Dictionary<string,string> parsQQ(string result)
        {
          
            JObject obj = JObject.Parse(result);
            if (int.Parse(obj["status"].ToString()) != 0)//0为正常
            {
                return null;
            }
            Dictionary<string, string> reDic = new Dictionary<string, string>();
            JObject objResult =(JObject) obj["result"];
            JObject location = (JObject)objResult["location"];
            reDic.Add("x", location["lng"].ToString());
            reDic.Add("y", location["lat"].ToString());
            reDic.Add("similarity", objResult["similarity"].ToString());
            reDic.Add("deviation", objResult["deviation"].ToString());
            reDic.Add("reliability", objResult["reliability"].ToString());

            return reDic;
        }
        /// <summary>
        /// AGS地址解析
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private string AGSGeocode(string address)
        {
            //  Encoding myEncoding = Encoding.GetEncoding("UTF8");
            //string addStr=  myEncoding.GetString(Encoding.UTF8.GetBytes(address));
            string key = "91ba0bc2b233eb48a6323db64fea6599609bb8b034ce2e82a4f69cce3deaceeb";
            string Url = "http://beta.arcgisonline.cn/geocode/"+key+"/single";
            string param = "?citycode=610100&f=json&score=80&Address=" + address + "&queryStr=" + address;
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(Url + param);
            req.Method = "GET";
            using (WebResponse wr = req.GetResponse())
            {
                StreamReader strd = new StreamReader(wr.GetResponseStream(), Encoding.UTF8);
                return strd.ReadToEnd();
            }
        }
        /// <summary>
        /// 解析AGS返回结果
        /// </summary>
        /// <param name="result"></param>
        /// <param name="XorY"></param>
        /// <returns></returns>
        private string parsonAGS(string result, string XorY)
        {
            
            string coord = null;
            JObject obj = JObject.Parse(result);
         JArray array=  (JArray) obj["result"];
            
            if (XorY.ToLower().Equals("x"))
            {
                coord = array[0]["longitude"].ToString();
            }
            else
            {
                coord = array[0]["longitude"].ToString();
            }
            return coord;
        }

        /// <summary>
        /// 百度地址解析
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private string baiduGeocode(string address)
        {
            string Url = "http://api.map.baidu.com/geocoder/v2/";
            string param = "?address=" + address + "&output=json&ak="+ this.tbx_key.Text;
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(Url + param);
            req.Method = "GET";
            using (WebResponse wr = req.GetResponse())
            {
                StreamReader strd = new StreamReader(wr.GetResponseStream(), Encoding.UTF8);
                return strd.ReadToEnd();
            }
        }
        /// <summary>
        /// 解析百度返回结果
        /// </summary>
        /// <param name="result"></param>
        /// <param name="XorY"></param>
        /// <returns></returns>
        private Dictionary<string, string> parsonBaidu(string result)
        {
            string coord = "0";
            try
            {
                JObject obj = JObject.Parse(result);
                if (obj == null) return null;
                if (int.Parse(obj["status"].ToString()) != 0) return null;

                Dictionary<string, string> reDic = new Dictionary<string, string>();
                    reDic.Add("x", obj["result"]["location"]["lng"].ToString());
                    reDic.Add("y", obj["result"]["location"]["lat"].ToString());
                    reDic.Add("precise", obj["result"]["precise"].ToString());
                    reDic.Add("confidence", obj["result"]["confidence"].ToString());
                    return reDic;
            }
            catch(Exception ex){

            }
            return null;
        }

        /// <summary>
        /// 百高德址解析
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private string AmapGeocode(string address)
        {

            //  Encoding myEncoding = Encoding.GetEncoding("UTF8");
            //string addStr=  myEncoding.GetString(Encoding.UTF8.GetBytes(address));
           string Url = "http://restapi.amap.com/v3/geocode/geo?";
            string param = "address=" + address + "&key=23ffc1fbd8cc3f199c86eba87f3c98cc&s=rsv3";
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(Url + param);
            req.Method = "GET";
            using (WebResponse wr = req.GetResponse())
            {
                StreamReader strd = new StreamReader(wr.GetResponseStream(), Encoding.UTF8);
                return strd.ReadToEnd();
            }
        }
        /// <summary>
        /// 解析高德返回结果
        /// </summary>
        /// <param name="result"></param>
        /// <param name="XorY"></param>
        /// <returns></returns>
        private string parsonAMap(string result, string XorY)
        {
            string coord = null;
            JObject obj = JObject.Parse(result);
            coord=(string)obj["count"];
            int count = Convert.ToInt16(coord);
            if (count < 1)
                return "0";
            coord = (string)obj["geocodes"][count - 1]["location"];
            string[] xy=coord.Split(',');
            if (XorY.ToLower().Equals("x"))
            {
                coord = xy[0];
            }
            else
            {
                coord = xy[1]; ;
            }
            return coord;
        }
        //选择数据库
        private void btn_selMDB_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "mdb files (*.mdb)|*.mdb";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.tbx_mdb.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            addressGeocoding();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if(e.UserState=="1"){
                progressBar1.Value = e.ProgressPercentage;
            }
            if (e.UserState == "2")
            {
                progressBar2.Value = e.ProgressPercentage;
            }
            if (e.UserState=="3")
            {
                label3.Text = "当前 " + this.progressBar1.Value + "/" +e.ProgressPercentage;
            }
            if (e.UserState == "4")
            {
                this.label2.Text = progressBar2.Value + "/" + e.ProgressPercentage / 1000 + 1; ;
            }
            if (e.UserState == "5")
            {
                this.progressBar2.Maximum = e.ProgressPercentage;
            }
            if (e.UserState == "6")
            {
                this.progressBar1.Maximum = e.ProgressPercentage;
            }
            if (e.UserState == "7")
            {
                this.progressBar2.Value = e.ProgressPercentage;
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("请求取消");
            }
            else
                MessageBox.Show("解析完成");
    
        }

        private void btn_stop_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sql="select * from 人口 where 级别='5'";
         DataTable dt=   DB.GetAccessData(sql, @"E:\SXGIS\四库\20140625 - 人口-行政区划数据处理.mdb");
         string name = "";
         int cun = 0;
         int zgdw = 0;
            try
            {
                foreach (DataRow dr in dt.Rows)
                {

                    name = dr["行政区划名称"].ToString();
                    if (name.Contains("村") || name.Contains("社区") || name.Contains("庄") ||
                        name.Contains("寨") || name.Contains("家") || name.Contains("居委会") ||
                        name.Contains("居委会") || name.Contains("金滹沱") || name.Contains("家")||
                        name.Contains("火石山"))
                    {
                        cun++;
                    }else
                    {

                    }
                }
            }
            catch (System.Exception ex)
            {
            	
            }

        }
    }


    /// <summary>
    /// 数据库操作类MySql 5.0
    /// </summary>
    class DB
    {
        public static void Excute(string sql)//数据库操作链接方法
        {
            string conn = Settings.Default.ConnectionString;
            MySqlConnection mysql = new MySqlConnection(conn);//实例化链接
            mysql.Open();//开启
            MySqlCommand comm = new MySqlCommand(sql, mysql);
            comm.ExecuteNonQuery();//执行
            mysql.Close();//关闭资源
        }
        public static void updateAdapter(string sql,DataTable dt)//显示操作
        {
            string conn = Settings.Default.ConnectionString;
            MySqlConnection mysql = new MySqlConnection(conn);//实例化链接
            mysql.Open();//开启
            MySqlDataAdapter mda = new MySqlDataAdapter(sql, mysql);
            MySqlCommandBuilder cmd = new MySqlCommandBuilder(mda);
            mda.Update(dt);
            mysql.Close();
            mysql.Dispose();
            //需要在调用的时候进行数据集填充

        }

        public static DataTable getAddress(string sql)//显示操作
        {
            List<string> coordLst = new List<string>();
            string conn = Settings.Default.ConnectionString;
            MySqlConnection mysql = new MySqlConnection(conn);//实例化链接
            mysql.Open();//开启
            //MySqlCommand comm = new MySqlCommand(sql, mysql);
            //comm.ExecuteNonQuery();
            MySqlDataAdapter mda = new MySqlDataAdapter(sql, mysql);
            mysql.Close();
            mysql.Dispose();
            DataSet ds = new DataSet();
            mda.Fill(ds, "address");
            DataTable dt = ds.Tables[0];
            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    foreach (DataRow dr in dt.Rows)
            //    {
            //        coordLst.Add(dr[0].ToString());
            //    }
            //}
            return dt;
        }

        public static DataTable GetAccessData(string sql,string mdbPath)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + mdbPath);
            con.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con);
            DataTable dt=new DataTable();
            da.Fill(dt);
            con.Close(); 
            da.Dispose();
            return dt;
        }

        public static void UpDateAccessAdapter(string sql, DataTable dt, string mdbPath)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + mdbPath);
            con.Open();
            //OleDbCommand cmd = new OleDbCommand(Sql, con);
            //cmd.ExecuteNonQuery();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con);
            OleDbCommandBuilder cb = new OleDbCommandBuilder(da);
            da.Update(dt);
            con.Close();
            da.Dispose();
        }
    }
}
