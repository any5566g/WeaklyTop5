using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;

namespace WeaklyTop5
{
    public partial class Form1
    {

        private int GetRepairData()
        {

            string postData = "action=search";
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);

            string address = "yee";

            address = "http://10.0.5.68/SFIS/Production/RepairDetail/resources/query2Xlsx.jsp?"+
                $"&profitCenter={ProdLine}"+
                "&Range=T" +
                $"&fromDate={FromDate.Value.ToString("yyyy")}%2F{FromDate.Value.ToString("MM")}%2F{FromDate.Value.ToString("dd")}+{FromTime.Value.ToString("HH")}" +
                $"&toDate={ToDate.Value.ToString("yyyy")}%2F{ToDate.Value.ToString("MM")}%2F{ToDate.Value.ToString("dd")}+{ToTime.Value.ToString("HH")}" +
                $"&sMoNumber={WorkOrder.Text.ToUpper()}" +
                $"&sModelSerial={USIPN.Text}" +
                $"&sModelName=" +
                $"&DateFormat=N" +
                $"&BU=" +
                $"&Result_Type=on" +
                $"&sGroupName={GroupName.Text.ToUpper()}" +
                $"&sMoNumber={WorkOrder.Text.ToUpper()}" +
                $"&sLine={LineName.Text.ToUpper()}";

            Console.WriteLine("Repair : " + address);

            Uri target = new Uri(address);
            WebRequest request = WebRequest.Create(target);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = byteArray.Length;
            request.Timeout = 120000;

            try
            {
                using (var dataStream = request.GetRequestStream())
                {
                    dataStream.Write(byteArray, 0, byteArray.Length);
                    Application.DoEvents();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+" || 請確認是否有連線到公司網路", "GetRepairData ERROR!");
                return 1;

            }

            string responseStr = "";
            //發出Request
            try
            {
                using (WebResponse response = request.GetResponse())
                {
                    using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        responseStr = sr.ReadToEnd();
                    }//end using  
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "GetRepairData ERROR!");
                return 1;

            }


            JObject obj = (JObject)JsonConvert.DeserializeObject(responseStr);
            string abc = Convert.ToString(obj["data"]);

             //2020-10-28 SFIS 資訊變動
            if ( obj.ToString().Contains("No Data") || abc=="")
            {
                MessageBox.Show(@"There is no Data on \SFIS\Repair明細表", "No Data Warning !");
                return 1;
           
            }

            Console.WriteLine(abc);

            string[] abcd = abc.Split('\n');
            char[] ignore = new char[] { ',', '[', ']', '\"', ' ' };
            int j = abcd[1].Length;
            abcd[1] = abcd[1].Substring(0, abcd[1].Length - 2).Trim(ignore);

            Console.WriteLine(abcd[1]);

            try
            {
                WebClient wc = new WebClient();
                wc.DownloadFile("http://10.0.5.68/" + abcd[1], Application.StartupPath + "\\RepairDetail.xlsx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message +"\n請確認關閉RepairDetail.xlsx檔案", "DownLoad RepairData ERROR!");//若例外狀況發生 可能是RepairDetail.xlsx被開啟
                return 1;
            }

            return 0;
        }

    }
}
