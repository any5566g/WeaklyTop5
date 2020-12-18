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
        private int GetQAQty()
        {

            //string postData = "ver=1&cmd=abf";
            string postData = "action=search";
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);

            string address = "yee";

            address = "http://10.0.5.68/SFIS/Yield/QA/resources/getQueryJSON.jsp?"+
                $"profitCenter={ProdLine}"+
                $"&Include_Rework_WO=Y&PROD_TYPE=ALL"+
                $"&fromDate={FromDate.Value.ToString("yyyy")}%2F{FromDate.Value.ToString("MM")}%2F{FromDate.Value.ToString("dd")}+{FromTime.Value.ToString("HH")}" +
                $"&toDate={ToDate.Value.ToString("yyyy")}%2F{ToDate.Value.ToString("MM")}%2F{ToDate.Value.ToString("dd")}+{ToTime.Value.ToString("HH")}"+
                $"&BU="+
                $"&Customer="+
                $"&MO={WorkOrder.Text.ToUpper()}"+
                $"&modelName="+
                $"&modelSerial={USIPN.Text}" +
                $"&majorProject=" +
                $"&projectName=" +
                $"&productName=" +
                $"&lineName={LineName.Text.ToUpper()}" +
                $"&sectionName=" +
                $"&profit=" +
                $"&profits=" +
                $"&groupName={GroupName.Text.ToUpper()}";

           // address = "http://10.0.5.68/SFIS/Yield/QA/resources/getQueryJSON.jsp?profitCenter=0000000031&Include_Rework_WO=Y&PROD_TYPE=ALL&fromDate=2020%2F07%2F03+04&toDate=2020%2F09%2F04+23&BU=&Customer=&MO=&modelName=&modelSerial=4551-184500-00&majorProject=&projectName=&productName=&lineName=&sectionName=&profit=&profits=&groupName=ALL";

            Console.WriteLine("QA : "+address);
            Uri target = new Uri(address);
            WebRequest request = WebRequest.Create(target);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            //request.ContentType = "application/json";
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
                MessageBox.Show(ex.Message+" || 請確認是否有連線到公司網路", "GetQADate ERROR!");
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
                MessageBox.Show(ex.Message, "GetQAData ERROR!");
                return 1;
            
            }
           
            /*
            sonReader reader = new JsonTextReader(new StringReader(responseStr));
            
            while (reader.Read())
            {
                Console.WriteLine("JsonReader 1: \n"+reader.TokenType + "\t\t" + reader.ValueType + "\t\t" + reader.Value);
            }*/


            JObject obj = (JObject)JsonConvert.DeserializeObject(responseStr);



            string final = Convert.ToString(obj["tableContents"]);

            if (final == "" || final == null)
            {
                MessageBox.Show(@"There is no Data on \SFIS\QA", "No Data Warning !");
                return 1;


            }
            //Regex rx1 = new Regex("[</tdr>class=\"\" nowap]");//會刪除太多 不知道怎麼用
            //final = rx1.Replace(final, "");
            
            List<string> QAList = System.Text.RegularExpressions.Regex.Replace(final, "<[^>]+>", "").Split('\n').ToList();
            int SetNum = (QAList.Count(item => item == "") - 1) / 2;
            QAList.RemoveAll(item => item == "");

            /*
            Console.WriteLine(final);
            Console.WriteLine("-------------------------------------------------------------------------------");
            List<string> fim = final.Split('\n').ToList();
           for (int i = 0; i < fim.Count(); i++)
            {

                if (fim[i].IndexOf("nowrap") != -1)
                {
                    fim[i] = fim[i].Substring(fim[i].IndexOf(">") + 1, (fim[i].IndexOf("</td>") - fim[i].IndexOf(">") - 1));
                }
                else
                {
                    fim[i] = "";
                }

           }
            fim.RemoveAll(item => item == "");

            foreach (var b in fim)
            {
                Console.WriteLine(b);
            }
            
            Console.WriteLine("-------------------------------------------------------------------------------");

            //Deal with data
            List<string> first = final.Split(new string[] { "</tr>" }, StringSplitOptions.RemoveEmptyEntries).ToList();

            // var aaa = fim.FindIndex(x => x == "CQC");
            ///Console.WriteLine(aaa.ToString());
            */


            //Buid DataTable 
            DataTable dt = new DataTable();
            DataRow row;

            //build column
            dt.Columns.Add("No.", typeof(string));
            dt.Columns.Add("Group Name", typeof(string));
            dt.Columns.Add("PASS" , typeof(string));
            dt.Columns.Add("FAIL" , typeof(string));
            dt.Columns.Add("TOTAL" , typeof(string));
            dt.Columns.Add("YIEDL RATE" , typeof(double));
            dt.Columns.Add("REPASS" , typeof(string));
            dt.Columns.Add("REFAIL" , typeof(string));
            dt.Columns.Add("PASS RATE" , typeof(string));



            //fill row
            try
            {
                /* //方法一
                for (int i = 0; i < QAList.Count() - 8; i += 9)
                {
                    row = dt.NewRow();
                    row["No."] = QAList[i];
                    row["Group Name"] = QAList[i+1];
                    row["PASS"] = QAList[i+2];
                    row["FAIL"] = QAList[i+3];
                    row["TOTAL"] = QAList[i+4];
                    row["YIEDL RATE"] = QAList[i+5];
                    row["REPASS"] = QAList[i+6];
                    row["REFAIL"] = QAList[i+7];
                    row["PASS RATE"] = QAList[i+8];

                    dt.Rows.Add(row);
                }*/
                //方法二
                for (int i = 0; i < SetNum; i++)
                {
                    row = dt.NewRow();
                    for (int j = 0; j < QAList.Count() / SetNum; j++)
                    {

                        row[j] = QAList[i * (QAList.Count() / SetNum) + j];

                    }
                    dt.Rows.Add(row);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "QAdata to Form ERROR!");
                return 1;
            }

            //Adjust width
            dataGridView2.DataSource = dt;          
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //dataGridView2.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //dataGridView2.Columns[0].Width = 50;



            //MessageBox.Show(Convert.ToString(obj["tableContents"]));

            //Console.WriteLine(Convert.ToString(obj["tableContents"]));//linq 方式
            //JObject jo = JObject.Parse(responseStr);
            return 0;
        }

    }
}
