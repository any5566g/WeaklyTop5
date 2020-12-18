using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.IO;
using System.Runtime;
using ExcelDataReader;
using System.Runtime.CompilerServices;
using OfficeOpenXml.Table.PivotTable;
using System.Security.Cryptography;
using System.Xml;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System.Diagnostics;

namespace WeaklyTop5
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        string ProdLine = "";




        public Form1()
        {
            InitializeComponent();
            //GetExcel();

            
            // Sort();
            //MakeExcel();
        }

        //private class ComboBox(string text , string value)
        





        private int GetExcel()
        {
            //string Path = @"C:\Users\Administrator\Desktop\CS1805\Repair.xlsx";
            string Path = Application.StartupPath + @"\RepairDetail.xlsx";

            try
            {
                using (var stream = new FileStream(Path, FileMode.Open, FileAccess.ReadWrite))
                {

                    IExcelDataReader reader = null;
                    /* 
                     reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                     reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                     {
                         FallbackEncoding = Encoding.GetEncoding("big5")
                     });*/ //.xls   binary 2003 以前版本

                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream); //.xlsx 2007年後的excel

                    if (reader == null)
                    {
                        Console.WriteLine("未知的處理檔案：");
                        Console.ReadKey();

                    }

                    using (reader)
                    {



                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            //UseColumnDataType = false,//true return type will be system.string , false will be system.object
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                //設定讀取資料時是否忽略標題
                                UseHeaderRow = true
                            }
                        });

                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        dataGridView1.DataSource = ds.Tables[0];

                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Read Data Fail \n{ex.Message} ||  Please put RepairDetail.xlsx to starup folder", "Error");
                return 1;
            }

            return 0;
        }




        private int Sort()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;//EEplus license

            var file = new FileInfo(Application.StartupPath + $"\\總表{DateTime.Now.ToString("yyyy-MM-dd")}.xlsx");

            using (var excel = new ExcelPackage())
            {

                var ws = excel.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B2"].LoadFromDataTable(ds.Tables[0], true);

                /*
                ExcelWorksheet worksheetPivot =excel.Workbook.Worksheets.Add("Pivot");
                var worksheetData = excel.Workbook.Worksheets.Add("Sheet2");

                worksheetData.Cells["A1"].Value = "Column A";
                worksheetData.Cells["A2"].Value = "Group A";
                worksheetData.Cells["A3"].Value = "Group B";
                worksheetData.Cells["A4"].Value = "Group C";
                worksheetData.Cells["A5"].Value = "Group A";
                worksheetData.Cells["A6"].Value = "Group B";
                worksheetData.Cells["A7"].Value = "Group C";
                worksheetData.Cells["A8"].Value = "Group A";
                worksheetData.Cells["A9"].Value = "Group B";
                worksheetData.Cells["A10"].Value = "Group C";
                worksheetData.Cells["A11"].Value = "Group D";

                worksheetData.Cells["B1"].Value = "Column B";
                worksheetData.Cells["B2"].Value = "emc";
                worksheetData.Cells["B3"].Value = "fma";
                worksheetData.Cells["B4"].Value = "h2o";
                worksheetData.Cells["B5"].Value = "emc";
                worksheetData.Cells["B6"].Value = "fma";
                worksheetData.Cells["B7"].Value = "h2o";
                worksheetData.Cells["B8"].Value = "emc";
                worksheetData.Cells["B9"].Value = "fma";
                worksheetData.Cells["B10"].Value = "h2o";
                worksheetData.Cells["B11"].Value = "emc";

                worksheetData.Cells["C1"].Value = "Column C";
                worksheetData.Cells["C2"].Value = 299;
                worksheetData.Cells["C3"].Value = 792;
                worksheetData.Cells["C4"].Value = 458;
                worksheetData.Cells["C5"].Value = 299;
                worksheetData.Cells["C6"].Value = 792;
                worksheetData.Cells["C7"].Value = 458;
                worksheetData.Cells["C8"].Value = 299;
                worksheetData.Cells["C9"].Value = 792;
                worksheetData.Cells["C10"].Value = 458;
                worksheetData.Cells["C11"].Value = 299;

                worksheetData.Cells["D1"].Value = "Column D";
                worksheetData.Cells["D2"].Value = 40075;
                worksheetData.Cells["D3"].Value = 31415;
                worksheetData.Cells["D4"].Value = 384400;
                worksheetData.Cells["D5"].Value = 40075;
                worksheetData.Cells["D6"].Value = 31415;
                worksheetData.Cells["D7"].Value = 384400;
                worksheetData.Cells["D8"].Value = 40075;
                worksheetData.Cells["D9"].Value = 31415;
                worksheetData.Cells["D10"].Value = 384400;
                worksheetData.Cells["D11"].Value = 40075;
                //---------------------------------------------------------------------------------------------------------------------------------
                var dataRange = worksheetData.Cells[worksheetData.Dimension.Address];


                var pivotTable = worksheetPivot.PivotTables.Add(worksheetPivot.Cells["B2"], dataRange, "PivotTable");

                //label field
                pivotTable.RowFields.Add(pivotTable.Fields["Column A"]);
                pivotTable.DataOnRows = false;

                //data fields
                var field = pivotTable.DataFields.Add(pivotTable.Fields["Column B"]);
                field.Name = "Count of Column B";
                field.Function = DataFieldFunctions.Count;
                

                field = pivotTable.DataFields.Add(pivotTable.Fields["Column C"]);
                field.Name = "Sum of Column C";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "0.00";

                field = pivotTable.DataFields.Add(pivotTable.Fields["Column D"]);
                field.Name = "Sum of Column D";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "€#,##0.00";

                */
                //------------------------------------------------------------------------------------------------------------------------------

                var dataRange1 = ws.Cells[ws.Dimension.Address];


                var pivotTable1 = ws.PivotTables.Add(ws.Cells["X2"], dataRange1, "PivotTable1");

                //label field
                pivotTable1.RowFields.Add(pivotTable1.Fields["ERROR_DESC"]);
                pivotTable1.RowFields.Add(pivotTable1.Fields["ERROR_CODE"]);
                pivotTable1.DataOnRows = false;

                //data fields
                var field1 = pivotTable1.DataFields.Add(pivotTable1.Fields["SERIAL_NUMBER"]);
                field1.Name = "Count of SERIAL_NUMBER";
                field1.Function = DataFieldFunctions.Count;


                //field1 = pivotTable1.DataFields.Add(pivotTable.Fields["Column C"]);
                //field1.Name = "Sum of Column C";
                //field1.Function = DataFieldFunctions.Sum;
                //field1.Format = "0.00";

                //field1 = pivotTable.DataFields.Add(pivotTable.Fields["Column D"]);
                // field1.Name = "Sum of Column D";
                // field1.Function = DataFieldFunctions.Sum;
                // field1.Format = "€#,##0.00";
                //worksheetPivot.Cells["A1:Z50"].AutoFitColumns

                // ws.Cells[ws.Dimension.Address].Sort(6);


                //----------------------------------------------------------------------------------------------------

                //data Cells
                //var wsData = excel.Workbook.Worksheets["Dataa"];
                //wsData.Cells.LoadFromDataTable(ds.Tables[0],true);
                ExcelRange dataCells = ws.Cells[ws.Dimension.Address];

                //pivot table 
                //ExcelRange pvtLocation = new ExcelPackage(n).Workbook.Worksheets.cell["B2"];
                //string pvtName = "Name";
                ExcelPivotTable pivotTable = ws.PivotTables.Add(ws.Cells["AB4"], dataCells, "myPvt");

                //header
                pivotTable.ShowHeaders = true;
                pivotTable.RowHeaderCaption = "ERROR";

                //grand total
                pivotTable.ColumnGrandTotals = true;
                pivotTable.GrandTotalCaption = "Total";

                //data field are plcae in columns
                pivotTable.DataOnRows = false;

                //style 
                pivotTable.TableStyle = OfficeOpenXml.Table.TableStyles.Medium9;

                //filter
                ExcelPivotTableField TitleField = pivotTable.PageFields.Add(pivotTable.Fields["TEST_GROUP"]);

                TitleField.Sort = eSortType.Ascending;

                //row:Error
                ExcelPivotTableField ErroeCodefiled = pivotTable.RowFields.Add(pivotTable.Fields["ERROR_CODE"]);
                ExcelPivotTableField ErroeCodefiled2 = pivotTable.RowFields.Add(pivotTable.Fields["ERROR_DESC"]);

                //valure serial number
                ExcelPivotTableDataField SerialnumberField = pivotTable.DataFields.Add(pivotTable.Fields["SERIAL_NUMBER"]);

                //column
                ExcelPivotTableField Dutyfield = pivotTable.ColumnFields.Add(pivotTable.Fields["DUTY_TYPE"]);

                SerialnumberField.Function = DataFieldFunctions.Count;
                //SerialnumberField.Format = "#,##0_);(#,##0)";
                SerialnumberField.Name = "EEEE";



                pivotTable.SortOnDataField(ErroeCodefiled, SerialnumberField, true);







                /*
                ExcelPivotTableField fields = ErroeCodefiled;
                ExcelPivotTableDataField dataField = SerialnumberField;

                bool descending = true;

                var xdoc = pivotTable.PivotTableXml;
                var nsm = new XmlNamespaceManager(xdoc.NameTable);

                var schemaMain = xdoc.DocumentElement.NamespaceURI;
                if (nsm.HasNamespace("x") == false)
                    nsm.AddNamespace("x", schemaMain);

                // <x:pivotField sortType="descending">
                var pivotField = xdoc.SelectSingleNode(
                    "/x:pivotTableDefinition/x:pivotFields/x:pivotField[position()=" + (fields.Index + 1) + "]",
                    nsm
                );


                
                pivotField.("sortType", (descending ? "descending" : "ascending"));

                // <x:autoSortScope>
                var autoSortScope = pivotField.AppendElement(schemaMain, "x:autoSortScope");

                // <x:pivotArea>
                var pivotArea = autoSortScope.AppendElement(schemaMain, "x:pivotArea");

                // <x:references count="1">
                var references = pivotArea.AppendElement(schemaMain, "x:references");
                references.AppendAttribute("count", "1");

                // <x:reference field="4294967294">
                var reference = references.AppendElement(schemaMain, "x:reference");
                // Specifies the index of the field to which this filter refers.
                // A value of -2 indicates the 'data' field.
                // int -> uint: -2 -> ((2^32)-2) = 4294967294
                reference.AppendAttribute("field", "4294967294");

                // <x:x v="0">
                var x = reference.AppendElement(schemaMain, "x:x");
                int v = 0;
                foreach (ExcelPivotTableDataField pivotDataField in pivotTable.DataFields)
                {
                    if (pivotDataField == dataField)
                    {
                        x.AppendAttribute("v", v.ToString());
                        break;
                    }
                    v++;
                }
                */
                ws.Cells[ws.Dimension.Start.Row, ws.Dimension.Start.Column, ws.Dimension.End.Row, ws.Dimension.End.Column].AutoFitColumns();//寬度自動調整

                try
                {
                    excel.SaveAs(file);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message , "Pivot Sort ERROR! || Save File ERROR");
                    return 1;
                }

            }
            return 0;
        }

        private int MakeExcel()
        {
            // 可以試試看 用list List<string> list = new List<string>() // 之後用list.add  
            HashSet<string> EC2 = new HashSet<string>();

            /* //List 特性
            List<List<string>> aaaa = new List<List<string>>();
            aaaa.Add(new List<string> { "aaa","bbbb"});
            aaaa.Add(new List<string> { "e", "ff", "ggg" });
            */

            List<String[]> Data = new List<string[]>();
            List<List<string>> FinalDutyCount = new List<List<string>>();


            //MessageBox.Show(EC3.Count().ToString());

            foreach (var a in (from i in ds.Tables[0].AsEnumerable()
                               where i.Field<string>("ERROR_CODE") != null
                               select i))
            {
                //Console.WriteLine($@"ErrorCode =  {a["ERROR_CODE"]} || RepairTime = {a["REPAIR_TIME"]} || Duty = {a["DUTY_TYPE"]}");

                EC2.Add(a["ERROR_CODE"].ToString());

            }

            

            foreach (var ECname in EC2)
            {
                var count = ds.Tables[0].AsEnumerable().Count(x => (string)x["ERROR_CODE"] == ECname);

                
                string Name = "";

                
                foreach (var duty in (from d in ds.Tables[0].AsEnumerable()
                                      where d["ERROR_CODE"].ToString()==ECname
                                      select d))
                {
                    Name = duty["ERROR_DESC"].ToString();
                
                }
                
                Data.Add(new string[] { ECname, count.ToString() , Name }); //Data Format Data(ErrorCode , Count , ErrorDescription)
                Console.WriteLine($@"ECNAME = {ECname} || count = {count} || ERROR_DESC = {Name}");

            }

            /*
            for (int i = 0; i < Data.Count; i++)
            {
                Console.WriteLine($@"Data = {Data[i][1]}");//[i][0]=errorcode [i][1]=count [i][2]=B [i][3]=N [i][4]=P [i][5]=W [i][6]=R
            }
            */

            //交換順序 Bubble Sort Big to Small
            Console.WriteLine($@"Data Length = {Data.Count}");

            List<string[]> temp = new List<string[]>();
            temp.Add(new string[] { "EC", "count" , "EC_DESC" });



            try
            {

                for (int i = 0; i < 3; i++)
                {
                    Console.WriteLine($@"Temp = {temp[0][i]}");
                }


                for (int i = 0; i < Data.Count; i++)
                {
                    for (int j = i; j < Data.Count; j++)
                    {


                        if (Convert.ToInt32(Data[j][1]) > Convert.ToInt32(Data[i][1]))
                        {
                            for (int k = 0; k < Data[j].Length; k++)
                            {
                                temp[0][k] = Data[j][k];
                            }

                            for (int k2 = 0; k2 < Data[i].Length; k2++)
                            {
                                Data[j][k2] = Data[i][k2];

                            }

                            for (int k3 = 0; k3 < temp[0].Length; k3++)
                            {
                                Data[i][k3] = temp[0][k3];
                            }

                        }


                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Sorting Fail");
                return 1;
            }


            //get total
            for (int i = 0; i < Data.Count; i++)
            {

                HashSet<string> Duty = new HashSet<string>();
                List<string> DNum = new List<string>(); //put number in here
                

                var Raw =  from d in ds.Tables[0].AsEnumerable()
                           where d["ERROR_CODE"].ToString() == Data[i][0]
                           select d;

                
                //Different Duty Type
                foreach (var duty in Raw)
                {                             
                   Duty.Add(duty["DUTY_TYPE"].ToString());

                }
                Duty.Remove("");
               
                //Count Duty Pcs
                foreach (var Type in Duty)
                {
                    //var DutyCount = Raw.Count(x => (x["DUTY_TYPE"]==DBNull.Value) ? (string)x["DUTY_TYPE"] == Type :  (string)x["DUTY_TYPE"] == Type); //Linq 空的CELL會報錯
                    var DutyCount = (from f in Raw
                                    where f["DUTY_TYPE"] != DBNull.Value
                                    where (string)f["DUTY_TYPE"] == Type
                                    select f).Count();

                    DNum.Add(DutyCount.ToString());
                    //Console.WriteLine($"ERROR = {Data[i][0]} || Duty = {Type} || count = {DutyCount}");
                    
                }
                // Write Total information
                FinalDutyCount.Add(Duty.ToList());
                FinalDutyCount.Add(DNum);

               
            }

            for (int i = 0; i < FinalDutyCount.Count; i += 2)
            {
                if (FinalDutyCount[i].Count != FinalDutyCount[i + 1].Count)//FinalDutyCount[odd] == 單一ErrorCode中不同的DutyType  FinalDutyCount[even] == 單一ErrorCode中不同DutyType的數量 
                {
                    MessageBox.Show("Duty Number dosen't Match !!", "FinalDataCount Error");
                
                }
                Console.Write(Data[i / 2][0] + " ");
                for (int j =0; j< FinalDutyCount[i].Count; j++)
                { 

                  Console.Write(FinalDutyCount[i][j] + " = ");
                  Console.Write(FinalDutyCount[i+1][j] + " || ");
                    
            
                    
                    
                }
                Console.WriteLine("");
            
        }




            //Output Errorcode & Each total count
            for (int i = 0; i < Data.Count; i++)
            {
                for (int j = 0; j < Data[i].Length; j++)
                {
                    Console.Write($"{Data[i][j]}  ");

                }
                Console.WriteLine("\n");
            }



            // save file
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;//EEplus license
            var file = new FileInfo(Application.StartupPath + $"\\Top5_std.xlsx");

            using (var excel = new ExcelPackage(file))
            {

                var ws = excel.Workbook.Worksheets[0];//可以寫進現有worksheet
                                                      // ws.Cells["A1"].LoadFromDataTable(ds2.Tables[0], true);

                

                //Top 5 Count information 
                for (int i = 0; (Data.Count<=5) ? i < Data.Count : i < 5 ; i++)
                {
                    ws.Cells[$"J{3 + i}"].Value = Data[i][2];
                    ws.Cells[$"K{3 + i}"].Value = double.Parse(Data[i][1]);
                    //ws.Cells[$"K{3 + i}"].Style.QuotePrefix = false;
                    ws.Cells[$"K{3 + i}"].Style.Numberformat.Format = "0";//數字樣式

                }

                /*
                int[] ss = { 18, 27, 36, 45, 54 };
                for (int j = 0; j < 5; j++)
                {
                    for (int k = 0; k < 8; k++)
                    {
                        ws.Cells[$"H{ss[j] + k}"].Value = double.Parse(Data[j][2 + k]);
                        ws.Cells[$"H{ss[j] + k}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                }*/


                //ws.InsertRow(19, 1,18); //Insert parament :: 1.which row 2.how many you want to add 2. duplicate  formula form which row

                
                //定位 插列
                int topN = (Data.Count > 5) ? 5 : Data.Count();

                var searchCell = from cell in ws.Cells[ws.Dimension.Start.Row, ws.Dimension.Start.Column, ws.Dimension.End.Row, ws.Dimension.End.Column] //you can define your own range of cells for lookup
                                 where cell.Text == "FA"
                                 select cell.Start.Row;


                foreach (var cc in searchCell.Reverse())
                {
                    Console.WriteLine("SearchCell \"FA\" ROW number : " + cc);
                    Console.WriteLine("\n");
                }


                var SearchReverseNew = (Data.Count >= 5) ? searchCell.Reverse() : searchCell.Reverse().Skip(5-Data.Count());  
                foreach (var value in SearchReverseNew)
                {
                   Console.WriteLine("SearchCell \"FA\" ROW number : "+value);
                   ws.InsertRow(value, (FinalDutyCount[topN * 2-1].Count)==0 ? 0 : (FinalDutyCount[topN * 2 - 1].Count - 1), 18); //insert((插哪一列(新增的會直接變成該列), 插幾列, 複製第幾列格式)) 2020-10-30
                   topN--;

                }
                
                //Fill the cell 
                List<string> TopNum = new List<string>() { "Top.1" , "Top.2", "Top.3", "Top.4", "Top.5"};

                for (int i = 0; (Data.Count <= 5) ? i < Data.Count : i < 5 ; i++)
                {

                    var SearchTopN = from cell in ws.Cells[ws.Dimension.Start.Row, ws.Dimension.Start.Column, ws.Dimension.End.Row, ws.Dimension.End.Column] //you can define your own range of cells for lookup
                                     where cell.Text == TopNum[i]
                                     select cell.Start.Row;
                    foreach (var vv in SearchTopN)
                    {
                        var one = ws.Cells[$"I{vv}"];
                        for (int j = 0; j < FinalDutyCount[i*2].Count(); j++)
                        {
                            ws.Cells[$"K{vv + j}"].Value = FinalDutyCount[i*2][j]; //Duty Type
                            ws.Cells[$"H{vv + j}"].Value = double.Parse(FinalDutyCount[i*2+1][j]); //Duty count
                            ws.Cells[$"K{vv + j}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            ws.Cells[$"H{vv + j}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // ws.Cells[$"I{vv + j}"].Formula = ws.Cells["I18"].Formula;
                            one.Copy(ws.Cells[$"I{vv + j}"]);

                        }

                    }
                   
                }
                
               


                ws.Cells["K9"].Value = double.Parse((from i in ds.Tables[0].AsEnumerable()
                                                     where i.Field<string>("NO") != null
                                                     select i).Count().ToString());
                //ws.Cells["I18"].Copy(ws.Cells["D35"]);
                // ws.Cells["D33"].Formula = ws.Cells["I18"].Formula;

                try
                {
                    //save the changes
                    excel.SaveAs(new FileInfo(Application.StartupPath + $"\\Top5_{DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss")}.xlsx"));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message , "MakeExcel ERROR || Save File ERROR!");
                    return 1;
                }

            }


            return 0;
        }
        private void Save()//沒用到
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;//EEplus license
            var file = new FileInfo(Application.StartupPath + $"\\Top5.xlsx");

            using (var excel = new ExcelPackage(file))
            {

                var ws = excel.Workbook.Worksheets[0];//可以寫盡現有worksheet
                                                      // ws.Cells["A1"].LoadFromDataTable(ds2.Tables[0], true);

                //add some data
                ws.Cells["R2"].Value = "Added data in Cell A4";
                ws.Cells["R3"].Value = "Added data in Cell B4";





                try
                {
                    //save the changes
                    excel.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Combobox Setting
            comboBox1.Items.Add(new ComboboxItem("ALL", "ALL"));
            comboBox1.Items.Add(new ComboboxItem("0000000027","CS"));
            comboBox1.Items.Add(new ComboboxItem("0000000031", "NK PD"));
            comboBox1.Items.Add(new ComboboxItem("0000000011", "PD1"));
            comboBox1.Items.Add(new ComboboxItem("0000000024", "PD2-EMS1"));
            comboBox1.Items.Add(new ComboboxItem("0000000023", "PD2-VPD"));
            comboBox1.Items.Add(new ComboboxItem("0000000022", "PD3"));
            comboBox1.Items.Add(new ComboboxItem("0000000021", "PD4"));
            comboBox1.Items.Add(new ComboboxItem("0000000025", "PD5-COM"));
            comboBox1.Items.Add(new ComboboxItem("0000000026", "PD5-SHD"));

            comboBox1.SelectedIndex = 0;

            //DateTime Setting
            //ProgressBar setting
            progressBar1.Value = 0;
            progressBar1.Maximum = 6;
            progressBar1.Step = 1;

            

            //MessageBox.Show(FromDate.Value.ToString("HH"));

        }

        private void InitialState()
        {
            progressBar1.Value = 0;

            //foreach (var process in Process.GetProcessesByName("NNNN"))
            //{
           //     process.Kill();
            //}

        }

        private class ComboboxItem
        {
            public ComboboxItem(string value, string text)
            {
                Value = value;
                Text = text;
            }
            public string Value
            {
                get;
                set;
            }

            public string Text
            {
                get;
                set;
            }

            public override string ToString()
            {
                return Text;
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ProdLine = (comboBox1.SelectedItem as ComboboxItem).Value;
            // MessageBox.Show(ProdLine);
           
        }

        private void StartBtn_Click(object sender, EventArgs e)
        {

           // var checkedBoxes = this.Controls.OfType<CheckBox>().Count(c => c.Checked);

            progressBar1.Value = 0;
            progressBar1.PerformStep();
            if (QAChk.Checked)
            {
                if (GetQAQty() != 0)
                {
                    InitialState();
                    return;
                }
            }
            progressBar1.PerformStep();

            if (ReChk.Checked)
            {
                if (GetRepairData() != 0)
                {
                    InitialState();
                    return;
                }

                progressBar1.PerformStep();


                if (GetExcel() != 0)
                {
                    InitialState();
                    return;
                }
                progressBar1.PerformStep();

            }
            else { progressBar1.Increment(2); }

            if (Top5Chk.Checked)
            {
                if (Sort() != 0)
                {
                    InitialState();
                    return;
                }
                progressBar1.PerformStep();

                if (MakeExcel() != 0)
                {
                    InitialState();
                    return;
                }
                progressBar1.PerformStep();
            }
            else { progressBar1.Increment(2); }

        }

        private void Top5Chk_CheckedChanged(object sender, EventArgs e)
        {
            if (Top5Chk.Checked == true)
            {
                ReChk.Checked = true;
            }
        }
        
    }
}
