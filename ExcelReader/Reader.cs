using log4net;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public partial class Reader : Form
    {

        public ILog log;
        public string Log4ConfigFilePath { get { return AppDomain.CurrentDomain.BaseDirectory + @"\\" + ConfigurationManager.AppSettings["Log4ConfigFileName"]; } }

        public Reader()
        {
            InitializeComponent();

            checkBox1.Checked = true;

            //掃描欄位的起始值初始
            numericUpDown1.Value = 4;
            numericUpDown1.Maximum = 999;
            numericUpDown1.Minimum = 1;

            //姓名欄位始值位置
            numericUpDown2.Value = 6;
            numericUpDown2.Maximum = 99;
            numericUpDown2.Minimum = 1;

            //點數欄位始值位置
            numericUpDown3.Value = 22;
            numericUpDown3.Maximum = 99;
            numericUpDown3.Minimum = 1;

            //獎金欄位始值位置
            numericUpDown4.Value = 24;
            numericUpDown4.Maximum = 99;
            numericUpDown4.Minimum = 1;

            //幣別欄位始值位置
            numericUpDown5.Value = 23;
            numericUpDown5.Maximum = 99;
            numericUpDown5.Minimum = 1;

            //進度顯示值的初始
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 10;
            progressBar1.Step = 1;

            label2.Text = "進度:";
        }

        //選擇檔案來源後的結果
        private void button1_Click(object sender, EventArgs e)
        {
            //只可選擇Excel相關的檔案
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FileName = "excelFile";

            //對話框OK後，將所選檔案的路徑寫入textBox1
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {             
                string directoryPath = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                this.textBox1.Text = directoryPath;
            }
        }

        //開始解析
        private void button2_Click(object sender, EventArgs e)
        {
            
            try
            {
            
                string path = textBox1.Text;
                if (string.IsNullOrEmpty(path) || path.Length <= 0)
                {
                    MessageBox.Show("未選擇上傳路徑。");
                    return;
                }
                else
                {
                    string fileName = Path.GetFileName(path); //把檔名存起來

                    #region ===驗證副檔名===
                    string ext = Path.GetExtension(path);
                    if (ext.Equals(".xls") || ext.Equals(".xls") || ext.Equals(".xlsm"))
                    {
                        //通過
                        progressBar1.PerformStep(); //進度值+1
                        label2.Text = "10%";
                    }
                    else
                    {
                        MessageBox.Show("副檔名" + ext + "不正確! 本程式僅接受開啟Excel相關的檔案。");
                        return;
                    }
                    #endregion

                    #region ===驗證檔案路徑是否存在===
                    FileInfo fi = new FileInfo(path);
                    //如果驗證的檔案路徑有問題，就跳警告阻擋。
                    if (!fi.Exists)
                    {
                        MessageBox.Show(path + "\n" + "路徑不正確，找不到指定的檔案! 請檢查是否輸入有誤。");
                        return;
                    }
                    else
                    {
                        progressBar1.PerformStep(); //進度值+1
                        label2.Text = "20%";
                    }
                    #endregion

                    #region ===宣告要寫出的Excel資料=====
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook workbook = excel.Workbooks.Open(path);
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1); //讀取第一個活頁簿的內容

                    //int sheetRowsCount = sheet.UsedRange.Rows.Count;
                    Excel.Range range = sheet.UsedRange; //找到已寫入的範圍

                    int rw = range.Rows.Count; //已被寫入的row
                    int cl = range.Columns.Count; //已被寫入的column
                    //MessageBox.Show(fileName+"已經被寫了: "+ rw + "row、 "+ cl+"column");

                    //宣告DataTable，欄與列
                    System.Data.DataTable table = new DataTable("ParentTable");
                    DataColumn column;
                    DataRow row;

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "Name";
                    column.AutoIncrement = false;
                    column.Caption = "Name";
                    column.ReadOnly = false;
                    column.Unique = false;
                    // Add the column to the table.
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.Double");
                    column.ColumnName = "Points";
                    column.AutoIncrement = false;
                    column.Caption = "Points";
                    column.ReadOnly = false;
                    column.Unique = false;
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "PriceTitle";
                    column.AutoIncrement = false;
                    column.Caption = "PriceTitle";
                    column.ReadOnly = false;
                    column.Unique = false;
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.Int32");
                    column.ColumnName = "Price";
                    column.AutoIncrement = false;
                    column.Caption = "Price";
                    column.ReadOnly = false;
                    column.Unique = false;
                    table.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "Currency";
                    column.AutoIncrement = false;
                    column.Caption = "Currency";
                    column.ReadOnly = false;
                    column.Unique = false;
                    table.Columns.Add(column);
                    #endregion

                    progressBar1.Increment(2);
                    label2.Text = "40%";

                    //取得掃描欄位的值
                    decimal val = numericUpDown1.Value;
                    int beginRow = Decimal.ToInt32(val);

                    //姓名欄位的起點位置
                    decimal val2 = numericUpDown2.Value;
                    int beginCol1 = Decimal.ToInt32(val2);

                    //點數欄位的起點位置
                    decimal val3 = numericUpDown3.Value;
                    int beginCol2 = Decimal.ToInt32(val3);

                    //獎金欄位的起點位置
                    decimal val4 = numericUpDown4.Value;
                    int beginCol3 = Decimal.ToInt32(val4);

                    //幣別欄位的起點位置
                    decimal val5 = numericUpDown5.Value;
                    int beginCol4 = Decimal.ToInt32(val5);


                    for (int i = 1; i <= rw; i++)
                    {
                        //beginRow以前的行數，都直接跳過
                        if (i < beginRow)
                        {
                            continue;
                        }

                        for (int j = 1; j <= cl; j++)
                        {

                            if (j == 6)
                            {
                                string str = (string)(range.Cells[i, beginCol1] as Excel.Range).Value2;

                                string currency = (string)(range.Cells[i, beginCol4] as Excel.Range).Value2;

                                double number;
                                int price = 0;

                                try
                                {

                                    number = double.Parse((range.Cells[i, beginCol2] as Excel.Range).Value2.ToString());
                                    price = Int32.Parse((range.Cells[i, beginCol3] as Excel.Range).Value2.ToString());

                                    //如果price是負數，就自動設回0。
                                    price = price < 0 ? 0 : price;

                                }
                                catch (Exception ex)
                                {
                                    //System.Diagnostics.Debug.WriteLine(str+"的數字轉換出現錯誤~" + ex);
                                    continue;
                                }


                                //如果是字串內容是空的或叫"姓名"，就直接跳下一筆。
                                if (string.IsNullOrEmpty(str) || str.Equals("姓名"))
                                {
                                    continue;
                                }
                                else
                                {
                                    //System.Diagnostics.Debug.WriteLine("第" + i + "行: " + str + "最終點數: "+ number+ " 金額:"+ price);

                                    row = table.NewRow();
                                    row["Name"] = str;
                                    row["Points"] = number;
                                    row["Price"] = price;
                                    row["Currency"] = currency;
                                    table.Rows.Add(row);

                                }
                            }
                        }
                    }

                    progressBar1.Increment(2);
                    label2.Text = "60%";

                    //要做Distinct，當然要設成true，其他參數是要做Group By的欄位名稱
                    DataTable dtGroup = table.DefaultView.ToTable(true, "Name");

                    //開始加欄位
                    dtGroup.Columns.Add("CountColumn");
                    dtGroup.Columns.Add("SumColumn");
                    dtGroup.Columns.Add("SumColumn2");
                    //dtGroup.Columns.Add("AvgColumn");

                    for (int i = 0; i < dtGroup.Rows.Count; i++)
                    {
                        //取資料，用String是因為上方加欄位時，沒指定型別為數字
                        string strCount = table.Select("Name='" + dtGroup.Rows[i]["Name"].ToString() + "'").Length.ToString();
                        string strSum = table.Compute("SUM(Points)", "Name='" + dtGroup.Rows[i]["Name"].ToString() + "'").ToString();
                        string strSum2 = table.Compute("SUM(Price)", "Name='" + dtGroup.Rows[i]["Name"].ToString() + "'").ToString();
                        //string strAvg = table.Compute("AVG(Points)", "Name='" + dtGroup.Rows[i]["Name"].ToString() + "'").ToString();

                        //設定資料
                        dtGroup.Rows[i]["CountColumn"] = (strCount == "" ? "0" : strCount);
                        dtGroup.Rows[i]["SumColumn"] = (strSum == "" ? "0" : strSum);
                        dtGroup.Rows[i]["SumColumn2"] = (strSum == "" ? "0" : strSum2);
                        //dtGroup.Rows[i]["AvgColumn"] = (strAvg == "" ? "0" : strAvg);
                    }

                    //印table的內容
                    //foreach (DataRow r in table.Rows){System.Diagnostics.Debug.WriteLine(r["Name"].ToString() +" "+ r["Points"].ToString());}

                    progressBar1.Increment(2);
                    label2.Text = "80%";

                    #region ===把所有結果寫成Excel======
                    Excel.Application App = new Excel.Application();
                    Excel.Workbook book = App.Workbooks.Add();

                    // Excel WorkBook，預設會產生一個 WorkSheet，索引從 1 開始，而非 0
                    Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets.Item[1];

                    //設置禁止彈出保存和覆蓋的詢問提示框
                    Sheet.Application.DisplayAlerts = false;
                    Sheet.Application.AlertBeforeOverwriting = false;

                    int drawRow = 1;

                    Sheet.Cells[1, 1] = "姓名:";
                    Sheet.Cells[1, 2] = "提案筆數:";
                    Sheet.Cells[1, 3] = "總點數:";
                    Sheet.Cells[1, 4] = "總獎金:";
                    Sheet.Cells[1, 5] = "幣別:";
                    drawRow++;

                    foreach (DataRow r in dtGroup.Rows)
                    {
                        #region ===透過人名，回頭到table找到對應名稱的第一筆資料所對應的Currency 
                        string name = r["Name"].ToString();
   
                        DataRow[] foundRow = table.Select("Name = '"+name+"'");

                        string currency = null;

                        if (foundRow.Length > 0)
                        {
                            int SelectedIndex =table.Rows.IndexOf(foundRow[0]);
                            currency = table.Rows[SelectedIndex]["Currency"].ToString();
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("透過Name沒有回頭找到對應的row。");
                        }
                        #endregion

                    //System.Diagnostics.Debug.WriteLine(r["Name"].ToString() + " " + r["CountColumn"].ToString() + " " + r["SumColumn"].ToString() + " " + r["SumColumn2"].ToString() + " "+ currency);
                    Sheet.Cells[drawRow, 1] = r["Name"].ToString();
                        Sheet.Cells[drawRow, 2] = r["CountColumn"].ToString();
                        Sheet.Cells[drawRow, 3] = r["SumColumn"].ToString();
                        Sheet.Cells[drawRow, 4] = r["SumColumn2"].ToString();
                        Sheet.Cells[drawRow, 5] = currency;
                        drawRow++;
                    }
                    #endregion

                    //寬度調整
                    Sheet.UsedRange.EntireColumn.AutoFit();

                    fileName = System.IO.Path.GetFileNameWithoutExtension(fileName); //移除副檔名

                    //預設存在資料夾路徑
                    string exportPath = Environment.CurrentDirectory + @"\" + fileName + "整合結果.xlsx";

                    //如果匯出至桌面被開啟，就改存到桌面。
                    if (checkBox1.Checked)
                    {
                        path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\" + fileName + "整合結果.xlsx";
                    }

                    progressBar1.Increment(2);
                    label2.Text = "100%";
                    MessageBox.Show("已解析完成!");


                    book.SaveAs(path); //儲存檔案
                    book.Close();      //關閉EXCEL
                    App.Quit();        //離開應用程式
                }
            
            }
            catch (Exception ex)
            {
                //log.Info("程式出錯: " + ex);
                MessageBox.Show("解析時出現錯誤!");
            }
            finally {
                progressBar1.Value = 0;
                label2.Text = "進度:";
            }
            
        }

    }
}
