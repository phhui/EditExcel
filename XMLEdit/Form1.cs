using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace XMLEdit
{
    public partial class Form1 : Form
    {
        private Dictionary<string,string> d = new Dictionary<string,string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_search_Click(object sender, EventArgs e)
        {
            if (txt_search.Text.Length < 1) return;
            foreach (KeyValuePair<string, string> a in d)
            {
                if (a.Key == txt_search.Text)
                    txt_output.Text = a.Value;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string str = File.ReadAllText("df.txt", System.Text.Encoding.GetEncoding("gb2312"));
            string[] arr;
            string[] res = str.Replace("\r\n", "#").Split('#');
            int n = res.Length;
            for (int i = 0; i < n; i++)
            {
                arr = res[i].Replace("\t","$").Split('$');
                if(arr.Length>12&&!d.ContainsKey(arr[12]))d.Add(arr[12], res[i]);
            }
            ReadExcel("stuInfoCountry.xls", dataGridView1);
        }
        private void ReadExcel(string sExcelFile,DataGridView dg)
        {
            DataTable ExcelTable;
            DataSet ds = new DataSet();
            //Excel的连接
            OleDbConnection objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sExcelFile + ";" + "Extended Properties=Excel 8.0;");
            objConn.Open();
            DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            string tableName = schemaTable.Rows[0][2].ToString().Trim();//获取 Excel 的表名，默认值是sheet1
            string strSql = "select * from [" + tableName + "]";
            OleDbCommand objCmd = new OleDbCommand(strSql, objConn);
            OleDbDataAdapter myData = new OleDbDataAdapter(strSql, objConn);
            myData.Fill(ds, tableName);//填充数据
            objConn.Close();

            ExcelTable = ds.Tables[tableName];
            dg.DataSource = ExcelTable;
            int iColums = ExcelTable.Columns.Count;//列数
            int iRows = ExcelTable.Rows.Count;//行数

            //定义二维数组存储 Excel 表中读取的数据
            string[,] storedata = new string[iRows, iColums];

            for (int i = 0; i < ExcelTable.Rows.Count; i++)
                for (int j = 0; j < ExcelTable.Columns.Count; j++)
                {
                    //将Excel表中的数据存储到数组
                    storedata[i, j] = ExcelTable.Rows[i][j].ToString();

                }
            int excelBom = 0;//记录表中有用信息的行数，有用信息是指除去表的标题和表的栏目，本例中表的用用信息是从第三行开始
            //确定有用的行数
            for (int k = 2; k < ExcelTable.Rows.Count; k++)
                if (storedata[k, 1] != "")
                    excelBom++;
            if (excelBom == 0)
            {
                MessageBox.Show("<script language=javascript>alert('您导入的表格不合格式！')</script>");
            }
            else
            {
                //LoadDataToDataBase(storedata，excelBom)//该函数主要负责将 storedata 中有用的数据写入到数据库中，在此不是问题的关键省略 
            }
            txt_output.Text = storedata[1,1];
            ExportDataGridViewToExcel(dg);
        }
        public static void ExportDataGridViewToExcel(DataGridView dataGridview1)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Execl  files  (*.xls)|*.xls";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "导出Excel文件到";

            DateTime now = DateTime.Now;
            saveFileDialog.FileName = now.Year.ToString().PadLeft(2)
            + now.Month.ToString().PadLeft(2, '0')
            + now.Day.ToString().PadLeft(2, '0') + "-"
            + now.Hour.ToString().PadLeft(2, '0')
            + now.Minute.ToString().PadLeft(2, '0')
            + now.Second.ToString().PadLeft(2, '0');

            saveFileDialog.ShowDialog();

            Stream myStream;
            myStream = saveFileDialog.OpenFile();
            StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding("gb2312"));
            string str = "";
            try
            {
                //写标题    
                for (int i = 0; i < dataGridview1.ColumnCount; i++)
                {
                    if (i > 0)
                    {
                        str += "\t";
                    }
                    str += dataGridview1.Columns[i].HeaderText;
                }

                sw.WriteLine(str);
                //写内容  
                for (int j = 0; j < dataGridview1.Rows.Count; j++)
                {
                    string tempStr = "";
                    for (int k = 0; k < dataGridview1.Columns.Count; k++)
                    {
                        if (k > 0)
                        {
                            tempStr += "\t";
                        }
                        tempStr += dataGridview1.Rows[j].Cells[k].Value.ToString();
                    }
                    sw.WriteLine(tempStr);
                }
                sw.Close();
                myStream.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                sw.Close();
                myStream.Close();
            }
        }  
    }
}
