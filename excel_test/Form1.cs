using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel;//excel需要

namespace excel_test
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)//窗体载入时处理函数
        {
            
        }

        private void btnTest1_Click(object sender, EventArgs e)
        {
            Excel.Application m_excel;
            Excel.Workbook m_workbook;

            m_excel = new Excel.Application();//对象实例化
            m_excel.Application.Workbooks.Add(true);

            int col;
            for (col = 0; col < 10; col++)
            {
                m_excel.Cells[1, col + 1] = col;
            }

            m_excel.Visible = true;//显示Excel内容
        }

        private void btnTest2_Click(object sender, EventArgs e)
        {
            Excel.Application m_excel;
            Excel.Workbook m_workbook;

            m_excel = new Excel.Application();//对象实例化

            m_workbook = m_excel.Workbooks.Open("D:\\test.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //m_workbook = m_excel.Workbooks.Open("D:\\test.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            m_excel.Visible = true;//显示Excel内容(为false时，excel.exe进程在后台运行，未找到关闭方法)

            Range rng;
            object obj;
            String str;
            rng = (Excel.Range)m_excel.Cells[2, 2];
            obj = rng.Value2;
            System.Diagnostics.Debug.WriteLine(obj.ToString());//输出调试信息
            str = rng.NumberFormatLocal;
            System.Diagnostics.Debug.WriteLine(str);//输出调试信息
            this.txtTest2.Text = obj.ToString();

            m_excel.DisplayAlerts = false; //设置禁止弹出保存和覆盖的询问提示框
            m_excel.AlertBeforeOverwriting = false;
            m_excel.Workbooks.Close();
            //object oV = System.Reflection.Missing.Value; //反复用到
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
           //System.Runtime.InteropServices.Marshal.ReleaseComObject(m_excel);
            //GC.Collect();
            m_excel.Quit();
        }

        public static System.Data.DataTable LoadDataFromExcel(string filePath)
        {
            Excel.Application m_excel;
            Excel.Workbook m_workbook;
            Excel.Worksheet worksheet;

            m_excel = new Excel.Application();//对象实例化

            m_workbook = m_excel.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            worksheet = (Worksheet)m_workbook.Worksheets[1];

            m_excel.Visible = true;//显示Excel内容(为false时，excel.exe进程在后台运行，未找到关闭方法)

            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;

            Excel.Range range1;

            System.Data.DataTable dt = new System.Data.DataTable();

            for (int i = 0; i < colCount; i++)
            {
                    //range1 = worksheet.get_Range(worksheet.Cells[1, i + 1], worksheet.Cells[1, i + 1]);
                    range1 = (Excel.Range)m_excel.Cells[1, i+1];
                    dt.Columns.Add(range1.Value2.ToString());
             }
             for (int j = 1; j < rowCount; j++)
             {
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < colCount; i++)
                    {
                        //range1 = worksheet.get_Range(worksheet.Cells[j + 1, i + 1], worksheet.Cells[j + 1, i + 1]);
                        range1 = (Excel.Range)m_excel.Cells[j+1, i + 1];
                        dr[i] = range1.Value2.ToString();
                    }
                    dt.Rows.Add(dr);
              }
              m_excel.DisplayAlerts = false; //设置禁止弹出保存和覆盖的询问提示框
              m_excel.AlertBeforeOverwriting = false;
              m_excel.Workbooks.Close();
              m_excel.Quit();
              return dt;
        }

        public static bool SaveDataTableToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Excel.Application app = new Excel.Application();//对象实例

            try
            {
                app.Visible = true;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int col = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelTable.Rows[i][j].ToString();
                            wSheet.Cells[i + 2, j + 1] = str;
                        }
                    }
                }

                int size = excelTable.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
                }

                app.DisplayAlerts = false; //设置禁止弹出保存和覆盖的询问提示框
                app.AlertBeforeOverwriting = false;

                wBook.SaveAs(filePath);//保存工作簿(当不指定FileFormat时，保存为默认格式，新建的文档默认格式为xlsx，即使filePath后缀为xls，内部数据格式仍为xlsx)
                wBook.Close();//关闭工作簿

                //app.Save(filePath);//保存excel文件
                //app.SaveWorkspace(filePath);
                app.Quit();
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }

        private void btnTest3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = LoadDataFromExcel("D:\\test.xls");
            //System.Data.DataTable dt = LoadDataFromExcel("D:\\test.xlsx");
            dgvTest3.AutoGenerateColumns = true;
            dgvTest3.DataSource = dt;
        }

        private void btnTest4_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetDgvToTable(dgvTest3);
            SaveDataTableToExcel(dt, "D:\\testsave.xlsx");
            //SaveDataTableToExcel(dt, "D:\\testsave.xls");
        }

        public System.Data.DataTable GetDgvToTable(DataGridView dgv)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
