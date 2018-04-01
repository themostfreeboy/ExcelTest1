﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace excel_test
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new Form1());
            }
            catch(Exception ex)
            {
                MessageBox.Show("程序出错：\n" + ex.Message + "\n程序退出");
            }
        }
    }
}