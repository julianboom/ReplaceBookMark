﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReplaceBookMark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = GetFilePath();
        }
        private string GetFilePath()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                this.textBox2.AppendText("你所选择的材料文件夹目录是：" + foldPath + "\r\n");
                //MessageBox.Show("已选择文件夹:" + foldPath, "选择文件夹提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return foldPath;
            }
            return "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.textBox2.Clear();
        }
    }
}
