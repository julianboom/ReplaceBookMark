using DriverHelper;
using System;
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
        public MySqlHelper mySqlHelper;
        public Form1()
        {
            InitializeComponent();
            mySqlHelper = new MySqlHelper();
            InitialMapSelection();
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
                this.textBox2.AppendText("你所选择的文书模板目录是：" + foldPath + "\r\n");
                //MessageBox.Show("已选择文件夹:" + foldPath, "选择文件夹提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return foldPath;
            }
            return "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.textBox2.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.textBox1.Text))
            {
                this.textBox2.AppendText("请先选择文书模板所在的文件夹位置\r\n");
                return;
            }
            if (string.IsNullOrWhiteSpace(this.MapSelection.SelectedValue.ToString()))
            {
                this.textBox2.AppendText("请先选择映射关系\r\n");
                return;
            }
            this.textBox2.AppendText("正在进行批量替换............\r\n");
        }
        /// <summary>
        /// 初始化映射关系选择下拉框
        /// </summary>
        public void InitialMapSelection()
        {
            try
            {
                DataTable dt = mySqlHelper.ExecuteDataTable("SELECT `GroupID`,`GroupName` FROM bookmarkmap GROUP BY GroupName Order By CreateTime Desc", null);
                DataRow defaultRow = dt.NewRow();
                defaultRow["GroupID"] = "";
                defaultRow["GroupName"] = "请选择";
                dt.Rows.InsertAt(defaultRow, 0);
                //this.MapSelection.Items.AddRange
                this.MapSelection.DataSource = dt;
                this.MapSelection.ValueMember = "GroupID";
                this.MapSelection.DisplayMember = "GroupName";

            }
            catch (Exception e)
            {

                this.textBox2.AppendText("连接映射关系数据源失败，请检查数据库：" + e.ToString() + "\r\n");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            InitialMapSelection();
        }

        private void MapSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            string groupID = this.MapSelection.SelectedValue.ToString();
            string groupName = this.MapSelection.Text;
            if (string.IsNullOrWhiteSpace(groupID))
            {
                this.textBox2.AppendText("您未选择分组，请选择分组\r\n");
            }
            else if (groupID != "System.Data.DataRowView")
            {
                this.textBox2.AppendText($"您选择的分组ID为  {groupID}，分组名为  {groupName}\r\n");
            }
        }
    }
}
