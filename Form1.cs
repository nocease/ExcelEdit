using System;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using System.Windows.Documents;
using System.Collections.Generic;

namespace ExcelEdit
{
    public partial class Excel处理 : Form
    {
        public Excel处理()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFolderPath = folderDialog.SelectedPath;
                    string[] fileNames = getFiles(selectedFolderPath);
                    if (fileNames != null)
                    {
                        foreach (string filename in fileNames)
                        {
                            updExcel(filename);
                        }
                        MessageBox.Show("执行完毕!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
            }

        }

        //获取文件夹内全部文件真实路径
        private string[] getFiles(string folde)
        {
            try
            {
                string folderPath = @folde;
                string[] filesName = Directory.GetFiles(folderPath);
                return filesName;
            }
            catch (Exception e)
            {
                MessageBox.Show("文件夹选择失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }

        //修改表格
        private void updExcel(string file)
        {
            try
            {
                //获取随机倍数
                List<double> numbers = new List<double> { 1.05, 1.06,1.07, 1.08, 1.09, 1.1, 1.11, 1.12 };
                Random random = new Random();
                int index = random.Next(numbers.Count);
                double selectedNumber = numbers[index];
                //修改表格
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var fileInfo = new FileInfo(@file);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    double sum = 0;
                    for (int i = 5; i <= 13; i++)
                    {
                        var cell = worksheet.Cells[$"G{i}"];
                        if (double.TryParse(cell.Text, out double value))
                        {
                            cell.Value = value * selectedNumber;
                            sum += cell.GetValue<double>();
                        }
                    }
                    if (sum != 0)
                    {
                        worksheet.Cells["G14"].Value = sum;
                    }
                    package.Save();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("有11111个文件处理失败："+ file, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void Excel处理_Load(object sender, EventArgs e)
        {

        }
    }
}
