using System;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using System.Windows.Documents;
using System.Collections.Generic;
using System.IO.Packaging;

namespace ExcelEdit
{
    public partial class Excel处理 : Form
    {
        public Excel处理()
        {
            // 设置LicenseContext为非商业使用
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "所有文件|*.*"; // 可以定义需要的文件类型过滤器
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;//获取文件路径
                    FileInfo fileInfo = new FileInfo(selectedFilePath);
                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // 假设数据在第一个工作表中
                        List<PersonData> personList = new List<PersonData>();
                        // 从Excel中读取数据并存储到集合中
                        for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var person = new PersonData
                            {
                                SerialNumber = worksheet.Cells[row, 1].Text,
                                Township = worksheet.Cells[row, 2].Text,
                                Village = worksheet.Cells[row, 3].Text,
                                Name = worksheet.Cells[row, 4].Text,
                                IDNumber = worksheet.Cells[row, 5].Text,
                                FamilyPopulation = worksheet.Cells[row, 6].Text,
                                Income2022 = worksheet.Cells[row, 7].Text
                            };
                            personList.Add(person);
                        }
                        MessageBox.Show("读取文件成功,请选择要输出的收入测算表文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        OpenFileDialog openFileDialog2 = new OpenFileDialog();
                        openFileDialog.Filter = "所有文件|*.*"; // 可以定义需要的文件类型过滤器
                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string selectedFilePath2 = openFileDialog.FileName;//获取文件路径
                            FileInfo fileInfo2 = new FileInfo(selectedFilePath2);
                            // 输出数据
                            foreach (var person in personList)
                            {

                                using (ExcelPackage package2 = new ExcelPackage(fileInfo2))
                                {
                                    ExcelWorksheet copiedWorksheet = package2.Workbook.Worksheets.Add(person.Name, package2.Workbook.Worksheets[0]);
                                    copiedWorksheet.Cells["A1"].Value = person.Village + "2023年度收入测算表";
                                    copiedWorksheet.Cells["A2"].Value = "户主姓名：" + person.Name;
                                    copiedWorksheet.Cells["C2"].Value = "年度家庭人口：" + person.FamilyPopulation;
                                    copiedWorksheet.Cells["C19"].Value = person.Income2022;

                                    // 保存Excel文件 
                                    package2.Save();
                                }
                            }
                            MessageBox.Show("运行结束。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("失败。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

            //修改表格
            private void updExcel(string file)
            {
                try
                {
                    //获取随机倍数
                    List<double> numbers = new List<double> { 1.05, 1.06, 1.07, 1.08, 1.09, 1.1, 1.11, 1.12 };
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
                    MessageBox.Show("有一个文件处理失败：" + file, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    public class PersonData
    {
        public string SerialNumber { get; set; }
        public string Township { get; set; }
        public string Village { get; set; }
        public string Name { get; set; }
        public string IDNumber { get; set; }
        public string FamilyPopulation { get; set; }
        public string Income2022 { get; set; }
    }

}
