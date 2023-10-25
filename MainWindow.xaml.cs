using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using MiniExcelLibs;
using Octokit;

namespace ProjectProgressExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            CheckUpdate();

            this.Title = title;
        }

        private readonly string title = "医学中心项目进度自动导出";
        private string path = ""; // 进度表文件路径
        private readonly Version localVersion = new Version("0.1.1"); // 本地版本号
        private Version? latestGithubVersion = null; // Github 最新版本号

        /// <summary>
        /// 导入医学中心进度工作簿按钮 点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnInputProgressTable_Click(object sender, RoutedEventArgs e)
        {
            // 打开进度表
            var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "xlsx 文件|*.xlsx"
            };
            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                path = openFileDialog.FileName;
            }
            else
            {
                return;
            }

            this.Title = title + " - " + path;
            ReadProgress(path);
        }


        /// <summary>
        /// 刷新按钮 点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(path))
            {
                BtnInputProgressTable_Click(sender, e); // 路径为空，直接调用 BtnInputProgressTable_Click 事件
            }
            else
            {
                ReadProgress(path);
            }
        }


        /// <summary>
        /// 更新按钮 点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {

            System.Diagnostics.Process.Start("explorer.exe", "https://github.com/Snoopy1866/ProjectProgressExport/releases/tag/" + latestGithubVersion);
        }


        /// <summary>
        /// 重试按钮 点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRetry_Click(object sender, RoutedEventArgs e)
        {
            CheckUpdate();
        }

        /// <summary>
        /// 读取两个工作表的进度
        /// </summary>
        /// <param name="path">进度表文件路径</param>
        private void ReadProgress(string path)
        {
            try
            {
                // 读取临床试验进度
                var clinicalSheetName = tbxExcelClinicalSheetName.Text;
                var clinicalDataTable = MiniExcel.QueryAsDataTable(path, useHeaderRow: true, sheetName: clinicalSheetName);

                var clinicalProgressTitle = txtExcelClinicalTitle.Text;
                var clinicalProjectColumnName = txtExcelClinicalProjectColumnName.Text;
                var clinicalMedicalColumnName = txtExcelClinicalMedicalColumnName.Text;
                var clinicalStatisticsColumnName = txtExcelClinicalStatisticsColumnName.Text;
                var clinicalDataManageColumnName = txtExcelClinicalDataManageColumnName.Text;

                var clinicalProgressTextDictionary = ReadProgressOfSingleSheet(clinicalDataTable, clinicalProgressTitle, clinicalProjectColumnName, clinicalMedicalColumnName, clinicalStatisticsColumnName, clinicalDataManageColumnName);

                // 读取CER进度
                var cerSheetName = tbxExcelCERSheetName.Text;
                var cerDataTable = MiniExcel.QueryAsDataTable(path, useHeaderRow: true, sheetName: cerSheetName);

                var cerProgressTitle = txtExcelCERTitle.Text;
                var cerProjectColumnName = txtExcelCERProjectColumnName.Text;
                var cerMedicalColumnName = txtExcelCERMedicalColumnName.Text;
                var cerStatisticsColumnName = txtExcelCERStatisticsColumnName.Text;

                var cerProgressTextDictionary = ReadProgressOfSingleSheet(cerDataTable, cerProgressTitle, cerProjectColumnName, cerMedicalColumnName, cerStatisticsColumnName);

                // 展示可复制到邮件的文本
                tbxProgressTextCopyToMail.Text = clinicalProgressTextDictionary["CopyToMail"].ToString() + cerProgressTextDictionary["CopyToMail"].ToString();
                tbxProgressTextCopyToMail.Foreground = Brushes.Black;

                // 展示可复制到工作表周进展的文本
                dgrClinicalProgress.ItemsSource = (List<ProgressInfo>)clinicalProgressTextDictionary["CopyToExcel"];
                dgrCERProgress.ItemsSource = (List<ProgressInfo>)cerProgressTextDictionary["CopyToExcel"];
            }
            catch (Exception err)
            {
                tbxProgressTextCopyToMail.Text = err.Message;
                tbxProgressTextCopyToMail.Foreground = Brushes.Red;
            }
        }

        /// <summary>
        /// 读取单个 sheet 的进度
        /// </summary>
        /// <param name="table">DataTable对象</param>
        /// <param name="title">段落标题</param>
        /// <param name="ExcelProjectColumnName">项目名称列名</param>
        /// <param name="ExcelMedicalColumnName">医学进度列名</param>
        /// <param name="ExcelStatisticsColumnName">统计进度列名</param>
        /// <param name="ExcelDataManageColumnName">数管进度列名</param>
        /// <returns></returns>
        private static Dictionary<string, object> ReadProgressOfSingleSheet(DataTable table, string title, string ExcelProjectColumnName, string ExcelMedicalColumnName, string ExcelStatisticsColumnName, string? ExcelDataManageColumnName = null)
        {
            // 获取项目名称、医学、数管、统计进度所在的列
            var dataColumnCollection = table.Rows[0].Table.Columns;
            var projectNameColumnIndex = dataColumnCollection.IndexOf(ExcelProjectColumnName);
            var mcColumnIndex = dataColumnCollection.IndexOf(ExcelMedicalColumnName);
            var stColumnIndex = dataColumnCollection.IndexOf(ExcelStatisticsColumnName);
            var dmColumnIndex = dataColumnCollection.IndexOf(ExcelDataManageColumnName);

            // 遍历所有行，提取进度文字
            var progressTextCopyToMail = "";
            var progressTextCopyToExcelList = new List<ProgressInfo>();
            var projectHasNewProgressIndex = 0;
            for (int i = 0; i <= table.Rows.Count - 1; i++)
            {
                var itemArray = table.Rows[i].ItemArray;
                var MCProgressExcelText = itemArray[mcColumnIndex]?.ToString()?.Replace("\n", "、").Trim();
                var STProgressExcelText = itemArray[stColumnIndex]?.ToString()?.Replace("\n", "、").Trim();

                // 读取CER项目进度时，索引 dmColumnIndex 为 -1，需要额外处理
                string? DMProgressExcelText = null;
                if (dmColumnIndex >= 0)
                {
                    DMProgressExcelText = itemArray[dmColumnIndex]?.ToString()?.Replace("\n", "、").Trim();
                }


                // 检查是否有新进度
                if (!string.IsNullOrEmpty(MCProgressExcelText) || !string.IsNullOrEmpty(STProgressExcelText) || !string.IsNullOrEmpty(DMProgressExcelText))
                {
                    var projectName = itemArray[projectNameColumnIndex]?.ToString()?.Trim();
                    var projectProgressText = "";
                    progressTextCopyToMail = progressTextCopyToMail + projectHasNewProgressIndex + ". " + projectName + "\n";

                    projectHasNewProgressIndex++;

                    var indentAsciiCode = 96;
                    // 检查医学进度
                    if (!string.IsNullOrEmpty(MCProgressExcelText))
                    {
                        indentAsciiCode += 1;
                        progressTextCopyToMail = progressTextCopyToMail + "    " + Convert.ToChar(indentAsciiCode) + ")    " + MCProgressExcelText + "\n";
                        projectProgressText = projectProgressText + MCProgressExcelText + "\n";
                    }

                    // 检查统计进度
                    if (!string.IsNullOrEmpty(STProgressExcelText))
                    {
                        indentAsciiCode += 1;
                        progressTextCopyToMail = progressTextCopyToMail + "    " + Convert.ToChar(indentAsciiCode) + ")    " + STProgressExcelText + "\n";
                        projectProgressText = projectProgressText + STProgressExcelText + "\n";
                    }

                    // 检查数管进度
                    if (!string.IsNullOrEmpty(DMProgressExcelText))
                    {
                        indentAsciiCode += 1;
                        progressTextCopyToMail = progressTextCopyToMail + "    " + Convert.ToChar(indentAsciiCode) + ")    " + DMProgressExcelText + "\n";
                        projectProgressText = projectProgressText + DMProgressExcelText + "\n";
                    }

                    progressTextCopyToExcelList.Add(new ProgressInfo(projectHasNewProgressIndex, projectName, projectProgressText.TrimEnd(new char[] { '\n' })));
                }
            }

            progressTextCopyToMail = title + "\n\n" + progressTextCopyToMail + "\n";
            var progressTextDictionary = new Dictionary<string, object>
            {
                ["CopyToMail"] = progressTextCopyToMail,
                ["CopyToExcel"] = progressTextCopyToExcelList,
            };

            return (progressTextDictionary);
        }

        /// <summary>
        /// 检查更新
        /// </summary>
        private async void CheckUpdate()
        {
            try
            {
                var client = new GitHubClient(new ProductHeaderValue("Snoopy1866"));
                IReadOnlyList<Release> release = await client.Repository.Release.GetAll("Snoopy1866", "ProjectProgressExport");
                latestGithubVersion = new Version(release[0].TagName);

                if (latestGithubVersion.CompareTo(localVersion) > 0)
                {
                    txbUpdateInfo.Text = "有更新可用，版本号：" + latestGithubVersion.ToString();
                    txbUpdateInfo.Foreground = Brushes.Red;

                    btnUpdate.Visibility = Visibility.Visible;
                    btnRetry.Visibility = Visibility.Collapsed;
                    
                }
                else
                {
                    txbUpdateInfo.Text = "已是最新版本。";
                    txbUpdateInfo.Foreground = Brushes.Black;

                    btnUpdate.Visibility = Visibility.Collapsed;
                    btnRetry.Visibility = Visibility.Collapsed;
                }
            }
            catch (Exception)
            {
                txbUpdateInfo.Text = "检查更新失败！";
                txbUpdateInfo.Foreground = Brushes.Red;

                btnUpdate.Visibility = Visibility.Collapsed;
                btnRetry.Visibility = Visibility.Visible;
            }
            stpUpdatePanel.Visibility = Visibility.Visible;
        }


        /// <summary>
        /// 超链接文本 点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            var hyperlink = (Hyperlink)sender;
            System.Diagnostics.Process.Start("explorer.exe", hyperlink.NavigateUri.AbsoluteUri);
        }
    }


    /// <summary>
    /// 项目进展信息类
    /// </summary>
    public class ProgressInfo
    {
        public int ID { get; set; }
        public string? ProjectName { get; set; }
        public string ProjectProgressText { get; set; }

        public ProgressInfo(int ID, string? ProjectName, string ProjectProgressText)
        {
            this.ID = ID;
            this.ProjectName = ProjectName;
            this.ProjectProgressText = ProjectProgressText;
        }
    }
}
