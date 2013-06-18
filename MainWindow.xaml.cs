using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.ComponentModel;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Schedule_NPOI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BackgroundWorker backgroundWorker;

        public MainWindow()
        {
            InitializeComponent();

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            //backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            // backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);

        }

        private void btnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            if (backgroundWorker.IsBusy != true)
            {
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog();
            CommonFileDialogFilter filter = new CommonFileDialogFilter("*.xls文件", ".xls");
            openFileDialog.Filters.Add(filter);
            CommonFileDialogResult commonFileDialogResult = CommonFileDialogResult.None;
            App.Current.Dispatcher.Invoke(new Action(() =>
            {
                commonFileDialogResult = openFileDialog.ShowDialog();
            }));

            if (commonFileDialogResult == CommonFileDialogResult.Ok)
            {

                using (FileStream fileStream = File.OpenRead(@openFileDialog.FileName))
                {
                    IWorkbook workbook;


                    try
                    {
                        workbook = WorkbookFactory.Create(fileStream);

                        ISheet sheet = workbook.GetSheetAt(0);

                        App.Current.Dispatcher.Invoke(new Action(() =>
                        {
                            txtTextBox.Text += "====================\n" +
                                       "开始对文件 " + openFileDialog.FileName + "进行处理\n" +
                                       "====================\n";
                            txtTextBox.ScrollToEnd();
                        }));

                        IRow row;
                        ICell cell;
                        string strTeacherName = "";
                        string gradeNum = "";
                        string strClassList = "";

                        //处理老师列和年级更的空行，跳过第一行
                        for (int rowI = 1; rowI < sheet.LastRowNum; rowI++)
                        {
                            row = sheet.GetRow(rowI);


                            if (row != null)
                            {
                                //处理老师一行的空行
                                cell = row.GetCell(1);

                                if (cell.ToString() != "")
                                {
                                    strTeacherName = cell.ToString();

                                    App.Current.Dispatcher.Invoke(new Action(() =>
                                    {
                                        txtTextBox.Text += "正在对 " + strTeacherName + " 的信息进行处理\n";
                                        txtTextBox.ScrollToEnd();
                                    }));
                                }
                                else
                                {
                                    cell.SetCellValue(strTeacherName);
                                }

                                //处理年级一列的空行
                                cell = row.GetCell(2);

                                if (cell.ToString() != "")
                                {
                                    gradeNum = cell.ToString();
                                    gradeNum = gradeNum.Replace("小", "");
                                    gradeNum = gradeNum.Replace("级", "");
                                    cell.SetCellValue(gradeNum);
                                }
                                else
                                {
                                    cell.SetCellValue(gradeNum);
                                }

                                //改变任课班级一列的格式
                                cell = row.GetCell(5);

                                //第一行都有内容，不用判断是否为空
                                strClassList = cell.ToString();
                                strClassList = strClassList.Replace("小(", "");
                                strClassList = strClassList.Replace(")班", "");

                                cell.SetCellValue(strClassList);
                            }
                        }

                        //把一位老师在一门课中的多个班级的分成一个班级一行
                        int timesPerWeek = 0;
                        string strTempTeacher = "";
                        string[] classArray;
                        int classArrayLength;
                        string[] stringSeparators = new string[] { "、" };

                        for (int rowI = 1; rowI < sheet.LastRowNum; rowI++)
                        {
                            row = sheet.GetRow(rowI);
                            strClassList = row.GetCell(5).ToString();

                            classArray = strClassList.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                            classArrayLength = classArray.Length;

                            if (classArrayLength > 1)
                            {
                                //判断是否完成了一个人的信息处理
                                if (strTempTeacher != row.GetCell(1).ToString())
                                {
                                    strTempTeacher = row.GetCell(1).ToString();
                                }

                                App.Current.Dispatcher.Invoke(new Action(() =>
                                {
                                    txtTextBox.Text += "正在处理 " + strTempTeacher + " 的任课班级信息\n";
                                    txtTextBox.ScrollToEnd();
                                }));
                            }
                        }








                        CommonOpenFileDialog saveFileInFolderDialog = new CommonOpenFileDialog();
                        saveFileInFolderDialog.IsFolderPicker = true;
                        saveFileInFolderDialog.Title = "选择文件保存目录";

                        App.Current.Dispatcher.Invoke(new Action(() =>
                        {
                            commonFileDialogResult = saveFileInFolderDialog.ShowDialog();
                        }));


                        // DateTime dateTime = System.DateTime.Now;

                        if (commonFileDialogResult == CommonFileDialogResult.Ok)
                        {
                            FileStream fsSaveFile = File.Create(saveFileInFolderDialog.FileName + "教学计划" + DateTime.Now.ToString("yyyyMMddHHmmss")
                                                                                                             + ".xls");
                            workbook.Write(fsSaveFile);

                            fsSaveFile.Close();
                        }


                        /* IRow row;
                         ICell cell;
                         StringBuilder strTempString = new StringBuilder();
                        // string strCellString = "";

                         for (int rowI = 0; rowI <= sheet.LastRowNum; rowI++)
                         {
                             row = sheet.GetRow(rowI);

                             strTempString.Clear();

                             if(row != null)
                             {

                                 for (int cellI = 0; cellI <= row.LastCellNum; cellI++)
                                 {
                                     //if (row != null)
                                     //{
                                     cell = row.GetCell(cellI);

                                     if (cell == null)
                                     {
                                         strTempString.Append("[null]");
                                     }
                                     else
                                     {
                                         strTempString.Append(row.GetCell(cellI).ToString());
                                     }
                                 }
                             }

                             MessageBox.Show(strTempString.ToString());

                         } */

                    }
                    catch (InvalidOperationException ex)
                    {
                        if (ex.Message == "Unexpected record type (DefaultRowHeightRecord)")
                        {
                            MessageBox.Show("您打开的xls文件有内部错误\n错误信息为：" + ex.Message + "\n请使用Excel将此文件另存为之后再次尝试");
                        }

                    }



                    // MessageBox.Show(workbook.NumberOfSheets.ToString());

                }
            }
        }
    }
}
