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

                        for (int j = 0; j <= sheet.LastRowNum; j++)
                        {
                            IRow row = sheet.GetRow(j);

                            if (row != null)
                            {
                                ICell cell = row.GetCell(1);

                                if (cell == null)
                                {
                                    MessageBox.Show("[null]");
                                }
                                else
                                {
                                    MessageBox.Show(cell.ToString());
                                }
                            }
                        }

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
