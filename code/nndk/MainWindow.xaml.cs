using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading;
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
using MahApps.Metro.Controls;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Configuration;

namespace nndk
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            isExcuting.IsActive = false;
        }

        private void chooseFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "(*.xlsx)|*.xlsx|(*.xls)|*.xls";
            if (dlg.ShowDialog() == true)
            {
                srcFilePath.Text= dlg.FileName;
            }
        }

        private void startExcute(object sender, RoutedEventArgs e)
        {
            this.IsEnabled = false;
            isExcuting.IsActive = true;

            if (string.IsNullOrEmpty(srcFilePath.Text))
            {
                MessageBox.Show("请先选择一个文件");
                isExcuting.IsActive = false;
                this.IsEnabled = true;
                return;
            }
            string path = srcFilePath.Text;
            
            new Thread(() =>
            {
                List<Record> resData = ReadExcel(path);
                List<List<Record>> lstlst = new List<List<Record>>();
                if (resData != null)
                {
                    foreach (var code in Common.GetCodes().Split(',').ToList())
                    {
                        lstlst.Add(resData.Where(t => t.code == code).ToList());
                    }
                }
                //lstlst 数据源
                CalcNormal(lstlst);
                //周内
                //周末
                //加班延点
                //节假日


                this.Dispatcher.Invoke(() =>
                {
                    isExcuting.IsActive = false;
                    this.IsEnabled = true;
                });

            }).Start();
        }
        //计算正常工时
        public static float CalcNormal(List<List<Record>> lstlst)
        {
            float sum = 0;
            foreach (var lstPersonData in lstlst)    //遍历每个人
            {
                foreach (var record in lstPersonData)
                {

                }
            }
            return sum;
        }
        //计算加班工时
        public static void CalcOverTime(List<List<Record>> lstlst)
        {
            foreach (var lstPersonData in lstlst)    //遍历每个人
            {
                foreach (var record in lstPersonData)
                {

                }
            }
        }
        public static List<Record> ReadExcel(string filePaht)
        {
            try
            {
                FileStream fs = File.OpenRead(filePaht);

                if (filePaht.Contains("xlsx"))
                    return ReadWorkBook(new XSSFWorkbook(fs));//xlsx
                else
                    return ReadWorkBook(new HSSFWorkbook(fs));
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("正由另一进程使用"))
                {
                    MessageBox.Show("请关闭打开的文件");
                }
                return null;
            }
        }
        public static List<Record> ReadWorkBook(IWorkbook wk)
        {
            List<Record> lstData = new List<Record>();
            ISheet sheet = wk.GetSheetAt(0);    //获取sheet
            for (int j = 0; j <= sheet.LastRowNum; j++)  //LastRowNum 是当前表的总行数
            {
                IRow row = sheet.GetRow(j);  //读取当前行数据
                if (row != null)
                {
                    Record tperson = new Record();
                    tperson.code = row.GetCell(1)?.StringCellValue;
                    tperson.name = row.GetCell(2)?.StringCellValue;
                    tperson.time = row.GetCell(3)?.StringCellValue.ToDateTime();
                    if (!string.IsNullOrEmpty(tperson.code) && !string.IsNullOrEmpty(tperson.name) && tperson.time != null)
                    {
                        lstData.Add(tperson);
                    }
                }
            }
            return lstData;
        }
    }
}
