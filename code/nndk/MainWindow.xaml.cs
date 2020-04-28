using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Permissions;
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
using MahApps.Metro.Controls;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
            //读excel
            //循环处理
            StringBuilder sbr = new StringBuilder();
            using (FileStream fs = File.OpenRead(srcFilePath.Text))   //打开myxls.xls文件
            {
                XSSFWorkbook wk = new XSSFWorkbook(fs);   //把xls文件中的数据写入wk中
                for (int i = 0; i < wk.NumberOfSheets; i++)  //NumberOfSheets是myxls.xls中总共的表数
                {
                    ISheet sheet = wk.GetSheetAt(i);   //读取当前表数据
                    for (int j = 0; j <= sheet.LastRowNum; j++)  //LastRowNum 是当前表的总行数
                    {
                        IRow row = sheet.GetRow(j);  //读取当前行数据
                        if (row != null)
                        {
                            sbr.Append("-------------------------------------\r\n"); //读取行与行之间的提示界限
                            for (int k = 0; k <= row.LastCellNum; k++)  //LastCellNum 是当前行的总列数
                            {
                                ICell cell = row.GetCell(k);  //当前表格
                                if (cell != null)
                                {
                                    sbr.Append(cell.ToString());   //获取表格中的数据并转换为字符串类型
                                }
                            }
                        }
                    }
                }
            }
            sbr.ToString();
            using (StreamWriter wr = new StreamWriter(new FileStream(@"E:/myText.txt", FileMode.Append)))  //把读取xls文件的数据写入myText.txt文件中
            {
                wr.Write(sbr.ToString());
                wr.Flush();
            }
            isExcuting.IsActive = false;
            this.IsEnabled = true;
        }
    }
}
