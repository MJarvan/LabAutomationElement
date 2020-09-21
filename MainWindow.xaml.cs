using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.Windows.Controls.Primitives;
using NPOI.SS.Util;
using System.Diagnostics;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;

namespace LabAutomationElement
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 检测化合物名称合计
        /// </summary>
        List<KeyValuePair<string,string>> compoundsNameList = new List<KeyValuePair<string,string>>();
        List<KeyValuePair<string,string>> preCompoundsNameList = new List<KeyValuePair<string,string>>();

        /// <summary>
        /// 委托单号
        /// </summary>
        string ReportNo = string.Empty;

        //调整一个竖表格的总列数
        int horizontalSheetColumnCount = 10;
        /// <summary>
        /// 每个化合物的datatable
        /// </summary>
        DataSet compoundsDataSet = new DataSet();

        /// <summary>
        /// 火焰元素的datatable合集
        /// </summary>
        DataSet FiresDataSet = new DataSet();

        /// <summary>
        /// 石墨元素的datatable合集
        /// </summary>
        DataSet GraphiteDataSet = new DataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender,RoutedEventArgs e)
        {
            topScrollViewer.DragEnter += scDragEnter;
            topScrollViewer.Drop += scDrop;
            mainScrollViewer.DragEnter += scDragEnter;
            mainScrollViewer.Drop += scDrop;
            samplingquantityLabel.Tag = 0;
            dilutionratioLabel.Tag = 1;
            LDMCLabel.Tag = 2;
            constantvolumeLabel.Tag = 3;
        }

        /// <summary>
        /// 拖动进入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scDragEnter(object sender,DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 拖动放下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scDrop(object sender,DragEventArgs e)
        {
            //foreach(string str in e.Data.GetFormats())
            //{
            //	MessageBox.Show(str);
            //}
            ScrollViewer scrollViewer = sender as ScrollViewer;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;

                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (int.Parse(scrollViewer.Tag.ToString()) == 0)
                {
                    //导入模板
                    CreateTemplate(paths[0]);
                }
                else if (int.Parse(scrollViewer.Tag.ToString()) == 1)
                {
                    //创建数据结构
                    CreateExcel(paths[0]);
                }
            }
            e.Handled = true;
        }

        /// <summary>
        /// 导入模板到程序里面
        /// </summary>
        /// <param name="path"></param>
        private void CreateTemplate(string path)
        {
            string symbol = "：";
            if (File.Exists(path))
            {
                List<string> alldata = File.ReadAllLines(path,Encoding.UTF8).ToList();
                foreach (string data in alldata)
                {
                    //加载常规设置项
                    string key = data.Split(symbol)[0];
                    string value = data.Split(symbol)[1];
                    if (key == samplingquantityLabel.Content.ToString())
                    {
                        foreach (ComboBoxItem comboBoxItem in samplingquantityComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    else if (key == constantvolumeLabel.Content.ToString())
                    {
                        foreach (ComboBoxItem comboBoxItem in constantvolumeComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    else if ((key + symbol) == FireLabel.Content.ToString())
                    {
                        foreach (ComboBoxItem comboBoxItem in FireComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    else if ((key + symbol) == GraphiteLabel.Content.ToString())
                    {
                        foreach (ComboBoxItem comboBoxItem in GraphiteComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    else if ((key + symbol) == AccuracyLabel.Content.ToString())
                    {
                        foreach (ComboBoxItem comboBoxItem in AccuracyComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    else if ((key + symbol) == FormulaLabel.Content.ToString())
                    {
                        foreach (ComboBoxItem comboBoxItem in FormulaComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    else if (key == testJCRadioButton.Content.ToString())
                    {
                        testJCRadioButton.IsChecked = true;
                        foreach (ComboBoxItem comboBoxItem in ZDJCCompanyComboBox.Items)
                        {
                            if (value == comboBoxItem.Content.ToString())
                            {
                                comboBoxItem.IsSelected = true;
                            }
                        }
                    }
                    //加载化合物项
                    else
                    {
                        //有没有添加化合物的文档
                        if (maingrid.Children.Count > 0)
                        {
                            TabControl tabControl = maingrid.Children[0] as TabControl;
                            foreach (TabItem tabItem in tabControl.Items)
                            {
                                StackPanel stackPanel = tabItem.Header as StackPanel;
                                Label label = stackPanel.Children[1] as Label;
                                TextBox textBox = stackPanel.Children[2] as TextBox;
                                if (label.Content.ToString() == key)
                                {
                                    textBox.Text = value;
                                }
                            }
                        }
                        else
                        {
                            KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(key,value);
                            preCompoundsNameList.Add(keyValuePair);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 通过文本创造核心内容
        /// </summary>
        /// <param name="path"></param>
        private void CreateExcel(string path)
        {
            AllClear();
            IWorkbook workbook = null;
            TabControl tabControl = new TabControl();
            tabControl.Name = "tabControl";
            if (File.Exists(path))
            {
                using (FileStream fs = File.OpenRead(path))
                {
                    // 2007版本
                    if (path.Contains(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    // 2003版本
                    else if (path.Contains(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    if (workbook != null)
                    {
                        ISheet sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        IRow firstrow = sheet.GetRow(sheet.FirstRowNum);
                        int Num = 0;
                        //第一行是啊化合物名称
                        int columnCount = firstrow.LastCellNum;//多少列
                        for (int j = 0; j < columnCount; j++)
                        {
                            ICell firstCell = firstrow.GetCell(j);
                            if (firstCell != null)
                            {
                                if (firstCell.ColumnIndex == 0)
                                {
                                    ReportNo = firstCell.StringCellValue;
                                }
                                else if (firstCell.StringCellValue != string.Empty && firstCell.StringCellValue != "" && firstCell.StringCellValue != "分析物名称")
                                {
                                    string compoundName = firstCell.StringCellValue.Trim();
                                    DataTable dataTable = new DataTable();
                                    dataTable.TableName = compoundName;
                                    CreateDataTable(tabControl,sheet,dataTable,firstCell.ColumnIndex,Num);
                                    Num++;
                                }
                            }
                        }
                    }


                    maingrid.Children.Add(tabControl);
                    ReportNoLabel.Content = ReportNo;
                }
            }
        }

        /// <summary>
        /// 全部清空,重置
        /// </summary>
        private void AllClear()
        {
            compoundsNameList.Clear();
            ReportNo = string.Empty;
            ReportNoLabel.Content = ReportNo;
            FiresDataSet.Tables.Clear();
            GraphiteDataSet.Tables.Clear();
            maingrid.Children.Clear();
        }

        private void CreateDataTable(TabControl tabControl,ISheet sheet,DataTable dataTable,int compoundsNum,int keyValueNum)
        {
            int rowCount = sheet.LastRowNum;//总行数
            string compoundName = dataTable.TableName;
            int skipNum = 2;
            //分辨火焰或石墨
            //左火焰右石墨
            if (compoundName.Contains("["))
            {
                FiresDataSet.Tables.Add(dataTable);
            }
            else if(compoundName.Contains("]"))
            {
                skipNum = 1;
                GraphiteDataSet.Tables.Add(dataTable);
            }
            compoundName = compoundName.Remove(0,2);

            for (int i = 1; i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                //由于Excel在非数据区进行了格式设置，那么sheet.LastRowNum 得到的值就会与实际得到的值不符。从而因有非空验证，造成导入失败。
                //所以直接先判断第一个单元格是否为空，在进行后面的操作
                ICell firstCell = row.GetCell(0);
                if (firstCell != null)
                {
                    //第二行都是表头，要组成datatable
                    if (i == 1)
                    {
                        for (int j = 0; j < 6; j++)
                        {
                            ICell secondCell = row.GetCell(j);
                            if (secondCell != null)
                            {
                                if (secondCell.StringCellValue != string.Empty && secondCell.StringCellValue != "")
                                {
                                    dataTable.Columns.Add(secondCell.StringCellValue);
                                }
                            }
                        }
                        dataTable.Columns.Add("浓度");
                    }
                    //第三行开始是数据
                    else
                    {
                        DataRow dataRow = dataTable.NewRow();
                        for (int k = 0; k < 6; k++)
                        {
                            ICell cell = row.GetCell(k);
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    dataRow[k] = cell.NumericCellValue;
                                }
                                else
                                {
                                    dataRow[k] = cell.StringCellValue.Trim();
                                }
                            }
                        }
                        ICell newCell = row.GetCell(compoundsNum);
                        if (newCell.CellType == CellType.Numeric)
                        {
                            dataRow[dataRow.ItemArray.Length - 1] = newCell.NumericCellValue;
                        }
                        else
                        {
                            dataRow[dataRow.ItemArray.Length - 1] = newCell.StringCellValue;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            //删掉火焰石墨不要的列
            dataTable.Columns.RemoveAt(skipNum);
            //删掉每个元素浓度为空的行
            for (int i = dataTable.Rows.Count -1; i >=0; i--)
            {
                string value = dataTable.Rows[i][dataTable.Columns.Count - 1].ToString();
                if (value == "/")
                {
                    dataTable.Rows.RemoveAt(i);
                }
            }
            //把样品名称调到最开始
            dataTable.Columns[dataTable.Columns.Count - 2].SetOrdinal(0);
            /*dataTable.Columns[dataTable.Columns.Count - 1].SetOrdinal(0);
            if (sampleNameList.Count != dataTable.Rows.Count)
            {
                for (int l = 0; l < dataTable.Rows.Count; l++)
                {
                    sampleNameList.Add(dataTable.Rows[l][0].ToString());
                }
            }*/
            AddParallelSamplesToList(dataTable);

            TabItem tabItem = new TabItem();
            //tabItem.Header = name[1] + " | " + name[2];
            StackPanel stackPanel = CreateStackPanel(compoundName,keyValueNum);
            tabItem.Header = stackPanel;
            DataGrid dg = new DataGrid();
            dg.Name = "dataGrid";
            dg.ItemsSource = dataTable.DefaultView;
            dg.CanUserSortColumns = true;
            dg.CanUserReorderColumns = true;
            tabItem.Content = dg;
            tabControl.Items.Add(tabItem);
            //compoundsDataSet.Tables.Add(dataTable);
        }

        /// <summary>
        /// 创建tabheader用的stackpanel
        /// </summary>
        /// <returns></returns>
        private StackPanel CreateStackPanel(string compoundsName,int num)
        {
            StackPanel stackPanel = new StackPanel();
            stackPanel.Orientation = Orientation.Horizontal;
            stackPanel.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            stackPanel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;

            Label numLabel = new Label();
            numLabel.Content = (num + 1).ToString() + ".";
            numLabel.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
            numLabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;

            Label compoundslabel = new Label();
            compoundslabel.Content = compoundsName;
            compoundslabel.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
            compoundslabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            TextBox textBox = new TextBox();
            textBox.Width = 50;
            textBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            textBox.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            textBox.KeyUp += Tab_TextBox_KeyUp;
            if (preCompoundsNameList.Count > 0)
            {
                foreach (KeyValuePair<string,string> keyValuePair in preCompoundsNameList)
                {
                    if (keyValuePair.Key == compoundsName)
                    {
                        textBox.Text = keyValuePair.Value;
                    }
                }
            }

            stackPanel.Children.Add(numLabel);
            stackPanel.Children.Add(compoundslabel);
            stackPanel.Children.Add(textBox);

            return stackPanel;
        }

        /// <summary>
        /// enter切换检出限
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tab_TextBox_KeyUp(object sender,KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBox textbox = sender as TextBox;
                StackPanel stackPanel = textbox.Parent as StackPanel;
                TabItem tabItem = stackPanel.Parent as TabItem;
                TabControl tabControl = tabItem.Parent as TabControl;
                int tabNum = tabControl.Items.IndexOf(tabItem);
                //到达最大值
                TabItem nextTabItem;
                if (tabNum == tabControl.Items.Count - 1)
                {
                    nextTabItem = tabControl.Items[0] as TabItem;
                }
                else
                {
                    nextTabItem = tabControl.Items[tabNum + 1] as TabItem;
                }
                StackPanel nextStackPanel = nextTabItem.Header as StackPanel;
                foreach (var item in nextStackPanel.Children)
                {
                    if (item.GetType() == typeof(TextBox))
                    {
                        TextBox nextTextBox = item as TextBox;
                        Keyboard.Focus(nextTextBox);
                        nextTextBox.Focus();
                    }
                }

            }
        }

        /// <summary>
        /// 添加平行样
        /// </summary>
        private void AddParallelSamplesToList(DataTable dataTable)
        {
            //由于只有竖表不用分组
            for (int i = 0; i < dataTable.Rows.Count - 1; i++)
            {
                string value = dataTable.Rows[i]["试样识别码"].ToString();
                if (value.Contains("Dup"))
                {
                    DataRow dataRow = dataTable.NewRow();
                    dataRow[0] = value.Replace("Dup","平均");
                    dataRow[1] = "/";
                    dataRow[2] = "/";
                    dataRow[3] = "/";
                    dataRow[4] = "/";
                    dataRow[5] = "/";
                    dataTable.Rows.InsertAt(dataRow,i + 1);
                }
            }
        }


        /// <summary>
        /// 导出生成Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importExcel_Click(object sender,RoutedEventArgs e)
        {
            if (FiresDataSet.Tables.Count == 0 && GraphiteDataSet.Tables.Count == 0)
            {
                return;
            }
            //只有竖表没有横表
            var workbook = new HSSFWorkbook();
            //先出火焰的再出石墨的
            foreach (DataTable firetable in FiresDataSet.Tables)
            {
                var sheet = workbook.CreateSheet(firetable.TableName.Remove(0,2));
                sheet.ForceFormulaRecalculation = true;
                CreateHorizontalExcel(sheet,firetable);
            }
            foreach (DataTable graptable in GraphiteDataSet.Tables)
            {
                var sheet = workbook.CreateSheet(graptable.TableName.Remove(0,2));
                sheet.ForceFormulaRecalculation = true;
                CreateHorizontalExcel(sheet,graptable);
            }
            ExportToExcel(workbook);
        }

        /// <summary>
        /// 添加自己填的检出限
        /// </summary>
        private void AddDetectionLimit()
        {
            compoundsNameList.Clear();
            TabControl tabControl = maingrid.Children[0] as TabControl;
            foreach (TabItem tabItem in tabControl.Items)
            {
                string compoundsName = string.Empty;
                string modelC = string.Empty;
                StackPanel stackPanel = tabItem.Header as StackPanel;
                foreach (var item in stackPanel.Children)
                {
                    if (item.GetType() == typeof(Label))
                    {
                        compoundsName = (item as Label).Content.ToString();
                    }
                    else if (item.GetType() == typeof(TextBox))
                    {
                        if ((item as TextBox).Text != null && (item as TextBox).Text != "" && (item as TextBox).Text != string.Empty)
                        {

                            modelC = (item as TextBox).Text;
                        }
                    }
                }
                KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(compoundsName,modelC);
                compoundsNameList.Add(keyValuePair);
            }

            //if (compoundsNameList.Count > 2)
            //{
            //    KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>("以下空白",string.Empty);
            //    compoundsNameList.Add(keyValuePair);
            //}
        }

        /// <summary>
        /// 创建竖表Excel
        /// </summary>
        private void CreateHorizontalExcel(ISheet sheet,DataTable datatable)
        {
            HSSFWorkbook workbook = sheet.Workbook as HSSFWorkbook;
            //设置顶部大标题样式
            HSSFCellStyle cellStyle = CreateStyle(workbook);
            HSSFCellStyle bordercellStyle = CreateStyle(workbook);
            bordercellStyle.BorderLeft = BorderStyle.Thin;
            bordercellStyle.BorderTop = BorderStyle.Thin;
            bordercellStyle.BorderLeft = BorderStyle.Thin;
            bordercellStyle.BorderRight = BorderStyle.Thin;
            //int Count = 0;
            //前五行 大表头
            for (int i = 0; i < 6; i++)
            {
                //第一行最右显示委托单号
                HSSFRow row = (HSSFRow)sheet.CreateRow(i); //创建行或者获取行
                row.HeightInPoints = 20;
                switch (i)
                {
                    case 0:
                        {
                            row.HeightInPoints = 30;

                            var nameCell = row.CreateCell(0);
                            HSSFCellStyle newcellStyle = CreateStyle(workbook);
                            var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建字体
                            cellStyleFont.IsBold = true; //字体加粗
                            cellStyleFont.FontHeightInPoints = 20; //字体大小
                            newcellStyle.SetFont(cellStyleFont); //将字体绑定到样式
                            nameCell.CellStyle = newcellStyle;
                            nameCell.SetCellValue("原子吸收分光光度法分析原始记录表（土）");
                            CellRangeAddress region = new CellRangeAddress(i,i,0,horizontalSheetColumnCount - 1);
                            sheet.AddMergedRegion(region);
                            break;
                        }
                    case 1:
                        {
                            var firstCell = row.CreateCell(0);
                            firstCell.CellStyle = cellStyle;
                            firstCell.SetCellValue("样品类别：");
                            var secondCell = row.CreateCell(2);
                            secondCell.CellStyle = cellStyle;
                            secondCell.SetCellValue("分析项目：");
                            var thirdCell = row.CreateCell(4);
                            thirdCell.CellStyle = cellStyle;
                            thirdCell.SetCellValue("收样日期：");
                            var fourthCell = row.CreateCell(6);
                            fourthCell.CellStyle = cellStyle;
                            fourthCell.SetCellValue("分析日期：");
                            var fifthCell = row.CreateCell(8);
                            fifthCell.CellStyle = cellStyle;
                            fifthCell.SetCellValue("委托编号：");
                            var reportnoCell = row.CreateCell(9);
                            reportnoCell.CellStyle = cellStyle;
                            reportnoCell.SetCellValue(ReportNo);
                            break;
                        }
                    case 2:
                        {
                            var firstCell = row.CreateCell(0);
                            firstCell.CellStyle = cellStyle;
                            firstCell.SetCellValue("方法依据：");
                            var secondCell = row.CreateCell(2);
                            secondCell.CellStyle = cellStyle;
                            secondCell.SetCellValue("仪器型号：");
                            var thirdCell = row.CreateCell(4);
                            thirdCell.CellStyle = cellStyle;
                            thirdCell.SetCellValue("仪器编号：");
                            var fourthCell = row.CreateCell(6);
                            fourthCell.CellStyle = cellStyle;
                            fourthCell.SetCellValue("方法检出限：");
                            var ZDJCCompanyCell = row.CreateCell(7);
                            ZDJCCompanyCell.CellStyle = cellStyle;
                            foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
                            {
                                if (keyValuePair.Key == sheet.SheetName)
                                {
                                    ZDJCCompanyCell.SetCellValue(keyValuePair.Value + ZDJCCompanyComboBox.Text);
                                }
                            }
                            var fifthCell = row.CreateCell(8);
                            fifthCell.CellStyle = cellStyle;
                            fifthCell.SetCellValue("计算公式：");
                            var FormulaCell = row.CreateCell(9);
                            FormulaCell.CellStyle = cellStyle;
                            if (FiresDataSet.Tables.Contains(datatable.TableName))
                            {
                                ComboBoxItem item = FormulaComboBox.Items[0] as ComboBoxItem;
                                FormulaCell.SetCellValue(item.Content.ToString());
                            }
                            else if (GraphiteDataSet.Tables.Contains(datatable.TableName))
                            {
                                ComboBoxItem item = FormulaComboBox.Items[1] as ComboBoxItem;
                                FormulaCell.SetCellValue(item.Content.ToString());
                            }
                            CellRangeAddress secondregion = new CellRangeAddress(i,i + 1,horizontalSheetColumnCount - 1,horizontalSheetColumnCount - 1);
                            sheet.AddMergedRegion(secondregion);
                            break;
                        }
                    case 3:
                        {
                            var firstCell = row.CreateCell(0);
                            firstCell.CellStyle = cellStyle;
                            firstCell.SetCellValue("测定波长：");
                            var secondCell = row.CreateCell(2);
                            secondCell.CellStyle = cellStyle;
                            secondCell.SetCellValue("狭缝：");
                            var thirdCell = row.CreateCell(4);
                            thirdCell.CellStyle = cellStyle;
                            thirdCell.SetCellValue("火焰法：");
                            var fourthCell = row.CreateCell(6);
                            fourthCell.CellStyle = cellStyle;
                            fourthCell.SetCellValue("石墨炉法：");
                            if (FiresDataSet.Tables.Contains(datatable.TableName))
                            {
                                var fifthCell = row.CreateCell(5);
                                fifthCell.CellStyle = cellStyle;
                                fifthCell.SetCellValue("√");
                            }
                            else if (GraphiteDataSet.Tables.Contains(datatable.TableName))
                            {
                                var fifthCell = row.CreateCell(7);
                                fifthCell.CellStyle = cellStyle;
                                fifthCell.SetCellValue("√");
                            }
                            break;
                        }
                    case 4:
                        {
                            for (int j = 0; j < horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == 0)
                                {
                                    cell.SetCellValue("分析编号");
                                }
                                else if (j == 8)
                                {
                                    string value = "土壤样品浓度\n"
                                        + "C" + ZDJCCompanyComboBox.Text;
                                    cell.SetCellValue(value);
                                }
                                else if (j == 9)
                                {
                                    string value = "备注";
                                    cell.SetCellValue(value);
                                }
                                else if (j == 1)
                                {
                                    cell.SetCellValue(datatable.Columns[j - 1].ColumnName);
                                    CellRangeAddress secondregion = new CellRangeAddress(4,5,j,j + 1);
                                    sheet.AddMergedRegion(secondregion);
                                }
                                else
                                {
                                    cell.SetCellValue(datatable.Columns[j - 2].ColumnName);
                                }

                                if (j != 1 && j != 2)
                                {
                                    CellRangeAddress secondregion = new CellRangeAddress(4,5,j,j);
                                    sheet.AddMergedRegion(secondregion);
                                }
                            }
                            break;
                        }
                    default:
                        {
                            for (int j = 0; j < horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                cell.SetCellValue(string.Empty);
                            }
                            break;
                        }
                }
            }

            //正规表格
            for (int i = 0; i < datatable.Rows.Count+1; i++)
            {
                if (i == datatable.Rows.Count)
                {
                }
                else
                {
                    int rowNum = i + 6;
                    HSSFRow row = (HSSFRow)sheet.CreateRow(rowNum); //创建行或者获取行
                    row.HeightInPoints = 20;
                    for (int j = 0; j < horizontalSheetColumnCount; j++)
                    {
                        var cell = row.CreateCell(j);
                        cell.CellStyle = bordercellStyle;
                        switch (j)
                        {
                            case 0:
                                {
                                    cell.SetCellValue(i + 1);
                                    break;
                                }
                            case 1:
                                {
                                    cell.SetCellValue(datatable.Rows[i][j - 1].ToString());
                                    CellRangeAddress secondregion = new CellRangeAddress(rowNum,rowNum,j,j + 1);
                                    sheet.AddMergedRegion(secondregion);
                                    break;
                                }
                            case 3:
                                {
                                    //计算精度函数
                                    string value = datatable.Rows[i][j - 2].ToString();
                                    if (!value.Contains("/"))
                                    {
                                        value = CalculateAccuracyCFour(value);
                                    }
                                    cell.SetCellValue(value);
                                    break;
                                }
                            case 4:
                                {
                                    //计算精度函数
                                    string value = datatable.Rows[i][j - 2].ToString();
                                    if (!value.Contains("/") && !value.Contains("%"))
                                    {
                                        value = CalculateAccuracyCPercent(value);
                                    }
                                    cell.SetCellValue(value);
                                    break;
                                }
                            case 8:
                                {
                                    //计算精度函数
                                    string value = string.Empty;
                                    cell.SetCellValue(value);
                                    break;
                                }
                            case 9:
                                {
                                    //备注
                                    string value = string.Empty;
                                    cell.SetCellValue(value);
                                    break;
                                }
                            default:
                                {
                                    cell.SetCellValue(datatable.Rows[i][j - 2].ToString());
                                    break;
                                }
                        }
                    }
                }
            }
            

            //自动调整列距
            for (int i = 0; i < horizontalSheetColumnCount; i++)
            {
                sheet.AutoSizeColumn(i);
                int width = sheet.GetColumnWidth(i);
                if (width < 20 * 256)
                {
                    sheet.SetColumnWidth(i,20 * 256);
                }
            }
        }


        /// <summary>
        /// 导出到Excel
        /// </summary>
        /// <param name="workbook"></param>
        private void ExportToExcel(HSSFWorkbook workbook)
        {
            //自己选位置
            /*System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
			fbd.ShowDialog();
			if (fbd.SelectedPath != string.Empty)
			{
				string filename = sheet.SheetName + ".xls";
				string path = System.IO.Path.Combine(fbd.SelectedPath,filename);
				using (FileStream stream = new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite))
				{
					workbook.Write(stream);
					stream.Flush();
				}
			}*/
            //特定位置
            try
            {
                string path = @"E:\CreateExcel\" + ReportNo + @"\";
                //创建用户临时图片文件夹或者清空临时文件夹所有文件
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string filename = ReportNo + "-" + workbook.GetSheetAt(0).SheetName + ".xls";
                string fullpath = System.IO.Path.Combine(path,filename);
                if (File.Exists(fullpath))
                {
                    File.Delete(fullpath);
                }
                using (FileStream stream = new FileStream(fullpath,FileMode.OpenOrCreate,FileAccess.ReadWrite))
                {
                    workbook.Write(stream);
                    stream.Flush();
                }
                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo(fullpath);
                processStartInfo.UseShellExecute = true;
                process.StartInfo = processStartInfo;
                process.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private string CalculateAccuracySu(double c)
        {
            string newC = c.ToString();
            string[] strC = newC.Split(".");
            //有小数点
            if (strC.Length > 1)
            {
                string leftC = strC[0];
                string rightC = strC[1];
                if (leftC.Length > 3)
                {
                    return leftC;
                }
                else
                {
                    newC = Math.Round(c,3 - leftC.Length).ToString();
                }
            }
            //计算完毕都需要补零
            if (newC.Length < 4 && newC.Contains("."))
            {
                newC += "0";
            }
            else if (newC.Length < 4 && !newC.Contains("."))
            {
                newC += ".";
                int lenC = newC.Length;
                for (int i = 0; i < 4 - lenC; i++)
                {
                    newC += "0";
                }
            }

            return newC;
        }

        /// <summary>
        /// 科学计数法
        /// </summary>
        /// <param name="testNum"></param>5
        /// <returns></returns>
        private string ScientificCounting(double testNum)
        {
            string returnnum = string.Empty;
            string oneNum = "1";
            if (testNum.ToString().Length > 4)
            {
                for (int i = 0; i < testNum.ToString().Length - 1; i++)
                {
                    oneNum += "0";
                }

                double onenum = double.Parse(oneNum);
                returnnum = (testNum / onenum).ToString() + "×" + "10" + (testNum.ToString().Length - 1).ToString();
            }
            return returnnum;
        }

        /// <summary>
        /// 计算平行样浓度平均值
        /// </summary>
        /// <param name="compoundName"></param>
        /// <param name="modelC"></param>
        /// <param name="sampleName1"></param>
        /// <param name="sampleName2"></param>
        /// <returns></returns>
        private string CompareCompoundWithFormula(string compoundName,string modelC,string sampleName1,string sampleName2)
        {
            //计算公式C = Ci × f × V1 / (m × Wdm)
            //稀释倍数
            double f = double.NaN;
            //定容体积
            double V1 = double.NaN;
            //取样量
            double m = double.NaN;
            //干物质含量
            double Wdm = double.NaN;

            //目标物上机测定浓度
            double Ci;
            double C1 = double.NaN;
            double C2 = double.NaN;
            foreach (DataTable dataTable in compoundsDataSet.Tables)
            {
                //找到该化合物对应的datatable
                if (dataTable.TableName == compoundName)
                {
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            //找到该化合物对应的样品编号和浓度数据
                            string dtsampleName = dataTable.Rows[i][j].ToString();

                            if (dtsampleName == sampleName1)
                            {
                                string potency = dataTable.Rows[i]["含量"].ToString();
                                if (!potency.Contains("-"))
                                {
                                    Ci = double.Parse(potency);
                                    //公式计算
                                    //先用写死的，然后之后学习反射
                                    double moleculeV1 = double.Parse((constantvolumeComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double denominatorM = double.Parse((samplingquantityComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double taggetC = double.Parse((FireComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double ZDJCCi = double.Parse((ZDJCCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    //单位换算
                                    double k = taggetC / ZDJCCi * (moleculeV1 / denominatorM);

                                    C1 = (Ci * f * V1 / (m * Wdm)) * k;

                                }
                            }
                            if (dtsampleName == sampleName2)
                            {
                                string potency = dataTable.Rows[i]["含量"].ToString();
                                if (!potency.Contains("-"))
                                {
                                    Ci = double.Parse(potency);
                                    //公式计算
                                    //先用写死的，然后之后学习反射
                                    double moleculeV1 = double.Parse((constantvolumeComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double denominatorM = double.Parse((samplingquantityComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double taggetC = double.Parse((FireComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double ZDJCCi = double.Parse((ZDJCCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    //单位换算
                                    double k = taggetC / ZDJCCi * (moleculeV1 / denominatorM);

                                    C2 = (Ci * f * V1 / (m * Wdm)) * k;
                                }
                            }
                        }
                    }
                }
            }

            double C = (C1 + C2) / 2;
            if (C > double.Parse(modelC))
            {
                string realC = CalculateAccuracyC(compoundName,C.ToString());
                return realC;
            }

            return "ND";
        }


        /// <summary>
        /// 计算目标化合物浓度
        /// </summary>
        /// <param name="compoundName"></param>
        /// <param name="modelC"></param>
        /// <param name="sampleName"></param>
        /// <returns></returns>
        private string CompareCompoundWithFormula(string compoundName,string modelC,string sampleName)
        {
            //计算公式C = Ci × f × V1 / (m × Wdm)
            //稀释倍数
            double f = double.NaN;
            //定容体积
            double V1 = double.NaN;
            //取样量
            double m = double.NaN;
            //干物质含量
            double Wdm = double.NaN;

            //目标物上机测定浓度
            double Ci;
            foreach (DataTable dataTable in compoundsDataSet.Tables)
            {
                //找到该化合物对应的datatable
                if (dataTable.TableName == compoundName)
                {
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            //找到该化合物对应的样品编号和浓度数据
                            string dtsampleName = dataTable.Rows[i][j].ToString();
                            if (dtsampleName == sampleName)
                            {
                                string potency = dataTable.Rows[i]["含量"].ToString();
                                if (!potency.Contains("-"))
                                {
                                    Ci = double.Parse(potency);
                                    //公式计算
                                    //先用写死的，然后之后学习反射
                                    double C = double.NaN;
                                    double moleculeV1 = double.Parse((constantvolumeComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double denominatorM = double.Parse((samplingquantityComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double taggetC = double.Parse((FireComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    double ZDJCCi = double.Parse((ZDJCCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
                                    //单位换算
                                    double k = taggetC / ZDJCCi * (moleculeV1 / denominatorM);

                                    C = (Ci * f * V1 / (m * Wdm)) * k;

                                    if (C > double.Parse(modelC))
                                    {
                                        string realC = CalculateAccuracyC(compoundName,C.ToString());
                                        return realC;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return "ND";
        }

        private HSSFCellStyle CreateStyle(HSSFWorkbook workbook)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
            cellStyle.WrapText = true;//自动换行
                                      //cellStyle.BorderBottom = BorderStyle.Thin;
                                      //cellStyle.BorderRight = BorderStyle.Thin;
                                      //cellStyle.BorderTop = BorderStyle.Thin;
                                      //cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
            cellStyle.RightBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;

            return cellStyle;
        }

        private HSSFCellStyle CreateGreyStyle(HSSFWorkbook workbook)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
            var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建变色字体
            cellStyleFont.Color = HSSFColor.Grey25Percent.Index;
            cellStyle.SetFont(cellStyleFont); //将字
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
            cellStyle.WrapText = true;//自动换行
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
            cellStyle.RightBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;

            return cellStyle;
        }

        private HSSFCellStyle CreateRedStyle(HSSFWorkbook workbook)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
            var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建变色字体
            cellStyleFont.Color = HSSFColor.Red.Index;
            cellStyle.SetFont(cellStyleFont); //将字
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
            cellStyle.WrapText = true;//自动换行
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
            cellStyle.RightBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;

            return cellStyle;
        }

        private HSSFCellStyle CreateColorStyle(HSSFWorkbook workbook,string setvalue)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
            var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建变色字体

            if (setvalue.Contains("Dup"))
            {
                cellStyleFont.Color = HSSFColor.SeaGreen.Index;
            }
            else if (setvalue.Contains("MS"))
            {
                cellStyleFont.Color = HSSFColor.Red.Index;
            }
            else if (setvalue.Length > 3 && setvalue.Remove(0,setvalue.Length - 3).Contains("102"))
            {
                cellStyleFont.Color = HSSFColor.SkyBlue.Index;
            }

            cellStyle.SetFont(cellStyleFont); //将字
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
            cellStyle.WrapText = true;//自动换行
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
            cellStyle.RightBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;

            return cellStyle;
        }
        private void ComplierCode(string expression)
        {
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

            CompilerParameters objCompilerParameters = new CompilerParameters();

            //添加需要引用的dll
            objCompilerParameters.ReferencedAssemblies.Add("System.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Windows.Forms.dll");
            //是否生成可执行文件
            objCompilerParameters.GenerateExecutable = false;
            //是否生成在内存中
            objCompilerParameters.GenerateInMemory = true;

            //编译代码
            CompilerResults cr = objCSharpCodePrivoder.CompileAssemblyFromSource(objCompilerParameters,FormulaComboBox.Text);

            if (cr.Errors.HasErrors)
            {
                var msg = string.Join(Environment.NewLine,cr.Errors.Cast<CompilerError>().Select(err => err.ErrorText));
                MessageBox.Show(msg,"编译错误");
            }
            else
            {
                Assembly objAssembly = cr.CompiledAssembly;
                object objHelloWorld = objAssembly.CreateInstance("Test");
                MethodInfo objMI = objHelloWorld.GetType().GetMethod("Hello");
                objMI.Invoke(objHelloWorld,null);
            }
        }

        /// <summary>
        /// 补齐四位数的零
        /// </summary>
        /// <param name="compoundName"></param>
        /// <param name="v"></param>
        /// <returns></returns>
        private string CalculateAccuracyCPercent(string value)
        {
            double num = double.Parse(value);
            num *= 100;
            value = num.ToString() + "%";
            return value;
        }

        /// <summary>
        /// 补齐四位数的零
        /// </summary>
        /// <param name="compoundName"></param>
        /// <param name="v"></param>
        /// <returns></returns>
        private string CalculateAccuracyCFour(string value)
        {
            string[] beforeValue = value.Split(".");
            int num;
            //没有小数点的
            if (beforeValue.Length < 2)
            {
                num = 4;
            }
            else
            {
                num = 4 - beforeValue[beforeValue.Length - 1].Length;
            }
            //计算后补零
            if (num != 0)
            {
                if (value.ToString().Contains("."))
                {
                    string answer = value.ToString();
                    for (int i = 0; i < num; i++)
                    {
                        answer += "0";
                    }
                    return answer;
                }
                else
                {
                    string answer = value.ToString() + ".";
                    for (int i = 0; i < num; i++)
                    {
                        answer += "0";
                    }
                    return answer;
                }
            }

            return value;
        }

        /// <summary>
        /// 计算精度
        /// </summary>
        /// <param name="compoundName"></param>
        /// <param name="C"></param>
        /// <returns></returns>
        private string CalculateAccuracyC(string compoundName,string C)
        {
            double answer = double.NaN;
            string accuracy = AccuracyComboBox.SelectedItem.ToString();
            int num = 0;
            //选择默认方式
            if (accuracy == AccuracyComboBox.Items[0].ToString())
            {
                foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
                {

                    if (keyValuePair.Key == compoundName)
                    {
                        string[] beforeValue = keyValuePair.Value.Split(".");
                        //没有小数点的
                        if (beforeValue.Length < 2)
                        {
                            answer = Math.Round(double.Parse(C),0);
                        }
                        else
                        {
                            answer = Math.Round(double.Parse(C),beforeValue[beforeValue.Length - 1].Length);
                        }
                        string[] afterValue = answer.ToString().Trim().Split(".");
                        if (afterValue.Length < 2)
                        {
                            num = beforeValue[beforeValue.Length - 1].Length;
                        }
                        else
                        {
                            num = beforeValue[beforeValue.Length - 1].Length - afterValue[afterValue.Length - 1].Length;
                        }
                    }
                }
            }
            //选择其他位数
            else
            {
                string[] beforeValue = accuracy.Split(":");
                answer = Math.Round(double.Parse(C),int.Parse(beforeValue[beforeValue.Length - 1]));
                string[] afterValue = answer.ToString().Trim().Split(".");
                num = int.Parse(beforeValue[beforeValue.Length - 1]) - afterValue[afterValue.Length - 1].Length;
            }
            //计算后补零
            if (num != 0)
            {
                if (answer.ToString().Contains("."))
                {
                    string newanswer = answer.ToString();
                    for (int i = 0; i < num; i++)
                    {
                        newanswer += "0";
                    }
                    return newanswer;
                }
                else
                {
                    string newanswer = answer.ToString() + ".";
                    for (int i = 0; i < num; i++)
                    {
                        newanswer += "0";
                    }
                    return newanswer;
                }
            }
            return answer.ToString().Trim();
        }


        /// <summary>
        /// 导出模板按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importAll_Click(object sender,RoutedEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "文本文件(*.txt)|*.txt|所有文件|*.*";//设置文件类型
                                                      //sfd.FileName = "保存";//设置默认文件名
            sfd.DefaultExt = "txt";//设置默认格式（可以不设）
            sfd.AddExtension = true;//设置自动在文件名中添加扩展名
            sfd.ShowDialog();
            if (sfd.FileName != string.Empty)
            {
                string fullpath = sfd.FileName;
                using (FileStream stream = new FileStream(fullpath,FileMode.Create,FileAccess.ReadWrite))
                {
                    StreamWriter streamWriter = new StreamWriter(stream);
                    //streamWriter.WriteLine(strReportNoLabel.Content + ReportNo);
                    streamWriter.WriteLine(samplingquantityLabel.Content + "：" + samplingquantityComboBox.Text);
                    streamWriter.WriteLine(constantvolumeLabel.Content + "：" + constantvolumeComboBox.Text);
                    streamWriter.WriteLine(FireLabel.Content + FireComboBox.Text);
                    streamWriter.WriteLine(GraphiteLabel.Content + GraphiteComboBox.Text);
                    streamWriter.WriteLine(AccuracyLabel.Content + AccuracyComboBox.Text);
                    streamWriter.WriteLine(FormulaLabel.Content + FormulaComboBox.Text);
                    streamWriter.WriteLine(testJCRadioButton.Content + "：" + ZDJCCompanyComboBox.Text);

                    foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
                    {
                        streamWriter.WriteLine(keyValuePair.Key + "：" + keyValuePair.Value);
                    }
                    streamWriter.Flush();
                    stream.Flush();
                }
                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo(fullpath);
                processStartInfo.UseShellExecute = true;
                process.StartInfo = processStartInfo;
                process.Start();
            }
        }

        /// <summary>
        /// 生成compoundsNameList
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importExcel_MouseMove(object sender,MouseEventArgs e)
        {
            if (FiresDataSet.Tables.Count == 0 && GraphiteDataSet.Tables.Count == 0)
            {
                return;
            }
            AddDetectionLimit();
        }

        private void TextBox_TextChanged(object sender,TextChangedEventArgs e)
        {
            TextBox textbox = sender as TextBox;
            string text = textbox.Text.Trim();
            if (text != null && text != "")
            {
                Regex r = new Regex(@"^([0-9\.]*)$");
                if (r.IsMatch(textbox.Text.Trim()) == false)
                {
                    textbox.Text = textbox.Text.Remove(textbox.Text.Length - 1,1);
                }

                int numDecimal = 0;
                for (int i = 0; i < textbox.Text.Length; i++)
                {
                    if (textbox.Text[i].ToString() == ".")
                    {
                        numDecimal++;
                    }
                }

                if (numDecimal > 1)
                {
                    textbox.Text = textbox.Text.Remove(textbox.Text.Length - 1,1);
                }
            }
            textbox.SelectionStart = textbox.Text.Length;
        }
        private static string lastText = string.Empty;
        private void TextBox_KeyUp(object sender,KeyEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            //判断按键是不是要输入的类型。
            if (textBox.Text != "" && textBox.Text != string.Empty && textBox.Text != lastText)
            {
                if (e.Key == Key.Decimal)
                {

                }
                //0-9
                else if (((int)e.Key < 34 || (int)e.Key > 43) && ((int)e.Key < 74 || (int)e.Key > 83))
                {
                    textBox.Text = textBox.Text.Remove(textBox.Text.Length - 1,1);
                }
            }

            textBox.SelectionStart = textBox.Text.Length;
            lastText = textBox.Text;
            e.Handled = true;
        }

        /// <summary>
        /// 搜索
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void searchTextBox_TextChanged(object sender,RoutedEventArgs e)
        {
            string searchText = searchTextBox.Text;
            TabControl tabControl = GetVisualChild<TabControl>(maingrid);
            if (tabControl != null)
            {
                foreach (TabItem tabItem in tabControl.Items)
                {
                    if (tabItem.IsSelected)
                    {
                        string header = tabItem.Header.ToString();
                        DataGrid dataGrid = tabItem.Content as DataGrid;
                        if (searchText != null && searchText != "")
                        {
                            for (int i = 0; i < dataGrid.ItemContainerGenerator.Items.Count - 1; i++)
                            {
                                dataGrid.ScrollIntoView(dataGrid.Items[i]);
                                DataGridRow dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                if (dgv == null)
                                {
                                    dataGrid.UpdateLayout();
                                    dataGrid.ScrollIntoView(dataGrid.Items[i]);
                                    dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                }
                                bool showdgv = false;
                                DataRow dr = (dgv.Item as DataRowView).Row;
                                for (int j = 0; j < dr.ItemArray.Length; j++)
                                {
                                    dgv.UpdateLayout();
                                    DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(dgv);
                                    DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(j);
                                    string cellcontent = dr[j].ToString().Trim();
                                    if (cellcontent.ToLower().Contains(searchText.ToLower()))
                                    {
                                        cell.Background = new SolidColorBrush(Colors.Orange);
                                        showdgv = true;
                                    }
                                    else
                                    {
                                        cell.Background = null;
                                    }
                                }
                                if (showdgv)
                                {
                                    dgv.Visibility = Visibility.Visible;
                                }
                                else
                                {
                                    dgv.Visibility = Visibility.Collapsed;
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < dataGrid.ItemContainerGenerator.Items.Count - 1; i++)
                            {
                                DataGridRow dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                if (dgv == null)
                                {
                                    dataGrid.UpdateLayout();
                                    dataGrid.ScrollIntoView(dataGrid.Items[i]);
                                    dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                }
                                dgv.Visibility = Visibility.Visible;
                                DataRow dr = (dgv.Item as DataRowView).Row;
                                for (int j = 0; j < dr.ItemArray.Length; j++)
                                {
                                    dgv.UpdateLayout();
                                    DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(dgv);
                                    DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(j);
                                    cell.Background = null;
                                }
                            }
                        }
                    }
                }
            }
        }

        #region 辅助函数
        /// <summary>
        /// 获取父可视对象中第一个指定类型的子可视对象
        /// </summary>
        /// <typeparam name="T">可视对象类型</typeparam>
        /// <param name="parent">父可视对象</param>
        /// <returns>第一个指定类型的子可视对象</returns>
        public static T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent,i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }

        /// <summary>
        /// 父控件+控件名找到子控件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public T GetChildObject<T>(DependencyObject obj,string name) where T : FrameworkElement
        {
            DependencyObject child = null;
            T grandChild = null;
            for (int i = 0; i <= VisualTreeHelper.GetChildrenCount(obj) - 1; i++)
            {
                child = VisualTreeHelper.GetChild(obj,i);
                if (child is T && (((T)child).Name == name || string.IsNullOrEmpty(name)))
                {
                    return (T)child;
                }
                else
                {
                    grandChild = GetChildObject<T>(child,name);
                    if (grandChild != null)
                        return grandChild;
                }
            }
            return null;
        }


        #endregion
    }
}
