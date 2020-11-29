
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Reflection;

using System.Text.RegularExpressions;

namespace Excelavg
{
    class Program
    {
        static string excelname;
        const string extension = ".xlsx";
        //const string extension = ".csv";
        //const string extension = ".xls";

        const int minRow = 1;
        const int maxRow = 999;/////////////

        const int minColumn = 1;
        const int maxColumn = 99;///////////


        static int Main(string[] args)
        {
            int startRow = minRow;
            int endRow = maxRow;
            int startRow1 = 0;
            int endRow1 = 0;


            int leftColumn = minColumn;
            int rightColumn = maxColumn;

            int maxFile = 99;//一个文件夹最大文件数量
            bool ifNewExcel = true;


            //get directory info -----------------------------------------------------------------
            DirectoryInfo directoryInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);


            //get xlFileFormat -------------------------------------------------------------------
            XlFileFormat xlFileFormat = XlFileFormat.xlCurrentPlatformText;
            if (extension == ".xlsx")
            {
                xlFileFormat = XlFileFormat.xlOpenXMLWorkbook;
            }
            else if (extension == ".csv")
            {
                xlFileFormat = XlFileFormat.xlCSV;
            }
            else if (extension == ".xls")
            {
                Excel.Application Application = new Excel.Application();
                string Version = Application.Version;

                if (Convert.ToDouble(Version) < 12)//use Excel 97-2003
                {
                    xlFileFormat = XlFileFormat.xlWorkbookNormal;//-4143
                }
                else//use excel 2007 or later
                {
                    xlFileFormat = XlFileFormat.xlExcel8;//56
                }
                xlFileFormat = XlFileFormat.xlWorkbookNormal;

                Application.Quit();
            }
            else
            {
                Console.WriteLine("不支持 “.xlsx” “.csv” “.xls” 以外的格式 ！！！");
                Console.ReadKey();
                return 0;
            }


            //get all ".xlsx" --------------------------------------------------------
            FileInfo[] files = directoryInfo.GetFiles();
            List<FileInfo> listFileInfos = new List<FileInfo>();
            for (int i = 0; i < (files.Length > maxFile ? maxFile : files.Length); i++)
            {
                if (files[i].Extension == extension)
                {
                    listFileInfos.Add(files[i]);
                    Console.WriteLine(listFileInfos.Last());
                }
            }


            //get startRow and endRow from XXXXXX.xlsx --------------------------------------------
            int excelType = 0;
            foreach (var fi in listFileInfos)
            {
                startRow = minRow;
                endRow = maxRow;
                excelname = fi.Name.TrimEnd(extension.ToCharArray());//example = "a .xlsx" >> "b "

                if (!(excelname.Contains(" ")))
                {
                    excelType = excelType <= 0 ? 0 : excelType;
                    continue;
                }
                else
                {
                    if (excelname.Split(' ').Length == 1)//error example = "a .xlsx"
                    {
                        //if (int.TryParse(excelname.Split(' ').First(), out startRow)) 
                        //{
                        //    if (startRow < minRow || startRow > maxRow)
                        //    {
                        //        startRow = minRow;
                        //    }
                        //    break;
                        //}
                        //else
                        //{
                        //    startRow = minRow;
                        //    continue;
                        //}
                    }
                    else if (excelname.Split(' ').Length > 1)
                    {
                        //first
                        if (int.TryParse(excelname.Split(' ').First(), out startRow))
                        {
                            if (startRow < -maxRow || startRow == 0)
                            {
                                excelType = excelType <= 1 ? 1 : excelType;
                                continue;
                            }
                            else if (-maxRow <= startRow && startRow < 0)
                            {
                                //“-a空格”  “-a空格......空格”
                                if (excelname.Last() == ' ')
                                {
                                    excelType = 5;
                                    break;
                                }
                                else
                                {
                                    excelType = excelType <= 2 ? 2 : excelType;
                                    continue;
                                }
                            }
                            else if (minRow <= startRow && startRow <= maxRow)
                            {
                                //last
                                if (excelname.Last() == ' ')
                                {
                                    excelType = 3;
                                    break;
                                }
                                else if (int.TryParse(excelname.Split(' ').Last(), out endRow))
                                {
                                    //“a空格b”  “a空格......空格b”
                                    if (endRow < startRow || endRow > maxRow)
                                    {
                                        excelType = excelType <= 2 ? 2 : excelType;
                                        continue;
                                    }
                                    else
                                    {
                                        //success "a b.xlsx"
                                        excelType = 4;
                                        break;
                                    }
                                }
                                else
                                {
                                    //“a空格”    “a空格......”
                                    excelType = excelType <= 2 ? 2 : excelType;
                                    break;
                                }

                            }

                        }

                    }

                }

            }
            //get results from  XXX XXX.xlsx
            if (excelType <= 2)
            {
                Console.WriteLine("\r\n******    当前文件夹没有找到 “a空格b" + extension +"” 或者 “a空格" + extension + "”类似的文件！    ******");
                Console.WriteLine("\r\n  为了调整开始行，和结束行，您需要修改文件名的【输入格式 1】：\r\n" +
                                  "\r\n  从 [a] 行到 [b] 行, 或者 [a] 行到末尾行, 1 <= [a] <= [b] <= 999：\r\n" +
                                  "\r\n “a空格b" + extension + "” 或者 “a空格" + extension + "”\r\n" +
                                  "\r\n  【输入格式 2】:\r\n" +
                                  "\r\n  从倒数 [-a] 行到末尾行，-999 <= [a] < 0：\r\n" +
                                  "\r\n “a空格" + extension + "”\r\n");
            }
            else if(excelType >= 3)
            {
                Console.WriteLine("\r\n******    恭喜您已经找到文件！    “" +excelname + extension + "”    ******");
                ifNewExcel = false;
            }
            Console.WriteLine("********************************************************\r\n");

            //Input numbers for startRow and endRow
            while (excelType <= 2)
            {
                startRow = 1;
                endRow = 999;
                Console.WriteLine("请按要求输入行数范围：“a空格b” 或者 “a空格” \r\n");
                excelname = Console.ReadLine();
                //' '
                if (!(excelname.Contains(" ")))
                {
                    excelType = excelType <= 0 ? 0 : excelType;
                    continue;
                }
                else
                {   //a
                    if (int.TryParse(excelname.Split(' ').First(), out startRow))
                    {
                        if (startRow < -maxRow || startRow == 0)
                        {
                            continue;
                        }
                        else if (-maxRow <= startRow && startRow < 0)
                        {
                            if (excelname.Last() == ' ')
                            {
                                excelType = 5;
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else if (minRow <= startRow && startRow <= maxRow)
                        {
                            if (excelname.Last() == ' ')
                            {
                                excelType = 3;
                            }
                            //b
                            else if (int.TryParse(excelname.Split(' ').Last(), out endRow))
                            {
                                if (startRow <= endRow && endRow <= maxRow)
                                {
                                    excelType = 4;
                                }
                            }
                        }
                    }
                }

            }

            //adjust startRow + endRow
            if (excelType == 3)
            {
                startRow = int.Parse(excelname.Split(' ').First());
                endRow = maxRow;
            }
            else if (excelType == 4)
            {
                startRow = int.Parse(excelname.Split(' ').First());
                endRow = int.Parse(excelname.Split(' ').Last());
            }
            else if (excelType == 5)
            {
                startRow = int.Parse(excelname.Split(' ').First());
                endRow = maxRow;

                startRow1 = startRow;
                endRow1 = endRow;
            }

            Console.WriteLine("\r\n********  恭喜您成功创建/找到文件：  “" + excelname + extension + "”  ********\r\n" +
                              "\r\n\t  开始行  =  " + string.Format("{0:G}", startRow) +
                              "\t\t 结束行  =  " + string.Format("{0:G}", endRow) +
                              "\r\n\r\n********************************************************");
            //Console.ReadKey();
            //return 0;


            object missing = System.Reflection.Missing.Value;
            //new excel *****************************************************************************
            //Excel.Application excel0 = new Excel.Application();
            //excel0.DisplayAlerts = false;//No warning, overwrite file

            //Microsoft.Office.Interop.Excel.Workbook workbook0 = excel0.Workbooks.Add(missing);
            //workbook0.SaveAs(AppDomain.CurrentDomain.BaseDirectory + excelname + extension,
            //                 missing, missing, missing, missing, missing, XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

            //Microsoft.Office.Interop.Excel.Sheets sheets0 = workbook0.Worksheets;

            //Excel.Worksheet worksheet0 = (Excel.Worksheet)sheets0.get_Item(1);

            //read excel **********************************************************************
            //Excel.Application excel1 = new Excel.Application();

            //Microsoft.Office.Interop.Excel.Workbook workbook1 = excel1.Workbooks.Open(excelPath,
            //        missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            //Microsoft.Office.Interop.Excel.Sheets sheets1 = workbook1.Worksheets;

            //Excel.Worksheet worksheet1 = (Excel.Worksheet)sheets1.get_Item(1);

            //if new excel
            Excel.Application excel0 = new Excel.Application();
            excel0.DisplayAlerts = false;//No warning, overwrite file
            Microsoft.Office.Interop.Excel.Workbook workbook0;
            if (ifNewExcel)
            {
                workbook0 = excel0.Workbooks.Add(missing);
                //第二个参数使用xlWorkbookNormal,则输出的是xls格式    //如果使用的是missing则输出系统中带有的EXCEL支持的格式
                workbook0.SaveAs(AppDomain.CurrentDomain.BaseDirectory + excelname + extension,
                                 xlFileFormat, 
                                 missing, missing, missing, missing, XlSaveAsAccessMode.xlNoChange, missing, missing, missing);
            }
            else
            {
                //read by open
                workbook0 = excel0.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + excelname + extension,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            }
            Microsoft.Office.Interop.Excel.Sheets sheets0 = workbook0.Worksheets;

            Excel.Worksheet worksheet0 = (Excel.Worksheet)workbook0.Worksheets[1];
            //choose blank worksheet0
            for (int i = 1; i <= sheets0.Count; i++)
            {
                worksheet0 = (Excel.Worksheet)workbook0.Worksheets[i];
                Range range = (Range)worksheet0.Cells[1, 1];

                if (range.Value == null)
                {
                    break;//if range.value == null    then begin calculate data of excels
                }
                else
                {
                    if (i == sheets0.Count)//if range.value   then add worksheet
                    {
                        workbook0.Worksheets.Add(missing, workbook0.Worksheets[i], missing, missing);
                        worksheet0 = (Excel.Worksheet)workbook0.Worksheets[i+1];
                        break;
                    }
                }
            }

            //worksheet0.Cells[1, 1] = "Time: "+ DateTime.Now.ToString("MM月dd日 HHmmss");
            worksheet0.Cells[1, 1] = (DateTime.Now.ToString("MM月dd日 HHmmss")).ToString();

            Excel.Application excel1 = new Excel.Application();
            int columnStatistics = 1;
            foreach (var listFile in listFileInfos)
            {
                if (listFile.Name.Equals(excelname+extension))
                {
                    continue;
                }

                Microsoft.Office.Interop.Excel.Workbook workbook1 = excel1.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + listFile.Name,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

                Microsoft.Office.Interop.Excel.Sheets sheets1 = workbook1.Worksheets;

                Excel.Worksheet worksheet1 = (Excel.Worksheet)sheets1.get_Item(1);


                int useRowCount = worksheet1.UsedRange.Rows.Count;
                int useColCount = worksheet1.UsedRange.Columns.Count;

                //calculate row
                //restore the old startRow + endRow
                startRow = startRow1 == 0 ? startRow : startRow1;
                endRow = endRow1 == 0 ? endRow : endRow1;
                bool isNormalRow = true;
                if (excelType == 3)//"a .xlsx"
                {
                    if (startRow <= (useRowCount < maxRow ? useRowCount : maxRow))
                    {
                        endRow = useRowCount < maxRow ? useRowCount : maxRow;
                    }
                    else
                    {
                        isNormalRow = false;
                    }
                }
                else if (excelType == 4)//"a b.xlsx"
                {
                    if (endRow <= (useRowCount < maxRow ? useRowCount : maxRow))
                    {

                    }
                    else
                    {
                        isNormalRow = false;
                    }
                }
                else if (excelType == 5)//"-a .xlsx"
                {
                    if (-startRow <= (useRowCount < maxRow ? useRowCount : maxRow))
                    {
                        startRow = (useRowCount < maxRow ? useRowCount : maxRow) + startRow + 1;
                        endRow = useRowCount < maxRow ? useRowCount : maxRow;
                    }
                    else
                    {
                        isNormalRow = false;
                    }
                }
                else
                {
                    isNormalRow = false;
                }
                if (isNormalRow == false)
                {
                    Console.WriteLine("\r\n 当前表格为： " + listFile.Name + ",\r\n" +
                                      "\r\n 您的开始行，和结束行，超出这个 Excel 总行数:" + " 1 ~ " + useRowCount + " \r\n" +
                                      "\r\n 请根据Excel，调整输入行数范围：“a空格b” 或者 “a空格” \r\n");
                    workbook1.Close();
                    continue;
                }

                //calculate left Column
                bool isDataColumn = true;
                Excel.Range range1_0;
                Excel.Range range1_1;
                for (int col = minColumn; col <= (useColCount<maxColumn ? useColCount : maxColumn); col++)
                {
                    //
                    range1_0 = (Range)worksheet1.Cells[startRow, col];
                    range1_1 = (Range)worksheet1.Cells[endRow, col];

                    if (range1_0.Value == null || range1_1.Value == null)
                    {
                        if (col >= (useColCount < maxColumn ? useColCount : maxColumn))
                        {
                            Console.WriteLine("\r\n当前表格为： " + listFile.Name + ",\r\n" +
                                              " column 1 - 99 have nothing!\r\n");
                            isDataColumn = false;
                        }
                        //workbook1.Close();
                        continue;
                    }
                    else if (Program.IsNumber(range1_0.Value.ToString()) &&
                             Program.IsNumber(range1_1.Value.ToString()))
                    {
                        leftColumn = col;///////////////////////////////////////////////
                    }
                    else
                    {
                        if (col >= (useColCount < maxColumn ? useColCount : maxColumn))
                        {
                            Console.WriteLine("\r\n column 1 - 99 isn't data!\r\n");
                            isDataColumn = false;
                        }
                        //workbook1.Close();
                        continue;
                    }
                    Console.WriteLine("\r\n\r\n leftColumn = col = " + col);
                    //Console.WriteLine("\r\n\r\n 最后一列的头尾数据," + range1_0.Value.ToString() + " , " + range1_1.Value.ToString());
                    //workbook1.Close();
                    break;
                }
                //only test left Column is Data？
                if (isDataColumn == false)
                {
                    //Console.ReadLine();
                    workbook1.Close();
                    continue;
                }

                //calculate right Column
                for (int col = useColCount < maxColumn ? useColCount : maxColumn; col > leftColumn; col--)
                {
                    //
                    range1_0 = (Range)worksheet1.Cells[startRow, col];
                    range1_1 = (Range)worksheet1.Cells[endRow, col];

                    if (range1_0.Value == null || range1_1.Value == null)
                    {
                        //workbook1.Close();
                        continue;
                    }
                    else
                    {
                        if (Program.IsNumber(range1_0.Value.ToString()) &&
                            Program.IsNumber(range1_1.Value.ToString()))
                        {
                            rightColumn = col;///////////////////////////
                        }
                        else
                        {
                            //workbook1.Close();
                            continue;
                        }
                    }

                    Console.WriteLine("\r\n\r\n rightColumn = col = " + col);
                    //workbook1.Close();
                    break;
                }


                //add data to statistics
                columnStatistics++;
                {
                    Range range = (Range)(worksheet0.Cells[columnStatistics, 1]);
                    range.Value = listFile.Name;
                }
                for (int i = leftColumn; i <=rightColumn; i++)
                {
                    var data = 0.0;
                    Range range = (Range)(worksheet0.Cells[leftColumn, startRow]);

                    for (int j = startRow; j <= endRow; j++)
                    {
                        range = (Range)(worksheet1.Cells[j, i]);
                        data += range.Value;
                    }

                    range = (Range)worksheet0.Cells[columnStatistics, i];
                    range.Value = data / (endRow - startRow + 1);
                }

                //just close workbook1
                workbook1.Save();
                workbook1.Close();
            }
            excel1.Quit();


            workbook0.Save();
            workbook0.Close();
            excel0.Quit();

            Console.ReadKey();
            return 0;
        }


        public void newExcel()
        {
            object missing = System.Reflection.Missing.Value;
            //new excel
            Excel.Application excel0 = new Excel.Application();
            excel0.DisplayAlerts = false;//No warning, overwrite file
            //
            Microsoft.Office.Interop.Excel.Workbook workbook0 = excel0.Workbooks.Add(missing); ;
            //
            Microsoft.Office.Interop.Excel.Sheets sheets0 = workbook0.Worksheets;
            //
            Excel.Worksheet worksheet0 = (Excel.Worksheet)sheets0.get_Item(1);

            string newExcelName = AppDomain.CurrentDomain.BaseDirectory + excelname + DateTime.Now.ToString("yyyy_MM_dd   HH_mm_ss") + extension;
            //第二个参数使用xlWorkbookNormal,则输出的是xls格式
            //如果使用的是missing则输出系统中带有的EXCEL支持的格式
            workbook0.SaveAs(AppDomain.CurrentDomain.BaseDirectory + excelname + extension,
                             missing, missing, missing, missing, missing, XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

        }


        public static bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");

            return !objNotNumberPattern.IsMatch(strNumber) &&
            !objTwoDotPattern.IsMatch(strNumber) &&
            !objTwoMinusPattern.IsMatch(strNumber) &&
            objNumberPattern.IsMatch(strNumber);
        }

        public static dynamic GetDynamicData(Worksheet worksheet, int row, int column)
        {

            return 0;
        }
    }
}
