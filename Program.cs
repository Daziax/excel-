using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace excel汇总
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelOp eo = new ExcelOp();


            Console.WriteLine("请讲文件拖入此处\n");
            string fullname = Console.ReadLine();
            Console.WriteLine("请输入需要汇总的sheet位置(第几个)");
            int sheetnum = Convert.ToInt16(Console.ReadLine());
            Console.WriteLine("系统功能\n1.将前面Sheet汇总到一张\n2.将汇总的表格合并相同项。\n3.退出");
            string option = Console.ReadLine();
            Console.WriteLine("\n正在转换中，请耐心等待。。。");
            while (option == "1" || option == "2")
            {
                if (option == "1")
                {
                    object Nothing = System.Reflection.Missing.Value;
                    Excel.Application app = new Excel.Application();
                    app.Visible = false;

                    Excel.Workbook mybook = app.Workbooks.Open(fullname);

                    eo.CreateSheet(fullname, mybook);

                    Excel.Worksheet mysheet = mybook.Sheets[mybook.Sheets.Count];

                    

                    int i = 0, j = 0;
                    string name = "1", unit = "1";
                    double amount = 1, unitPrice = 1;
                    Excel.Worksheet thesheet = mybook.Worksheets[1];
                    int count = thesheet.UsedRange.Rows.Count;
                    try
                    {
                        for (i = sheetnum; i <= mybook.Sheets.Count - 1; i++)
                        {
                            thesheet = mybook.Worksheets[i];
                            thesheet.Activate();

                            count = thesheet.UsedRange.Rows.Count;
                            for (j = 4; j <= count; j++)
                            {

                                name = thesheet.Cells[j, "B"].Value;
                                if (name == null)
                                    break;
                                unit = thesheet.Cells[j, "C"].Value;
                                amount = thesheet.Cells[j, "J"].Value;
                                unitPrice = thesheet.Cells[j, "H"].Value;

                                eo.WriteToExcel(name, unit, amount, unitPrice, mybook, mysheet);
                            }
                            mybook.Save();
                            Console.WriteLine("  第{0}个：" + count, i);
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("导入第{0}个表格出错" + e.Message + "坐标((i,j)=({1},{2}))" + "当前name为'{3}' 后者为'{4}'", i, i, j, name, thesheet.Cells[j, "B"].Value);
                    }
                    finally
                    {
                        mybook.Save();
                        mybook.Close(false, Type.Missing, Type.Missing);
                        mybook = null;
                        app.Quit();
                        Console.WriteLine("系统功能\n1.将前面Sheet汇总到一张\n2.将汇总的表格合并相同项。\n3.退出");
                        option = Console.ReadLine();

                    }
                }
                else if (option == "2")
                {
                    option =eo.CombineSame(fullname);
                }
            }
            return;

        }
        class ExcelOp
        {
            public string Read(string path)
            {
                StreamReader sr = new StreamReader(path, Encoding.UTF8);
                return sr.ReadToEnd();
            }

            internal void CreateSheet(string FileName, Excel.Workbook workBook)
            {
                //create
                //object Nothing = System.Reflection.Missing.Value;
                //var app = new Excel.Application();
                //app.Visible = false;
                //Excel.Workbook workBook = app.Workbooks.Add(Nothing);


                workBook.Worksheets.Add(Type.Missing, workBook.Worksheets[workBook.Worksheets.Count], 1, Type.Missing);
                ((Worksheet)workBook.Worksheets[workBook.Worksheets.Count]).Name = "汇总";

                workBook.Save();
                Excel.Worksheet worksheet = workBook.Sheets[workBook.Sheets.Count];


                //headline
                worksheet.Cells[1, 1] = "名称";
                worksheet.Cells[1, 2] = "单位";
                //worksheet.Cells[1, 2].numberFormatting = "@";
                worksheet.Cells[1, 3] = "数量";
                worksheet.Cells[1, 4] = "单价";





                //workBook.Close(false, Type.Missing, Type.Missing);


                //app.Quit();

            }
            internal void WriteToExcel(string name, string unit, double amount, double unitPrice, Excel.Workbook mybook, Worksheet mysheet)
            {
                //open
                //mysheet.Activate();
                //get activate sheet max row count
                int maxrow = mysheet.UsedRange.Rows.Count + 1;
                mysheet.Cells[maxrow, 1] = name;
                mysheet.Cells[maxrow, 2] = unit;
                mysheet.Cells[maxrow, 3] = amount;
                mysheet.Cells[maxrow, 4] = unitPrice;

                //mybook.Save();
                //mybook.Close(false, Type.Missing, Type.Missing);
                //mybook = null;
            }
            internal string CombineSame(string fullname)
            {
                string option = null;
                Excel.Application app = new Excel.Application();
                app.Visible = false;
                Excel.Workbook mybook = app.Workbooks.Open(fullname);
                Excel.Worksheet mysheet = mybook.Sheets[mybook.Sheets.Count];

                try
                {
                    string tempName = "";
                    int count = mysheet.UsedRange.Rows.Count, tempCount = 0;
                    for (int i = 1; i <= count; i++)
                    {
                        tempName = mysheet.Cells[i, "A"].Value;
                        tempCount = i;
                        for (int k = i + 1; k <= count; k++)
                        {
                            if (mysheet.Cells[k, "A"].Value == tempName)
                            {
                                mysheet.Cells[tempCount, "C"] = mysheet.Cells[tempCount, "C"].Value+mysheet.Cells[k, "C"].Value;
                                mysheet.Rows[k].Delete();
                                count--;
                            }
                            //if (mysheet.Cells[k, "A"].Value == null)
                            //    break;
                        }
                        Console.WriteLine("已合并{0}个", i);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                finally
                {
                    mybook.Save();
                    mybook.Close(false, Type.Missing, Type.Missing);
                    mybook = null;
                    app.Quit();
                    Console.WriteLine("系统功能\n1.将前面Sheet汇总到一张\n2.将汇总的表格合并相同项。\n3.退出");
                    option = Console.ReadLine();
                }
                return option;





            }
        }
    }
}
