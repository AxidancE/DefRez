using AutoUpdaterDotNET;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Wall_def
{
    public partial class Form1 : Form
    {
        public int global1;

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            
            AutoUpdater.Synchronous = true;
            AutoUpdater.ShowSkipButton = false;
            AutoUpdater.ShowRemindLaterButton = false;
            AutoUpdater.Start("https://raw.githubusercontent.com/AxidancE/DefRez/main/Version.xml");

            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.WorkerReportsProgress = true;
        }

        private Excel.Application xlApp;
        private Excel.Workbook xlAppBook;
        private Excel.Workbooks xlAppBooks;
        private Excel.Sheets xlSheets;
        private int flagexcelapp = 0;
        private readonly String strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        public static string UserName { get; }
        public static string path = Application.StartupPath.ToString() + @"\textes";

        public void main_prog()
        {
            
            string Var_element = comboBox1.Text + ".Швы";
            if (comboBox1.Text == "Днище_центр")
            {
                Var_element = "Днище.Швы";
            }
            Excel.Worksheet Wall = (Excel.Worksheet)xlApp.Worksheets.get_Item(Var_element);//Поиск по имени листа
            Excel.Worksheet Defectes = (Excel.Worksheet)xlApp.Worksheets.get_Item("Дефекты_1");//Дефекты, поменять номер на тот что был в исходнике

            int Var_vertical,
                Var_horizontal,
                Var_vert_x,
                Var_horiz_y;
            string Var_f_range, Var_f0_range, Vert_string;

            if (comboBox1.SelectedIndex == 0)
            {
                Var_vertical = 26;
                Var_horizontal = 28;
                Var_vert_x = 27;
                Var_horiz_y = 29;

                Var_f_range = "B5";
                Var_f0_range = "B";
            }
            else if (comboBox1.SelectedIndex == 1 || comboBox1.SelectedIndex == 2)
            {
                Var_vertical = 22;
                Var_horizontal = 20;
                Var_vert_x = 21;
                Var_horiz_y = 23;

                Var_f_range = "C6";
                Var_f0_range = "C";
            }
            else
            {
                Var_vertical = 0;
                Var_horizontal = 0;
                Var_vert_x = 0;
                Var_horiz_y = 0;

                Var_f_range = "B6";
                Var_f0_range = "B";
                MessageBox.Show("Ошибка");
            }
            use_sheet("Дефекты_1");
            //Console.WriteLine(Defectes.UsedRange.Rows.Count);
            Excel.Range S_range = xlApp.get_Range("AI6", $"AI{Defectes.UsedRange.Rows.Count}");//"A6", $"A{Defectes.UsedRange.Rows.Count}"

            string text = File.ReadAllText(path + @"\mark.cdm", System.Text.Encoding.GetEncoding(1251));
            string b = "";
            string TTC_F = File.ReadAllText(path + @"\Arrayed.txt", System.Text.Encoding.GetEncoding(1251));
            //Console.WriteLine(Defectes.UsedRange.Rows.Count);
            //Console.WriteLine(Defectes.UsedRange.Rows.Count - 5);

            int number_of_defs = Defectes.UsedRange.Rows.Count - 5;

            //for (int j = 416; j <= 416; j++)//Defectes.UsedRange.Rows.Count-5
            for (int j = 1; j <= number_of_defs; j++)
            {
                b += $"#[{j}]" + "\n";
                int n = j / (number_of_defs / 100 + 1);
                String s = "Текстую " + n + "% ";
                backgroundWorker1.ReportProgress(n, s); // Отправляем данные в ProgressChanged
            }
            backgroundWorker1.ReportProgress(100, "Завершено.");
            TTC_F = TTC_F.Replace("#array_here", b);
            Excel.Range Find_in_Cycle;
            //for (int i = 416; i <= 416; i++)//i = 1
            for (int i = 1; i <= number_of_defs; i++)//i = 1
            {
                
                Find_in_Cycle = S_range.Find(i);//207 - проверочный
                                                //Excel.Range Ser_number = Defectes.Cells[Find_in_Cycle.Row, 35];//Номер п|п
                if (i == 1)
                {
                    Find_in_Cycle = S_range.Find(0.1);
                }

                //Console.WriteLine(Find_in_Cycle.Value2);
                Excel.Range Defect_number = Defectes.Cells[Find_in_Cycle.Row, 6];//Номер дефекта                
                Excel.Range Vertical = Defectes.Cells[Find_in_Cycle.Row, Var_vertical];//Вертикаль
                Excel.Range Horizon = Defectes.Cells[Find_in_Cycle.Row, Var_horizontal];//Горизонталь
                Excel.Range Vert_x_orig = Defectes.Cells[Find_in_Cycle.Row, Var_vert_x];//Расстояние от начала вертикали
                Excel.Range Horiz_y_orig = Defectes.Cells[Find_in_Cycle.Row, Var_horiz_y];//Расстояние от начала горизонтали

                use_sheet(Var_element);
                Excel.Range F_range = Wall.get_Range(Var_f_range, Var_f0_range + Wall.UsedRange.Rows.Count);//c6
                
                Excel.Range V_find = F_range.Find(Vertical, Type.Missing, Type.Missing, 1, Type.Missing, Excel.XlSearchDirection.xlNext, true);//26
                Excel.Range H_find = F_range.Find(Horizon, Type.Missing, Type.Missing, 1, Type.Missing, Excel.XlSearchDirection.xlNext, true);//28 //Type.Missing, s
                //Console.WriteLine(Horizon.Value2);
                //Console.WriteLine(H_find.Value2);

                //Excel.Range F_Seam = Wall.Cells[V_find.Row, 2];//
                //Excel.Range S_Seam = Wall.Cells[H_find.Row, 2];//

                //Начальная точка
                //"x" дефекта любой, кроме x2_H
                //"У" дефекта любой, кроме - y1_V
                //Console.WriteLine("I - " + i);

                Excel.Range Construct_element = Defectes.Cells[Find_in_Cycle.Row, 3];
                
                //Основные привязки для начальной точки.
                Excel.Range X_Main_orig = Wall.Cells[H_find.Row, 5];// //Темно синий //X_main
                Excel.Range Y_Main_orig = Wall.Cells[H_find.Row, 6];// //Бордовый //Y_main //V_find - ориг
                

                //проверено
                Excel.Range XH_Second_orig = Wall.Cells[H_find.Row, 7];
                Excel.Range YH_Second_orig = Wall.Cells[H_find.Row, 8];

                Excel.Range XV_Second_orig = Wall.Cells[V_find.Row, 5];
                Excel.Range YV_Second_orig = Wall.Cells[V_find.Row, 6];

                Excel.Range XV_Second_Additional = Wall.Cells[V_find.Row, 7];
                Excel.Range YV_Second_Additional = Wall.Cells[V_find.Row, 8];


                string Construct_element_str = Construct_element.Value2;
                //Console.WriteLine("--------------------------------");
                if (comboBox1.SelectedIndex == 1)
                {
                    Console.WriteLine(Vertical.Value2);
                    Vert_string = Vertical.Value2;
                    if (Vert_string[0] != 'L')
                    {
                        Console.WriteLine("1)Worked!");
                        

                        X_Main_orig = Wall.Cells[H_find.Row, 7];
                        Y_Main_orig = Wall.Cells[H_find.Row, 8];

                        XH_Second_orig = Wall.Cells[H_find.Row, 5];
                        YH_Second_orig = Wall.Cells[H_find.Row, 6];
                        
                        //Horiz_y_orig.Value2 = -Horiz_y_orig.Value2;
                        //Vert_x_orig.Value2 = -Vert_x_orig.Value2;



                        if (Vert_x_orig.Value2 <= 0)
                        {
                            XV_Second_orig = Wall.Cells[V_find.Row, 7];
                            YV_Second_orig = Wall.Cells[V_find.Row, 8];
                        }
                        else
                        {
                            XV_Second_orig = Wall.Cells[V_find.Row, 5];
                            YV_Second_orig = Wall.Cells[V_find.Row, 6];
                        }
                        
                    }
                }

                Excel.Range X_Additional = Wall.Cells[V_find.Row, 5];
                Excel.Range Y_Additional = Wall.Cells[H_find.Row, 6];

                if (comboBox1.SelectedIndex == 0)
                {
                    X_Main_orig = Wall.Cells[V_find.Row, 7];
                    Y_Main_orig = Wall.Cells[H_find.Row, 8];
                    //if (X_Additional.Value2 > X_Main_orig.Value2)//Заменить на "меньше"?
                    //{
                    //    X_Main_orig = Wall.Cells[V_find.Row, 5];
                    //}

                    //if (Y_Additional.Value2 > Y_Main_orig.Value2)
                    //{
                    //    Y_Main_orig = Wall.Cells[H_find.Row, 6];
                    //}
                }


                double X_Main = Convert.ToInt32(X_Main_orig.Value2);
                double Y_Main = Convert.ToInt32(Y_Main_orig.Value2);
                if(comboBox1.SelectedIndex == 1)
                {
                    if(X_Main == 0)
                    {
                        X_Main = 1;
                    }
                    if(Y_Main == 0)
                    {
                        Y_Main = 1;
                    }
                    
                    
                }
                double XH_Second = Convert.ToInt32(XH_Second_orig.Value2);
                double YH_Second = Convert.ToInt32(YH_Second_orig.Value2);

                double XV_Second_Add = Convert.ToInt32(XV_Second_Additional.Value2);
                double YV_Second_Add = Convert.ToInt32(YV_Second_Additional.Value2);

                double XV_Second = Convert.ToInt32(XV_Second_orig.Value2);
                double YV_Second = Convert.ToInt32(YV_Second_orig.Value2);
                //Console.WriteLine($"X - {X_Main}; Y - {Y_Main}");

                double Vert_x = Convert.ToInt32(Vert_x_orig.Value2);//27
                double Horiz_y = Convert.ToInt32(Horiz_y_orig.Value2);

                double dX_V = 0, dY_V = 0, dX_H = 0, dY_H = 0,
                    x0_v = 0, x0_h = 0, y0_v = 0, y0_h = 0;
                //&& asd != "Центральная часть днища"
                
                if (comboBox1.SelectedIndex == 1)
                {
                    if (XV_Second == X_Main && YV_Second == Y_Main)
                    {
                        XV_Second = XV_Second_Add;
                        YV_Second = YV_Second_Add;
                    }

                    X_Main = convert_number(X_Main);
                    Y_Main = convert_number(Y_Main);
                    XH_Second = convert_number(XH_Second);
                    YH_Second = convert_number(YH_Second);
                    XV_Second = convert_number(XV_Second);
                    YV_Second = convert_number(YV_Second);
                    Vert_x = convert_number(Vert_x);
                    Horiz_y = convert_number(Horiz_y);

                    Vert_x = -Vert_x; //Инверсия дефектов по горизонту

                    //x0|y0 - определяются через excel; || вводить (x1, y1)
                    //xH_0 | yV_0 - аналогично; (xd, yd)
                    //x0_v | y0_h - определение расстояния прямой проходящей по СВШ
                    //IX.X | IX.Y - конечная точка (x2, y2)
                    double const_aV = Vert_x,//От вертикали - горизонтальное расстояние
                           const_aH = Horiz_y,//От горизонта - вертикальное расстояние

                           x0, y0, xH_0, yH_0, xV_0, yV_0,
                           degV,
                           degH,
                           x1_v, y1_v, x1_h, y1_h,
                           x2_v, y2_v, x2_h, y2_h;

                    if (const_aV == 0)
                    {
                        const_aV = -0.01;
                    }
                    //x0 и y0 - начало пересечения свш
                    x0 = X_Main;
                    y0 = Y_Main;

                    //Вертикальный шов
                    xH_0 = XV_Second;
                    yH_0 = YV_Second;

                    //Горизонтальный шов
                    xV_0 = XH_Second;
                    yV_0 = YH_Second;

                    //здесь  поиск углов СВШ

                    degV = Find_deg(x0, y0, xV_0, yV_0);//113.9178
                    degH = Find_deg(x0, y0, xH_0, yH_0);//211.6147

                    //Console.WriteLine($"||||||||||||||| {x0}-{y0}; {xH_0}-{yH_0}");

                    //Console.WriteLine($"{degV} : {degH}\n------------");
                    //Console.WriteLine($"{const_quarter(degV, 1, const_aV)} : {const_quarter(degH, 2, const_aV)}\n------------");
                    //Console.WriteLine($"{const_quarter(degH, 1, const_aV)} : {const_quarter(degH, 2, const_aV)}\n------------");

                    //Нужна эта 4-ка - ага, для чего?
                    x0_h = Rotate_segment(x0, y0, x0 + Math.Abs(const_aV), degH, "x");//-92.0123
                    y0_h = Rotate_segment(x0, y0, x0 + Math.Abs(const_aV), degH, "y");

                    x0_v = Rotate_segment(x0, y0, x0 + Math.Abs(const_aH), degV, "x");//-92.0123
                    y0_v = Rotate_segment(x0, y0, x0 + Math.Abs(const_aH), degV, "y");//208

                    //Переменные для DX и DY размерных выносок
                    dX_V = Rotate_segment(0, 0, 3, degV, "x");
                    dX_H = Rotate_segment(0, 0, 3, degH, "x");
                    dY_H = Rotate_segment(0, 0, 3, degH, "y");
                    dY_V = Rotate_segment(0, 0, 3, degV, "y");

                    //Console.WriteLine($"{x0_h} : {y0_h}");
                    //Console.WriteLine($"{x0_v} : {y0_v}\n------------");

                    //Нахождение координаты относительно длин сторон от которых отступают (прога рисует линии как квадрат)
                    x1_v = Rotate_segment(x0_v, y0_v, x0_v + Math.Abs(const_aV), const_quarter(degV, 1, const_aV), "x");//-98
                    y1_v = Rotate_segment(x0_v, y0_v, x0_v + Math.Abs(const_aV), const_quarter(degV, 1, const_aV), "y");//205 //от вертикала

                    x1_h = Rotate_segment(x0_h, y0_h, x0_h + Math.Abs(const_aH), const_quarter(degH, 2, const_aV), "x");//-92.0123
                    y1_h = Rotate_segment(x0_h, y0_h, x0_h + Math.Abs(const_aH), const_quarter(degH, 2, const_aV), "y");//от горизонта

                    //Console.WriteLine($"{x1_h} : {y1_h}");
                    //Console.WriteLine($"{x1_v} : {y1_v}\n------------");

                    x2_v = Rotate_segment(x1_v, y1_v, x1_v + const_aH, const_quarter(degV, 3, const_aV), "x");//-98
                    y2_v = Rotate_segment(x1_v, y1_v, x1_v + const_aH, const_quarter(degV, 3, const_aV), "y");//205 //от вертикала

                    x2_h = Rotate_segment(x1_h, y1_h, x1_h + const_aV, const_quarter(degH, 3, const_aV), "x");//-92.0123
                    y2_h = Rotate_segment(x1_h, y1_h, x1_h + const_aV, const_quarter(degH, 3, const_aV), "y");//от горизонта

                    Point A = new Point(x1_v, y1_v);
                    Point B = new Point(x2_v, y2_v);
                    Point C = new Point(x1_h, y1_h);
                    Point D = new Point(x2_h, y2_h);

                    Point IX = Intersection(A, B, C, D);

                    //Console.WriteLine("{0} {1}", IX.X, IX.Y);
                    X_Main = IX.X;
                    Y_Main = IX.Y;
                }
                else
                {
                    X_Main = convert_number(X_Main);
                    Y_Main = convert_number(Y_Main);
                }
                
                double convert_number(double a)
                {
                    a /= 100;
                    return a;
                }

                ChangeText();
                //text = text.Replace("#" + OriginalName, TextToChange_F);

                string ChangeText()
                {
                    string AllTextered, AllText_T = "", AllText_Fo = "", AllText_F = "", AllText_S = "";

                    if (comboBox1.SelectedIndex == 0) 
                    {
                        ChangeText_in_cycle("marker", 1, out AllText_F);
                        ChangeText_in_cycle("circle", 2, out AllText_S);

                        if (Vert_x != 0)
                        {
                            
                            ChangeText_in_cycle("Horizon", 3, out AllText_T);
                        }

                        if (Horiz_y != 0)
                        {
                            ChangeText_in_cycle("Vertical", 4, out AllText_Fo);
                        }
                    }
                    else //if(comboBox1.SelectedIndex == 1) // && asd != "Центральная часть днища"
                    {
                        ChangeText_in_cycle("marker", 5, out AllText_F);
                        ChangeText_in_cycle("circle", 6, out AllText_S);
                        Console.WriteLine(Vert_x);
                        Console.WriteLine(Horiz_y);
                        if (Horiz_y != 0 && Math.Abs(Horiz_y) != 0.01)
                        {
                            Console.WriteLine("Worked 0");
                            ChangeText_in_cycle("Horizon", 7, out AllText_T);
                        }

                        if (Vert_x != 0 && Math.Abs(Vert_x) != 0.01)
                        {
                            Console.WriteLine("Worked 1");
                            ChangeText_in_cycle("Vertical", 8, out AllText_Fo);
                        }
                    }

                    AllTextered = AllText_F + "\n" + AllText_S + "\n" + AllText_T + "\n" + AllText_Fo;
                    TTC_F = TTC_F.Replace($"#[{i}]", AllTextered);
                    return TTC_F;//Total Text Changed
                }

                void ChangeText_in_cycle(string TextToChange, int sw_case, out string AllText)
                {
                    AllText = "";
                    string OriginalName = TextToChange;
                    TextToChange = File.ReadAllText(path + @"\" + OriginalName + ".txt", System.Text.Encoding.GetEncoding(1251));
                    double X_Converted, Y_Converted;
                    //Console.WriteLine("_------------------------Main_orig(H)------------------------");
                    //Console.WriteLine(X_Main_orig.Value2);
                    //Console.WriteLine(Y_Main_orig.Value2);
                    //Console.WriteLine("_------------------------Vert_x------------------------");
                    //Console.WriteLine(Vert_x);
                    //Console.WriteLine("_------------------------Horiz_y------------------------");
                    //Console.WriteLine(Horiz_y);
                    if (comboBox1.SelectedIndex == 0)
                    {
                            X_Converted = (X_Main + Vert_x / 100); //5 - 27
                            Y_Converted = (Y_Main + Horiz_y / 100); //6 - 29
                    }
                    else
                    {
                        X_Converted = (X_Main + Vert_x); //5 - 27
                        Y_Converted = (Y_Main + Horiz_y); //6 - 29
                    }
                    

                    //Console.WriteLine("____________________________________");
                    //Console.WriteLine(X_Main);
                    //Console.WriteLine(Y_Main);
                    //Console.WriteLine(X_Converted);
                    //Console.WriteLine(Y_Converted);

                    //Console.WriteLine(X_Converted);

                    if (sw_case == 1)
                    {
                        TextToChange = TextToChange.Replace("x = 50.0", "x = " + X_Converted);
                        TextToChange = TextToChange.Replace("y = 46.0", "y = " + Y_Converted);
                        {
                            if (Vert_x > 0 && Horiz_y > 0)
                            {
                                //Console.WriteLine("x > 0 & y > 0");
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted + 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted + 2));
                                TextToChange = TextToChange.Replace("dirX = -1", "dirX = 1");
                            }
                            else if (Vert_x > 0 && Horiz_y < 0)
                            {
                                //Console.WriteLine("x > 0 & y < 0");
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted + 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted - 2));
                                TextToChange = TextToChange.Replace("dirX = -1", "dirX = 1");
                            }
                            else if (Vert_x < 0 && Horiz_y < 0)
                            {
                                //Console.WriteLine("x < 0 & y < 0");
                                //Console.WriteLine(TextToChange);
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted - 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted - 2));
                            }
                            else if (Vert_x < 0 && Horiz_y > 0)
                            {
                                //Console.WriteLine("x < 0 & y > 0");
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted - 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted + 2));
                            }
                            else if (Vert_x == 0 || Horiz_y == 0)
                            {
                                //Console.WriteLine("x = 0 || y = 0");
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted - 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted + 2));
                            }
                        }
                        TextToChange = TextToChange.Replace("iTextItemParam.s = \"208\"", $"iTextItemParam.s = \"{Convert.ToString(Defect_number.Value2)}\"");
                        AllText = TextToChange;
                    }
                    if (sw_case == 2) //Для округи с штриховкой
                    {
                        TextToChange = TextToChange.Replace("25.0", "" + X_Converted);
                        TextToChange = TextToChange.Replace("999.0", "" + Y_Converted);
                        TextToChange = TextToChange.Replace(
                            "qwe",
                            $"iDocument2D.ksArcByPoint({X_Converted}, {Y_Converted}, 0.25," +
                            $"{X_Converted - 0.25}, {Y_Converted - 0.25}, " +
                            $"{(X_Converted + 0.25)}, {Y_Converted + 0.25}, 1, 1 )");//Аналогично заменить здесь (2 окружности рисуется)

                        TextToChange = TextToChange.Replace(
                            "asd",
                            $"iDocument2D.ksArcByPoint({X_Converted}, {Y_Converted}, 0.25," +
                            $"{X_Converted + 0.25}, {Y_Converted + 0.25}, " +
                            $"{(X_Converted - 0.25)}, {Y_Converted - 0.25}, 1, 1 )");
                        AllText = TextToChange;
                        //TextToChange = TextToChange.Replace("x = 48.0", "x = " + polka);
                    }
                    if (sw_case == 3)
                    {
                        Console.WriteLine("sw_0");
                        Sw_cased(0);
                        if (Vert_x != 0)
                        {
                            TextToChange = TextToChange.Replace("2317", "" + Math.Abs(Vert_x));
                        }
                        AllText = TextToChange;
                    }
                    if (sw_case == 4)
                    {
                        Console.WriteLine("sw_1");
                        Sw_cased(1);
                        if (Horiz_y != 0)
                        {
                            TextToChange = TextToChange.Replace("2317", "" + Math.Abs(Horiz_y));
                        }
                        AllText = TextToChange;
                    }
                    if (sw_case == 5)
                    {
                        TextToChange = TextToChange.Replace("x = 50.0", "x = " + X_Main);
                        TextToChange = TextToChange.Replace("y = 46.0", "y = " + Y_Main);
                        {
                            if (Vert_x > 0 && Horiz_y > 0)
                            {
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Main + 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Main + 2));
                                TextToChange = TextToChange.Replace("dirX = -1", "dirX = 1");
                            }
                            else if (Vert_x > 0 && Horiz_y < 0)
                            {
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Main + 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Main - 2));
                                TextToChange = TextToChange.Replace("dirX = -1", "dirX = 1");
                            }
                            else if (Vert_x < 0 && Horiz_y < 0)
                            {
                                //Console.WriteLine(TextToChange);
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Main - 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Main - 2));
                            }
                            else if (Vert_x < 0 && Horiz_y > 0)
                            {
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Main - 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Main + 2));
                            }
                            else if (Vert_x == 0 || Horiz_y == 0)
                            {
                                TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Main - 2));
                                TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Main + 2));
                            }
                        }
                        TextToChange = TextToChange.Replace("iTextItemParam.s = \"208\"", $"iTextItemParam.s = \"{Convert.ToString(Defect_number.Value2)}\"");
                        AllText = TextToChange;
                    }
                    if (sw_case == 6) //Для округи с штриховкой
                    {
                        TextToChange = TextToChange.Replace("25.0", "" + X_Main);
                        TextToChange = TextToChange.Replace("999.0", "" + Y_Main);
                        TextToChange = TextToChange.Replace(
                            "qwe",
                            $"iDocument2D.ksArcByPoint({X_Main}, {Y_Main}, 0.25," +
                            $"{X_Main - 0.25}, {Y_Main - 0.25}, " +
                            $"{(X_Main + 0.25)}, {Y_Main + 0.25}, 1, 1 )");//Аналогично заменить здесь (2 окружности рисуется)

                        TextToChange = TextToChange.Replace(
                            "asd",
                            $"iDocument2D.ksArcByPoint({X_Main}, {Y_Main}, 0.25," +
                            $"{X_Main + 0.25}, {Y_Main + 0.25}, " +
                            $"{(X_Main - 0.25)}, {Y_Main - 0.25}, 1, 1 )");
                        AllText = TextToChange;
                        //TextToChange = TextToChange.Replace("x = 48.0", "x = " + polka);
                    }
                    if (sw_case == 7)
                    {
                        //Console.WriteLine(7);
                        Sw_cased(0);

                        if (Vert_x != 0 && Vert_x != 1)
                        {
                            TextToChange = TextToChange.Replace("2317", "" + Math.Abs(Horiz_y * 100));
                        }

                        AllText = TextToChange;
                    }
                    if (sw_case == 8)
                    {
                        //Console.WriteLine(8);
                        Sw_cased(1);

                        if (Horiz_y != 0 && Horiz_y != 1)
                        {
                            //Console.WriteLine("1" + Vert_x);
                            TextToChange = TextToChange.Replace("2317", "" + Math.Abs(Vert_x * 100));
                        }

                        AllText = TextToChange;
                    }

                    string Sw_cased(int Straight)
                    {
                        if (comboBox1.SelectedIndex == 0)
                        {
                            TextToChange = TextToChange.Replace("11.11", "" + X_Main);
                            TextToChange = TextToChange.Replace("12.22", "" + Y_Main);
                            TextToChange = TextToChange.Replace("13.33", "" + X_Converted);
                            TextToChange = TextToChange.Replace("14.44", "" + Y_Converted);
                            //Настройки указателей привязок
                            //Vert_x
                            //Horiz_y
                            {
                                if (Vert_x == 0 || Horiz_y == 0)
                                {
                                    //Console.WriteLine(0);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (-2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 2.5);
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x > 0 && Horiz_y > 0)
                                {
                                    //Console.WriteLine(1);
                                    //Console.WriteLine(Horiz_y);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (Horiz_y / 100 + 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Vert_x / 100 + 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x > 0 && Horiz_y < 0)
                                {
                                    //Console.WriteLine(2);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (Horiz_y / 100 - 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Vert_x / 100 + 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x < 0 && Horiz_y < 0)
                                {
                                   //Console.WriteLine(3);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (Horiz_y / 100 - 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Vert_x / 100 - 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x < 0 && Horiz_y > 0)
                                {
                                    //Console.WriteLine(4);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (Horiz_y / 100 + 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Vert_x / 100 - 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                            }
                        }
                        else if (comboBox1.SelectedIndex == 1)
                        {
                            TextToChange = TextToChange.Replace("13.33", "" + X_Main);
                            TextToChange = TextToChange.Replace("14.44", "" + Y_Main);

                            //Vert_x
                            //Horiz_y
                            {
                                if (Straight == 0)
                                {
                                    TextToChange = TextToChange.Replace("iLDimSourceParam.ps = 0", "iLDimSourceParam.ps = 3");
                                    TextToChange = TextToChange.Replace("11.11", "" + x0_h);
                                    TextToChange = TextToChange.Replace("12.22", "" + y0_h);

                                    TextToChange = TextToChange.Replace("00.00", "" + dX_H);
                                    TextToChange = TextToChange.Replace("01.01", "" + dY_H);
                                    //Console.WriteLine($"\n------------\n{dX_H} : {dY_H}");
                                }
                                if (Straight == 1)
                                {
                                    TextToChange = TextToChange.Replace("iLDimSourceParam.ps = 1", "iLDimSourceParam.ps = 3");
                                    TextToChange = TextToChange.Replace("11.11", "" + x0_v);
                                    TextToChange = TextToChange.Replace("12.22", "" + y0_v);

                                    TextToChange = TextToChange.Replace("00.00", "" + dX_V);
                                    TextToChange = TextToChange.Replace("01.01", "" + dY_V);

                                    //Console.WriteLine($"\n------------\n{dX_V} : {dY_V}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("------------");
                            TextToChange = TextToChange.Replace("11.11", "" + X_Main);
                            TextToChange = TextToChange.Replace("12.22", "" + Y_Main);
                            TextToChange = TextToChange.Replace("13.33", "" + X_Converted);
                            TextToChange = TextToChange.Replace("14.44", "" + Y_Converted);

                            //Vert_x
                            //Horiz_y
                            {
                                if (Vert_x == 0 || Horiz_y == 0)
                                {
                                    Console.WriteLine(0);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (-Vert_x / 100 - 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Horiz_y / 100 + 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x > 0 && Horiz_y > 0)
                                {
                                    Console.WriteLine(1);
                                    if (Straight == 0)//y
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (-Vert_x / 100 - 2.5));
                                    }
                                    if (Straight == 1)//x
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Horiz_y / 100 + 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x > 0 && Horiz_y < 0)
                                {
                                    Console.WriteLine(2);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (-Vert_x / 100 - 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Horiz_y / 100 - 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x < 0 && Horiz_y < 0)
                                {
                                    Console.WriteLine(3);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (-Vert_x / 100 + 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Horiz_y / 100 - 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                else if (Vert_x < 0 && Horiz_y > 0)
                                {
                                    Console.WriteLine(4);
                                    if (Straight == 0)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                        TextToChange = TextToChange.Replace("01.01", "" + (-Vert_x / 100 + 2.5));
                                    }
                                    if (Straight == 1)
                                    {
                                        TextToChange = TextToChange.Replace("00.00", "" + (Horiz_y / 100 + 2.5));
                                        TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                                    }
                                }
                                Console.WriteLine("------------\n");
                            }
                        }

                        return (TextToChange);
                    }

                    text = text.Replace("#" + OriginalName, TextToChange);
                }
                int n = i / (number_of_defs / 100 + 1);
                String s = "Выполнение " + n + "% ";
                backgroundWorker1.ReportProgress(n, s); // Отправляем данные в ProgressChanged backgroundWorker1.ReportProgress(100, "Выполнено.");
            }

            text = text.Replace("#Arrayed", TTC_F);
            File.WriteAllText(path + @"\mark_ch.cdm", text, Encoding.GetEncoding(1251));
            use_sheet("Дефекты_1");
            backgroundWorker1.ReportProgress(100, "Выполнено.");
            MessageBox.Show("Готово!",
                "Отчет",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {                
                main_prog();
                
            }
            catch (Exception)
            {
                button1.Enabled = false;
                MessageBox.Show("Найдены несоответствия. \nДля проверки нажмите \"ОК\".", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);//Поменять ОК на 2 кнопки
                CheckTable();
            }
        }

        private void use_sheet(string sheet_in_use)
        {
            //Функция активирует конкретный лист. Сделано чтобы не возникало ошибки присваивания объекта.
            ((Excel.Worksheet)this.xlApp.ActiveWorkbook.Sheets[sheet_in_use]).Select();
        }

        private void открытьExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            открытьExcelToolStripMenuItem.Enabled = false;
            openFileDialog1.Filter = "Excel Files(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm|All files(*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            string pathToXlsx = openFileDialog1.FileName;
            //string pathToXlsx = filename;

            try
            {// Присоединение к открытому приложению Excel (если оно открыто)
                xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                flagexcelapp = 1; // устанавливаем флаг в 1, будем знать что присоединились
            }
            catch
            {
                xlApp = new Excel.Application();// Если нет, то создаём новое приложение
            }
            finally
            {
                xlApp.Workbooks.Open(Path.GetFullPath(pathToXlsx));
                xlApp.Visible = false;
                xlAppBooks = xlApp.Workbooks; // Получаем список открытых книг
                xlAppBook = xlAppBooks[xlAppBooks.Count];
                xlSheets = xlAppBook.Worksheets;
            }
            открытьExcelToolStripMenuItem.Enabled = false;
            закрытьExcelToolStripMenuItem.Enabled = true;
            button1.Enabled = true;
            label1.Visible = true;
            label1.Text = "Excel подключен";
            label1.ForeColor = System.Drawing.Color.Green;
            закрытьExcelToolStripMenuItem.Enabled = true;
            создатьСтолбецToolStripMenuItem.Enabled = true;
            проверкаТаблицыToolStripMenuItem.Enabled = true;
            показатьОкноExcelToolStripMenuItem.Enabled = true;
        }

        private void закрытьExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            закрытьExcelToolStripMenuItem.Enabled = false;

            try
            {
                if (flagexcelapp == 0)
                {
                    xlAppBook.Close(false, false, false);
                    xlApp.Quit();
                    Process[] List;
                    List = Process.GetProcessesByName("EXCEL");
                    foreach (Process proc in List)
                    {
                        proc.Kill();
                    }
                }
                else
                {
                    xlAppBook.Close(false, false, false);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel уже закрыт.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            открытьExcelToolStripMenuItem.Enabled = true;
            label1.Text = "Excel отключен";
            label1.ForeColor = System.Drawing.Color.OrangeRed;
            открытьExcelToolStripMenuItem.Enabled = true;
            создатьСтолбецToolStripMenuItem.Enabled = false;
            проверкаТаблицыToolStripMenuItem.Enabled = false;
            показатьОкноExcelToolStripMenuItem.Enabled = false;
        }

        private void показатьОкноExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                xlApp.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Excel Не открыт");
            }
        }

        private void проверкаТаблицыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckTable();
        }

        private void CheckTable()
        {
            
                
                string Var_element = comboBox1.Text + ".Швы";
                if (comboBox1.Text == "Днище_центр")
            {
                Var_element = "Днище.Швы";
            }
                //Создавать новый текстовый, а не выводить в сообщении.
                int Errors_finded = 0;
                int Error_find_prev = 0;
                string total_defs;
                List<string> Errored_defect = new List<string> { };
                List<int> Errored_defect_numb = new List<int> { };
                string Var_f_range, Var_f0_range;
                int Var_horizontal, Var_vertical;
                if (comboBox1.SelectedIndex == 0)
                {
                    Var_vertical = 28;
                    Var_horizontal = 26;

                    Var_f_range = "B5";
                    Var_f0_range = "B";
                }
                else if (comboBox1.SelectedIndex == 1 || comboBox1.SelectedIndex == 2)
                {
                    Var_vertical = 22;
                    Var_horizontal = 20;

                    Var_f_range = "C6";//5?
                    Var_f0_range = "C";
                }
                else
                {
                    Var_horizontal = 0;
                    Var_vertical = 0;

                    Var_f_range = "B6";
                    Var_f0_range = "B";
                    MessageBox.Show("Ошибка");
                }
                Excel.Worksheet Defectes_try_catch = (Excel.Worksheet)xlApp.Worksheets.get_Item("Дефекты_1");//Дефекты, поменять номер на тот что был в исходнике
                Excel.Worksheet Wall_try_catch = (Excel.Worksheet)xlApp.Worksheets.get_Item(Var_element);//Стенка.швы (переделать на поиск по имени)
                use_sheet("Дефекты_1");
                Excel.Range S_range_try_catch = xlApp.get_Range("AI6", $"AI{Defectes_try_catch.UsedRange.Rows.Count}");//"A6", $"A{Defectes.UsedRange.Rows.Count}"
                use_sheet(Var_element);
                Excel.Range F_range_try_catch = Wall_try_catch.get_Range(Var_f_range, Var_f0_range + Wall_try_catch.UsedRange.Rows.Count);
                try
                {
                    Excel.Range Find_in_Cycle_try_catch = S_range_try_catch.Find(2);//207 - проверочный
                    Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, Var_horizontal];
                }
                catch (Exception)
                {
                    MessageBox.Show("Не найден поисковой столбец. Создание столбца...", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    создатьСтолбецToolStripMenuItem.PerformClick();
                }
                progressBar1.Value = 1;
                progressBar1.Maximum = Defectes_try_catch.UsedRange.Rows.Count - 4;
                int number_of_defs = Defectes_try_catch.UsedRange.Rows.Count - 5;
                //Console.WriteLine(number_of_defs);
                for (int k = 1; k <= number_of_defs; k++)//i = 1
                {
                

                progressBar1.Value++;
                    Excel.Range Find_in_Cycle_try_catch = S_range_try_catch.Find(k);//207 - проверочный
                if (k == 1)
                {
                    Find_in_Cycle_try_catch = S_range_try_catch.Find(0.1);
                }
                try
                    {
                   Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, Var_horizontal];//Вертикаль
                    Excel.Range V_find_try_catch = F_range_try_catch.Find(Vertical_try_catch);//26
                    Excel.Range Y_Main_orig_try_catch = Wall_try_catch.Cells[V_find_try_catch.Row, 6];// //Бордовый //Y_main
                                                                                                          //Console.WriteLine("Worked5");
                    
                }
                    catch (Exception)
                    {
                        Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, Var_horizontal];
                        Errors_finded++;
                        if (Vertical_try_catch.Value2 == null)
                        {
                            Errored_defect.Add("Нет значения");
                        }
                        else
                        {
                            Errored_defect.Add(Vertical_try_catch.Value2);
                        }
                        Errored_defect_numb.Add(k);
                        Error_find_prev = 1;
                    }
                if (Error_find_prev == 1)
                    {
                        Error_find_prev = 0;
                        continue;
                    }
                    try
                    {
                        Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, Var_vertical];//Вертикаль
                        Excel.Range V_find_try_catch = F_range_try_catch.Find(Vertical_try_catch);//26
                        Excel.Range Y_Main_orig_try_catch = Wall_try_catch.Cells[V_find_try_catch.Row, 6];// //Бордовый //Y_main
                    }
                    catch (Exception)
                    {
                        Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, Var_vertical];
                        Errors_finded++;
                        if (Vertical_try_catch.Value2 == null)
                        {
                            Errored_defect.Add("Нет значения");
                        }
                        else
                        {
                            Errored_defect.Add(Vertical_try_catch.Value2);
                        }
                        Errored_defect_numb.Add(k);
                    }

                int n = k / (number_of_defs / 100 + 1);
                    //Console.WriteLine(n);
                    String s = "Проверка " + n + "% ";
                    backgroundWorker1.ReportProgress(n, s); // Отправляем данные в ProgressChanged
                }
                backgroundWorker1.ReportProgress(100, "Проверено."); // Отправляем данные в ProgressChanged
            if (Errors_finded > 0)
                {
                    total_defs = $"Найдено - {Errors_finded} несоответствий.\n";
                    for (int i = 0; i < Errors_finded; i++)
                    {
                        total_defs += $"\n{i + 1}) {Errored_defect_numb[i]} - {Errored_defect[i]}; ";
                    }
                    if (Errors_finded > 10)
                    {
                        //System.IO.File.Create(Application.StartupPath.ToString() + @"\Отчеты");
                        StreamWriter file = new StreamWriter(Application.StartupPath.ToString() + @"\Отчет\Отчёт.txt");
                        file.Write(total_defs);
                        file.Close();

                        MessageBox.Show("Отчет создан в папке \"Отчёт\".", "Отчет");
                    }
                    else
                    {
                        MessageBox.Show(total_defs, "Отчет");
                    }
                }
                else
                {
                    MessageBox.Show("Несоответствий не найдено.", "Отчет", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    button1.Enabled = true;
                }
            
        }

        private void создатьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Worksheet Defectes_creating = (Excel.Worksheet)xlApp.Worksheets.get_Item("Дефекты_1");//Дефекты, поменять номер на тот что был в исходнике
            use_sheet("Дефекты_1");
            Excel.Range F_range_creating = Defectes_creating.get_Range("AI6", $"AI{Defectes_creating.UsedRange.Rows.Count}");//"A6", $"A{Defectes.UsedRange.Rows.Count}"
            F_range_creating.Cells[1, 1] = String.Format("'0.1");
            

            int number_of_defs = Defectes_creating.UsedRange.Rows.Count - 5;

            //Console.WriteLine(number_of_defs.ToString());

            for (int i = 2; i <= number_of_defs; i++)//i = 1
            {
                F_range_creating.Cells[i, 1] = i;

                int n = (i - 1) / (number_of_defs / 100 + 1);
                String s = "Создание " + n + "% ";
                backgroundWorker1.ReportProgress(n, s); // Отправляем данные в ProgressChanged
            }
            backgroundWorker1.ReportProgress(100, "Создано.");
            MessageBox.Show("Столбец создан", "Отчет", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void оПрограммеToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Автор: Мазиков А.С." +
                "\nВерсия: " + strVersion,
                "О программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage; // Меняю данные прогрессбара
            toolStripStatusLabel1.Text = (String)e.UserState; // Меняю значение метки
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            main_prog(); // Вызываем метод с расчетами
            // или прямо тут можно что-то считать =)
        }

        public static double Find_deg(double x_0, double y_0, double xv_1, double yv_1)
        {
            double res_H;
            //x_0 = -87.5451,
            //y_0 = 198.0650,

            //xv_1 = -93.5889,
            //yv_1 = 211.7389,

            //xh_1 = -135.7769,
            //yh_1 = 168.3756,
            res_H = (yv_1 - y_0) / (xv_1 - x_0);
            res_H = RadianToDegree_circle(res_H);
            res_H = Math.Round(res_H, 2);

            if ((xv_1 - x_0) < 0 && (yv_1 - y_0) > 0)
            {
                res_H += 180;
            }
            else if ((xv_1 - x_0) < 0 && (yv_1 - y_0) < 0)
            {
                res_H += 180;
            }
            else if ((xv_1 - x_0) > 0 && (yv_1 - y_0) < 0)
            {
                res_H += 360;
            }

            double RadianToDegree_circle(double angle)
            {
                angle = Math.Atan(angle);
                angle = angle * 180.0 / Math.PI;
                return angle;
            }

            return res_H;
        }

        public static double Rotate_segment(double x_0, double y_0, double x_Rotate, double deg, string axis)
        {
            double x_2, y_2, res_H, res_F;

            deg = DegreeToRadian(deg);
            x_2 = x_0 + (x_Rotate - x_0) * Math.Cos(deg);// - ()
            y_2 = y_0 + (x_Rotate - x_0) * Math.Sin(deg);

            x_2 = Math.Round(x_2, 4);
            y_2 = Math.Round(y_2, 4);

            if (axis == "x")
            {
                return x_2;
            }
            else
            {
                return y_2;
            }
        }

        public static double DegreeToRadian(double angle)
        {
            angle = angle * Math.PI / 180.0;
            return angle;
        }

        //public
        static private Point Intersection(Point A, Point B, Point C, Point D)
        {
            double xo = A.X, yo = A.Y;
            double p = B.X - A.X, q = B.Y - A.Y;

            double x1 = C.X, y1 = C.Y;
            double p1 = D.X - C.X, q1 = D.Y - C.Y;

            double x = (xo * q * p1 - x1 * q1 * p - yo * p * p1 + y1 * p * p1) /
                (q * p1 - q1 * p);
            double y = (yo * p * q1 - y1 * p1 * q - xo * q * q1 + x1 * q * q1) /
                (p * q1 - p1 * q);

            return new Point(x, y);
        }

        //left = x < 0; right = x > 0.
        static public double const_quarter(double income_deg, int step, double X_Left_Right = 0)
        {
            double outcome_deg;
            if (step == 1)
            {
                if (X_Left_Right > 0)
                {
                    outcome_deg = income_deg - 90;
                }
                else
                {
                    outcome_deg = income_deg + 90;
                }
            }
            else if (step == 2)
            {
                if (X_Left_Right > 0)
                {
                    outcome_deg = income_deg + 90;
                }
                else
                {
                    outcome_deg = income_deg - 90;
                }
            }
            else
            {
                outcome_deg = income_deg + 180;
            }
            return outcome_deg;
        }
    }

    internal class Point
    {
        public double X { get; set; }
        public double Y { get; set; }

        public Point()
        {
        }

        public Point(double x, double y)
        {
            X = x;
            Y = y;
        }
    }
}