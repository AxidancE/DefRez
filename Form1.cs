using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using AutoUpdaterDotNET;
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
            comboBox2.SelectedIndex = 0;
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
            Excel.Worksheet Wall = (Excel.Worksheet)xlApp.Worksheets.get_Item(9);//Стенка.швы (переделать на поиск по имени)
            Excel.Worksheet Defectes = (Excel.Worksheet)xlApp.Worksheets.get_Item(10);//Дефекты, поменять номер на тот что был в исходнике

            use_sheet(10);
            
            Excel.Range S_range = xlApp.get_Range("AI6", $"AI{Defectes.UsedRange.Rows.Count}");//"A6", $"A{Defectes.UsedRange.Rows.Count}"
            
            string text = File.ReadAllText(path + @"\mark.cdm", System.Text.Encoding.GetEncoding(1251));
            string b = "";
            string TTC_F = File.ReadAllText(path + @"\Arrayed.txt", System.Text.Encoding.GetEncoding(1251));
            //Console.WriteLine(Defectes.UsedRange.Rows.Count);
            //Console.WriteLine(Defectes.UsedRange.Rows.Count - 5);

            int number_of_defs = Defectes.UsedRange.Rows.Count - 5;
            
            //for (int j = 416; j <= 416; j++)//Defectes.UsedRange.Rows.Count-5
            for (int j = 1; j <= Defectes.UsedRange.Rows.Count - 5; j++)
            {                
                b += $"#[{j}]" + "\n";
                int n = j / (number_of_defs / 100 + 1);
                String s = "Текстую " + n + "% ";
                backgroundWorker1.ReportProgress(n, s); // Отправляем данные в ProgressChanged
            }
            backgroundWorker1.ReportProgress(100, "Завершено.");
            TTC_F = TTC_F.Replace("#array_here", b);

            //for (int i = 416; i <= 416; i++)//i = 1
            for (int i = 1; i <= Defectes.UsedRange.Rows.Count - 5; i++)//i = 1
            {
                Excel.Range Find_in_Cycle = S_range.Find(i);//207 - проверочный
                                                            //Excel.Range Ser_number = Defectes.Cells[Find_in_Cycle.Row, 35];//Номер п|п
                Excel.Range Defect_number = Defectes.Cells[Find_in_Cycle.Row, 6];//Номер дефекта
                Excel.Range Vertical = Defectes.Cells[Find_in_Cycle.Row, 28];//Вертикаль
                Excel.Range Horizon = Defectes.Cells[Find_in_Cycle.Row, 26];//Горизонталь
                Excel.Range Vert_x_orig = Defectes.Cells[Find_in_Cycle.Row, 27];//Расстояние от начала вертикали
                Excel.Range Horiz_y_orig = Defectes.Cells[Find_in_Cycle.Row, 29];//Расстояние от начала горизонтали

                use_sheet(9);
                Excel.Range F_range = Wall.get_Range("B5", $"B{Wall.UsedRange.Rows.Count}");

                Excel.Range V_find = F_range.Find(Vertical);//26
                Excel.Range H_find = F_range.Find(Horizon);//28

                //Excel.Range F_Seam = Wall.Cells[V_find.Row, 2];//
                //Excel.Range S_Seam = Wall.Cells[H_find.Row, 2];//

                //Начальная точка
                //"x" дефекта любой, кроме x2_H
                //"У" дефекта любой, кроме - y1_V
                //Console.WriteLine("I - " + i);
                
                Excel.Range X_Main_orig = Wall.Cells[H_find.Row, 5];// //Темно синий //X_main
                Excel.Range Y_Main_orig = Wall.Cells[V_find.Row, 6];// //Бордовый //Y_main

                Excel.Range X_Additional = Wall.Cells[V_find.Row, 5];
                Excel.Range Y_Additional = Wall.Cells[H_find.Row, 6];

                if(X_Additional.Value2 > X_Main_orig.Value2)//Заменить на "меньше"?
                {
                    X_Main_orig = Wall.Cells[V_find.Row, 5];
                }

                if(Y_Additional.Value2 > Y_Main_orig.Value2)
                {
                    Y_Main_orig = Wall.Cells[H_find.Row, 6];
                }




                double X_Main = Convert.ToInt32(X_Main_orig.Value2);
                double Y_Main = Convert.ToInt32(Y_Main_orig.Value2);
                //Console.WriteLine($"X - {X_Main}; Y - {Y_Main}");

                double Vert_x = Convert.ToInt32(Vert_x_orig.Value2);//27
                double Horiz_y = Convert.ToInt32(Horiz_y_orig.Value2);

                
                ChangeText();
                //text = text.Replace("#" + OriginalName, TextToChange_F);

                string ChangeText()
                {
                    string AllTextered, AllText_T = "", AllText_Fo = "";
                    
                    ChangeText_in_cycle("marker", 1, out string AllText_F);
                    ChangeText_in_cycle("circle", 2, out string AllText_S);
                    
                    if (Vert_x != 0)
                    {
                        ChangeText_in_cycle("Horizon", 3, out AllText_T);
                    }

                    if (Horiz_y != 0)
                    {
                        ChangeText_in_cycle("Vertical", 4, out AllText_Fo);
                    }

                    AllTextered = AllText_F + "\n" + AllText_S + "\n" + AllText_T + "\n" + AllText_Fo;
                    TTC_F = TTC_F.Replace($"#[{i}]", AllTextered);
                    return TTC_F;


                }

                void ChangeText_in_cycle(string TextToChange, int sw_case, out string AllText)
                {
                    AllText = "";
                    string OriginalName = TextToChange;
                    TextToChange = File.ReadAllText(path + @"\" + OriginalName + ".txt", System.Text.Encoding.GetEncoding(1251));

                    double X_Converted = (X_Main + Vert_x) / 100; //5 - 27
                    double Y_Converted = (Y_Main + Horiz_y) / 100; //6 - 29

                    //Console.WriteLine(X_Converted);

                    if (sw_case == 1)
                    {
                        TextToChange = TextToChange.Replace("x = 50.0", "x = " + X_Converted);
                        TextToChange = TextToChange.Replace("y = 46.0", "y = " + Y_Converted);
                        

                        if (Vert_x > 0 && Horiz_y > 0)
                        {
                            TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted + 2));
                            TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted + 2));
                            TextToChange = TextToChange.Replace("dirX = -1", "dirX = 1");

                        }
                        else if (Vert_x > 0 && Horiz_y < 0)
                        {
                            TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted + 2));
                            TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted - 2));
                            TextToChange = TextToChange.Replace("dirX = -1", "dirX = 1");
                        }
                        else if (Vert_x < 0 && Horiz_y < 0)
                        {

                            //Console.WriteLine(TextToChange);
                            TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted - 2));
                            TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted - 2));
                        }
                        else if(Vert_x < 0 && Horiz_y > 0)
                        {
                            TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted - 2));
                            TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted + 2));
                        }
                        else if (Vert_x == 0 || Horiz_y == 0)
                        {
                            TextToChange = TextToChange.Replace("x = 48.0", "x = " + (X_Converted - 2));
                            TextToChange = TextToChange.Replace("y = 44.0", "y = " + (Y_Converted + 2));
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
                        Sw_cased(0);

                        if (Vert_x != 0)
                        {
                            TextToChange = TextToChange.Replace("2317", "" + Math.Abs(Vert_x));
                        }

                        AllText = TextToChange;
                    }
                    if (sw_case == 4)
                    {
                        Sw_cased(1);

                        if (Horiz_y != 0)
                        {
                            TextToChange = TextToChange.Replace("2317", "" + Math.Abs(Horiz_y));
                        }


                        AllText = TextToChange;
                    }
                    string Sw_cased(int Straight)
                    {
                        TextToChange = TextToChange.Replace("11.11", "" + X_Main / 100);
                        TextToChange = TextToChange.Replace("12.22", "" + Y_Main / 100);
                        TextToChange = TextToChange.Replace("13.33", "" + X_Converted);
                        TextToChange = TextToChange.Replace("14.44", "" + Y_Converted);

                        //Vert_x
                        //Horiz_y

                        if (Vert_x == 0 || Horiz_y == 0)
                        {
                            //Console.WriteLine(0);
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
                        else if(Vert_x > 0 && Horiz_y > 0)
                        {
                            //Console.WriteLine(1);
                            if(Straight == 0)
                            {
                                TextToChange = TextToChange.Replace("00.00", "" + 0.0);
                                TextToChange = TextToChange.Replace("01.01", "" + (Horiz_y / 100 + 2.5));
                            }
                            if(Straight == 1)
                            {
                                TextToChange = TextToChange.Replace("00.00", "" + (Vert_x / 100 + 2.5));
                                TextToChange = TextToChange.Replace("01.01", "" + 0.0);
                            }
                        }
                        else if(Vert_x > 0 && Horiz_y < 0)
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
                        else if(Vert_x < 0 && Horiz_y < 0)
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
                        else if(Vert_x < 0 && Horiz_y > 0)
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
            use_sheet(10);
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
                проверкаТаблицыToolStripMenuItem.PerformClick();
            }

            
        }

        private void use_sheet(int sheet_in_use)
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
            //Создавать новый текстовый, а не выводить в сообщении.
            int Errors_finded = 0;
            int Error_find_prev = 0;
            string total_defs;
            List<string> Errored_defect = new List<string> { };
            List<int> Errored_defect_numb = new List<int> { };

            Excel.Worksheet Defectes_try_catch = (Excel.Worksheet)xlApp.Worksheets.get_Item(10);//Дефекты, поменять номер на тот что был в исходнике
            Excel.Worksheet Wall_try_catch = (Excel.Worksheet)xlApp.Worksheets.get_Item(9);//Стенка.швы (переделать на поиск по имени)
            use_sheet(10);
            Excel.Range S_range_try_catch = xlApp.get_Range("AI6", $"AI{Defectes_try_catch.UsedRange.Rows.Count}");//"A6", $"A{Defectes.UsedRange.Rows.Count}"
            
            use_sheet(9);
            Excel.Range F_range_try_catch = Wall_try_catch.get_Range("B5", $"B{Wall_try_catch.UsedRange.Rows.Count}");

            try
            {
                Excel.Range Find_in_Cycle_try_catch = S_range_try_catch.Find(2);//207 - проверочный
                Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, 26];
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
                Console.WriteLine(k);
                progressBar1.Value++;
                Excel.Range Find_in_Cycle_try_catch = S_range_try_catch.Find(k);//207 - проверочный
                try
                {                   
                    Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, 26];//Вертикаль
                    Excel.Range V_find_try_catch = F_range_try_catch.Find(Vertical_try_catch);//26
                    Excel.Range Y_Main_orig_try_catch = Wall_try_catch.Cells[V_find_try_catch.Row, 6];// //Бордовый //Y_main
                }
                catch (Exception)
                {
                    Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, 26];
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
                if(Error_find_prev == 1)
                {
                    Error_find_prev = 0;
                    continue;
                }
                try
                {
                    Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, 28];//Вертикаль
                    Excel.Range V_find_try_catch = F_range_try_catch.Find(Vertical_try_catch);//26
                    Excel.Range Y_Main_orig_try_catch = Wall_try_catch.Cells[V_find_try_catch.Row, 6];// //Бордовый //Y_main
                }
                catch (Exception)
                {
                    Excel.Range Vertical_try_catch = Defectes_try_catch.Cells[Find_in_Cycle_try_catch.Row, 28];
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
                    Console.WriteLine(n);
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
                if(Errors_finded > 40)
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
                button1.PerformClick();
            }
                
        }

        private void создатьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Worksheet Defectes_creating = (Excel.Worksheet)xlApp.Worksheets.get_Item(10);//Дефекты, поменять номер на тот что был в исходнике
            use_sheet(10);
            Excel.Range F_range_creating = Defectes_creating.get_Range("AI6", $"AI{Defectes_creating.UsedRange.Rows.Count}");//"A6", $"A{Defectes.UsedRange.Rows.Count}"
            F_range_creating.Cells[1,1] = String.Format("'0.1");            
            
            int number_of_defs = Defectes_creating.UsedRange.Rows.Count - 5;

            Console.WriteLine(number_of_defs.ToString());

            for (int i = 2; i <= number_of_defs; i++)//i = 1
            {
                F_range_creating.Cells[i, 1] = i;

                int n = (i-1) / (number_of_defs / 100 + 1);
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
    }
}