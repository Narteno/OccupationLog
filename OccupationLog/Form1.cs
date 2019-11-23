using System;
using System.Collections.Generic;
using System.Threading;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using _Word = Microsoft.Office.Interop.Word;
using _Excel = Microsoft.Office.Interop.Excel;

namespace OccupationLog
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public bool WebDriverIsOpen = false;
        public Label lb;
        public string NameOfPrepod;
        public string NameOfTheDiscipline;
        List<string> AllListLab = new List<string>();
        List<string> AllListLec = new List<string>();
        List<string> AllListPr = new List<string>();
        IWebDriver driver;
        ChromeOptions options;
        ChromeDriverService cds;
        private void AddToList(List<string> ls, _Word.Cell cell, int NumberStart, int NumberEnd)
        {
            for (int i = NumberStart; i < NumberEnd; i++)
            {
                ls[i] += ";" + cell.Range.Text.Substring(9,1);
            }
        }
        private void Button1_Click(object sender, EventArgs e) // Поиск преподавателя
        {
            try
            {
                panel1.Controls.Clear();
                if (tb_family_teacher.Text == "")
                {
                    MessageBox.Show("Пожалуйста, заполните необходимые поля", "Ошибка ввода");
                    return;
                }
                int flag = 1, x = 10, y = 10;
                progressBar1.Visible = true;
                options = new ChromeOptions();
                bool FIO = true;
                cds = ChromeDriverService.CreateDefaultService();
                cds.HideCommandPromptWindow = true;
                options.AddArguments("headless");
                driver = new ChromeDriver(cds, options);
                WebDriverIsOpen = true;
                driver.Url = @"http://schedule.tsu.ru/teachers_schedule";
                driver.FindElement(By.XPath(@".//div[@id='teachers_search']/form/input[@name='teachers_surname']")).SendKeys(tb_family_teacher.Text);
                driver.FindElement(By.XPath(@".//div[@id='teachers_search']/form/input[@type='submit']")).Click();
                var links = driver.FindElements(By.XPath(@".//div[@id='teachers_search_result']/div[@class='result-tr']"));
                progressBar1.Step = 50 / (links.Count - 1);
                List<string> URLPhoto = new List<string>();
                foreach (IWebElement link in links)
                {
                    if (!FIO) // чтоб не брать первую строчку "пустышку"
                    {
                        this.lb = new Label()
                        {

                            Name = "tb_serach" + flag.ToString(),
                            Location = new Point(x, y),
                            AutoSize = false,
                            Font = new Font("Times New Roman", 14, FontStyle.Underline),
                            Size = new Size(300, 20),
                            Text = link.Text.Replace("Смотреть расписание", ""),
                        };
                        lb.Click += Label_Click;
                        panel1.Controls.Add(this.lb);
                        var PhotoSearch = link.FindElement(By.TagName("a")).GetAttribute("href");
                        lb.Tag = PhotoSearch; URLPhoto.Add(PhotoSearch);
                        if (flag % 2 == 1)
                            x += 320;
                        else
                        {
                            x = 10; y += 210;
                        }
                        flag++;
                        progressBar1.PerformStep();
                        Thread.Sleep(1);
                    }
                    FIO = false;
                }
                progressBar1.PerformStep();
                flag = 1; x = 67; y = 40;
                for (int i = 0; i < URLPhoto.Count; i++)
                {
                    driver.Url = URLPhoto[i];
                    var photo = driver.FindElement(By.XPath(@".//div[@class='teacher-avatar']/img"));
                    panel1.Controls.Add(new PictureBox()
                    {
                        Name = "pb_teacher" + flag.ToString(),
                        Size = new System.Drawing.Size(150, 150),
                        Location = new Point(x, y),
                        ImageLocation = photo.GetAttribute("src"),
                        SizeMode = PictureBoxSizeMode.StretchImage
                    });
                    if (flag % 2 == 1)
                        x += 320;
                    else
                    {
                        x = 67; y += 210;
                    }
                    flag++;
                    progressBar1.PerformStep();
                    Thread.Sleep(1);
                }
                label5.Visible = true;
                lb = null;
                progressBar1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Возник сбой программы", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                driver.Quit();
                WebDriverIsOpen = false;
            }
        }
        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }
        public string ReturnTime(int NumberOfLesson)
        {
            string TimeOfLesson = "";
            switch(NumberOfLesson)
            {
                case 0: TimeOfLesson = "8:45 - 10:20"; break;
                case 1: TimeOfLesson = "10:35 - 12:10"; break;
                case 2: TimeOfLesson = "12:25 - 14:00"; break;
                case 3: TimeOfLesson = "14:45 - 16:20"; break;
                case 4: TimeOfLesson = "16:35 - 18:10"; break;
                case 5: TimeOfLesson = "18:25 - 20:00"; break;
            }
            return TimeOfLesson;
        }
        private void Label_Click(object sender, EventArgs e)
        {
            var c = GetAll(panel1, typeof(Label));
            foreach(Label lab in c)
            {
                lab.ForeColor = SystemColors.ControlText;
            }
            this.lb = (sender as Label);
            lb.ForeColor = Color.Green;
            NameOfPrepod = lb.Text;
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(WebDriverIsOpen)
            driver.Quit();
        }
        private void Button2_Click(object sender, EventArgs e) // Загрузка РПД
        {
            _Word.Application word = new _Word.Application();
            _Word.Document doc = new _Word.Document();

            int GetTblIndex = 0, LecStart = 0, LecEnd = 0, LabStart = 0, LabEnd = 0, PrStar = 0, PrEnd = 0; 
            Object confConv = false;
            Object readOnly = false;
            Object saveChanges = false;
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = ".docx|*.docx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //File.SetAttributes(openFileDialog.FileName, FileAttributes.);
                    string direct = openFileDialog.FileName;
                    label4.Text = direct.Split('\\').Last();
                    Object filename = direct;
                    doc = word.Documents.Open(ref filename, ref confConv, ref readOnly);
                    word.Visible = false;
                    _Word.Tables tbl = doc.Tables;
                    for (int i = 1; i < tbl.Count + 1; i++)
                    {
                        _Word.Range range = tbl[i].Cell(1, 1).Range;
                        if (range.Text.Substring(0, 27) == "Наименование разделов и тем")
                        {
                            GetTblIndex = i;
                            break;
                        }
                    }
                    object begCell1 = tbl[GetTblIndex].Cell(1, 3).Range.Start;
                    object endCell1 = tbl[GetTblIndex].Cell(2, 5).Range.End;
                    _Word.Range ranga = doc.Range(ref begCell1, ref endCell1);
                    ranga.Select();
                    word.Selection.Cells.Merge();
                    foreach (_Word.Row r in tbl[GetTblIndex].Rows)
                    {
                        if (r.Cells[1].Range.Text.Substring(0, 4) == "Тема")
                        {
                            if (r.Cells[3].Range.Text != "\r\a")
                            {
                                AllListLec.Add((r.Cells[3].Range.Text + ";" + r.Cells[1].Range.Text).Replace("\a", "").Replace("\r", ""));
                                LecEnd++;
                            }
                            if (r.Cells[4].Range.Text != "\r\a")
                            {
                                AllListLab.Add((r.Cells[4].Range.Text + ";" + r.Cells[1].Range.Text).Replace("\a", "").Replace("\r", ""));
                                LabEnd++;
                            }
                            if (r.Cells[5].Range.Text != "\r\a")
                            {
                                AllListPr.Add((r.Cells[5].Range.Text + ";" + r.Cells[1].Range.Text).Replace("\a", "").Replace("\r", ""));
                                PrEnd++;
                            }
                        }
                        else if (r.Cells[1].Range.Text.Substring(0, 5) == "Итого")
                        {
                            AddToList(AllListLec, r.Cells[1], LecStart, LecEnd);
                            AddToList(AllListLab, r.Cells[1], LabStart, LabEnd);
                            AddToList(AllListPr, r.Cells[1], PrStar, PrEnd);
                            LecStart = LecEnd; LabStart = LabEnd; PrStar = PrEnd;
                        }
                    }
                    object findText = "Код и наименование дисциплины";
                    Object unit = _Word.WdUnits.wdStory;
                    Object extend = _Word.WdMovementType.wdMove;
                    Object count = 1;
                    word.Selection.HomeKey(ref unit, ref extend);
                    word.Selection.Find.Execute(ref findText);
                    unit = _Word.WdUnits.wdLine;
                    word.Selection.MoveDown(ref unit, ref count, ref extend);
                    word.Selection.HomeKey(ref unit, ref extend);
                    extend = _Word.WdMovementType.wdExtend;
                    word.Selection.EndKey(ref unit, ref extend);
                    NameOfTheDiscipline = word.Selection.Text;
                    int maxLastIndexOf = -1;
                    for (int i = 0; i < 10; i++)
                    {
                        int counts = NameOfTheDiscipline.LastIndexOf(Convert.ToChar(Convert.ToString(i).Substring(0, 1)));
                        if (counts > maxLastIndexOf)
                            maxLastIndexOf = counts;
                    }
                    NameOfTheDiscipline = NameOfTheDiscipline.Substring(maxLastIndexOf + 2, (NameOfTheDiscipline.LastIndexOf('.') - maxLastIndexOf - 2));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Возник сбой программы",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            finally
            {
                doc.Close(ref saveChanges);
                word.Quit(ref saveChanges);
            }
        }        
        private void Button3_Click(object sender, EventArgs e) // Составляем журнал
        {
            int WeekTo=0,WeekStart=0;
            List<int> ls = new List<int>();
            List<string> AllInfo = new List<string>();
            if (lb == null) // ИСПРАВИТЬ НА !=
            {
                if (comboBox1.Text != "Выбрать семестр")
                {
                    try
                    {
                        // число, месяц - время - №группы - isлекция
                        driver = new ChromeDriver(); WebDriverIsOpen = true;
                        if (comboBox1.Text == "1" || comboBox1.Text == "3" || comboBox1.Text == "5" || comboBox1.Text == "7")
                        {
                            WeekTo = 18; WeekStart = 1;
                            //driver.Url = @"" + lb.Tag + "&week_num=1";
                        }
                        else if (comboBox1.Text == "2" || comboBox1.Text == "4" || comboBox1.Text == "6" || comboBox1.Text == "8")
                        {
                            WeekTo = 44; WeekStart = 24;
                            //driver.Url = @"" + lb.Tag + "&week_num=24";
                        }
                        for (int week = WeekStart; week < WeekTo; week++)
                        {
                            //driver.Url = @"" + lb.Tag + "&week_num=" + week; ВЕРНУТЬ!!!!
                            driver.Url = @"http://schedule.tsu.ru/teachers_schedule?teach_id=15513&week_num=" + week;
                            var GetDate = driver.FindElement(By.XPath(@".//div[@class='schedule-print-block']/span[@class='schedule-info-week']"));
                            var date = GetDate.Text.Substring(GetDate.Text.Length - 23, 23).Substring(0, 10).Replace(".", "/");
                            DateTime d = DateTime.Parse(date);
                            var GetDay = driver.FindElements(By.CssSelector(".weekday_line"));
                            for (int j = 0; j < 6; j++)
                            {
                                var tryCSS = GetDay[j].FindElements(By.CssSelector(".lessons_cell"));
                                for (int k = 0; k < 7; k++)
                                {
                                    try
                                    {
                                        var GetName = tryCSS[k].FindElement(By.TagName("label")).Text;//(By.XPath("./div[@class='all_lessons_info']/div[class='one_lesson_info']/label")).Text;
                                        if (GetName == NameOfTheDiscipline)
                                        {
                                            var ProverkaLekcii = tryCSS[k].FindElement(By.ClassName("type_employment")).GetAttribute("style");
                                            int pr = ProverkaLekcii.IndexOf("rgb(255, 0, 0)"); // лекции
                                            int pr2 = ProverkaLekcii.IndexOf("rgb(41, 109, 144)"); // практики
                                            int pr3 = ProverkaLekcii.IndexOf("rgb(46, 196, 228)"); // лабораторные
                                            string TypeOfLesson = "";
                                            if (pr != -1) TypeOfLesson = "л";
                                            else if (pr2 != -1) TypeOfLesson = "п";
                                            else if (pr3 != -1) TypeOfLesson = "лаб";
                                            var WhatIsGroup = tryCSS[k].FindElement(By.ClassName("groups")).Text;
                                            if (WhatIsGroup.Substring(0, 2) == "18")
                                            {
                                                d = d.AddDays(j);
                                                AllInfo.Add(d.ToString().Substring(0, d.ToString().Length - 8) + "," + ReturnTime(k) + "," + WhatIsGroup + "," + TypeOfLesson);
                                                d = DateTime.Parse(date);
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        int flazhok = 0;
                        List<string> proverka = new List<string>();
                        string Semestr = comboBox1.Text;
                        for (int i = 0; i < AllListLec.Count; i++)
                        {
                            string[] pr = new string[2];
                            pr = AllListLec[i].Split(';');
                            if (pr[0] != "" && pr[2] == Semestr)
                            {
                                for (int j = 0; j < Convert.ToInt32(pr[0]); j++)
                                {
                                    proverka.Add(pr[1]);
                                    flazhok++;
                                }
                            }
                        }
                        flazhok = 0;
                        List<string> ListOfPrakrics = new List<string>();
                        List<string> ListOfLections = new List<string>();
                        List<string> CountOfGroups = new List<string>();
                        List<string> CountOfGroupsLections = new List<string>();
                        string[] GetGroup = new string[4];
                        for (int i = 0; i < AllInfo.Count; i++)
                        {
                            GetGroup = AllInfo[i].Split(',');
                            if (GetGroup[3] == "п")
                                ListOfPrakrics.Add(AllInfo[i]);
                            else if (GetGroup[3] == "л")
                                ListOfLections.Add(AllInfo[i]);
                        }
                        GetGroup = ListOfPrakrics[0].Split(',');
                        CountOfGroups.Add(GetGroup[2]);
                        bool IsHave = true;
                        for (int i = 1; i < ListOfPrakrics.Count; i++)
                        {
                            GetGroup = ListOfPrakrics[i].Split(',');
                            for (int j = 0; j < CountOfGroups.Count; j++)
                            {
                                if (GetGroup[2] == CountOfGroups[j])
                                {
                                    IsHave = false;
                                    break;
                                }
                            }
                            if (IsHave)
                                CountOfGroups.Add(GetGroup[2]);
                            IsHave = true;
                        }
                        GetGroup = ListOfLections[0].Split(',');
                        CountOfGroupsLections.Add(GetGroup[2]);
                        IsHave = true;
                        for (int i = 1; i < ListOfLections.Count; i++)
                        {
                            GetGroup = ListOfLections[i].Split(',');
                            for (int j = 0; j < CountOfGroupsLections.Count; j++)
                            {
                                if (GetGroup[2] == CountOfGroupsLections[j])
                                {
                                    IsHave = false;
                                    break;
                                }
                            }
                            if (IsHave)
                                CountOfGroupsLections.Add(GetGroup[2]);
                            IsHave = true;
                        }
                        List<string> proverka2 = new List<string>();
                        for (int i = 0; i < AllListPr.Count; i++)
                        {
                            string[] pr = new string[2];
                            pr = AllListPr[i].Split(';');
                            if (pr[0] != "")
                            {
                                for (int j = 0; j < Convert.ToInt32(pr[0]); j++)
                                {
                                    proverka2.Add(pr[1]);
                                    flazhok++;
                                }
                            }
                        }
                        for (int i = 0; i < AllInfo.Count; i++)
                        {
                            string[] DataOfLesson = new string[4];
                            DataOfLesson = AllInfo[i].Split(',');
                            if (DataOfLesson[3] == "л")
                            {
                                if (proverka[0] == proverka[1])
                                {
                                    for (int j = 0; j < CountOfGroupsLections.Count; j++)
                                    {
                                        AllInfo[i + j] = proverka[0] + "," + AllInfo[i + j];
                                    }
                                    proverka.RemoveAt(0); proverka.RemoveAt(0);
                                    i += CountOfGroupsLections.Count - 1;

                                }
                                else
                                {
                                    for (int j = 0; j < CountOfGroupsLections.Count; j++)
                                    {
                                        AllInfo[i + j] = proverka[0] + "\n" + proverka[1] + "," + AllInfo[i + j];
                                    }
                                    proverka.RemoveAt(0); proverka.RemoveAt(0);
                                    i += CountOfGroupsLections.Count - 1;
                                }
                            }
                            else if (DataOfLesson[3] == "п")
                            {
                                if (proverka2[0] == proverka2[1])
                                {
                                    for (int j = 0; j < CountOfGroups.Count; j++)
                                    {
                                        AllInfo[i + j] = proverka2[0] + "," + AllInfo[i + j];
                                    }
                                    proverka2.RemoveAt(0); proverka2.RemoveAt(0);
                                    i += CountOfGroups.Count - 1;
                                }
                                else
                                {
                                    for (int j = 0; j < CountOfGroups.Count; j++)
                                    {
                                        AllInfo[i + j] = proverka2[0] + "\n" + proverka2[1] + "," + AllInfo[i + j];

                                    }
                                    proverka2.RemoveAt(0); proverka2.RemoveAt(0);
                                    i += CountOfGroups.Count - 1;
                                }
                            }
                        }
                        _Excel.Application xlApp = new _Excel.Application
                        {
                            WindowState = _Excel.XlWindowState.xlMaximized,
                            Visible = true
                        };
                        object misValue = System.Reflection.Missing.Value;
                        _Excel.Range range;
                        _Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
                        _Excel.Worksheet xlWorkSheet = (_Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        range = xlWorkSheet.get_Range("A1", "H82").Cells;
                        range.Font.Size = 11;
                        range.Font.Name = "Times New Roman";
                        range.HorizontalAlignment = 3;
                        range = xlWorkSheet.get_Range("A1", "A5").Cells; range.Font.Bold = true;
                        range = xlWorkSheet.get_Range("A1").Cells; range.FormulaR1C1 = "ФИО";
                        range = xlWorkSheet.get_Range("B1").Cells; range.FormulaR1C1 = NameOfPrepod;
                        range = xlWorkSheet.get_Range("A2").Cells; range.FormulaR1C1 = "Кафедра";
                        range = xlWorkSheet.get_Range("A3").Cells; range.FormulaR1C1 = "Должность";
                        range = xlWorkSheet.get_Range("A4").Cells; range.FormulaR1C1 = "Учёная степень";
                        range = xlWorkSheet.get_Range("A5").Cells; range.FormulaR1C1 = "Дисциплина";
                        range = xlWorkSheet.get_Range("B5").Cells; range.FormulaR1C1 = NameOfTheDiscipline;
                        range = xlWorkSheet.get_Range("B1", "D1").Cells; range.Merge(Type.Missing); range.HorizontalAlignment = 2;
                        range = xlWorkSheet.get_Range("B2", "D2").Cells; range.Merge(Type.Missing); range.HorizontalAlignment = 2;
                        range = xlWorkSheet.get_Range("B3", "D3").Cells; range.Merge(Type.Missing); range.HorizontalAlignment = 2;
                        range = xlWorkSheet.get_Range("B4", "D4").Cells; range.Merge(Type.Missing); range.HorizontalAlignment = 2;
                        range = xlWorkSheet.get_Range("B5", "D5").Cells; range.Merge(Type.Missing); range.HorizontalAlignment = 2;
                        range = xlWorkSheet.get_Range("A1", "D5").Cells;
                        range.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
                        range = xlWorkSheet.get_Range("A10", "H82").Cells;
                        range.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
                        range = xlWorkSheet.get_Range("A10", "H10").Cells; range.Font.Bold = true;
                        range = xlWorkSheet.get_Range("A10").Cells; range.FormulaR1C1 = "Число, месяц";
                        range = xlWorkSheet.get_Range("B10").Cells; range.FormulaR1C1 = "Время";
                        range = xlWorkSheet.get_Range("C10").Cells; range.FormulaR1C1 = "№ группы";
                        range = xlWorkSheet.get_Range("D10").Cells; range.FormulaR1C1 = "Тема занятия";
                        range = xlWorkSheet.get_Range("E10").Cells; range.FormulaR1C1 = "Лекция";
                        range = xlWorkSheet.get_Range("F10").Cells; range.FormulaR1C1 = "Семинарское занятие (практическое, лабораторное)";
                        range = xlWorkSheet.get_Range("G10").Cells; range.FormulaR1C1 = "Консультация";
                        range = xlWorkSheet.get_Range("H10").Cells; range.FormulaR1C1 = "Зачёт/Экзамен";
                        range = xlWorkSheet.get_Range("A1", "A82").Cells; range.ColumnWidth = 16;
                        range = xlWorkSheet.get_Range("B1", "B82").Cells; range.ColumnWidth = 17;
                        range = xlWorkSheet.get_Range("C1", "C82").Cells; range.ColumnWidth = 14;
                        range = xlWorkSheet.get_Range("D1", "D82").Cells; range.ColumnWidth = 50;
                        range = xlWorkSheet.get_Range("E1", "E82").Cells; range.ColumnWidth = 13;
                        range = xlWorkSheet.get_Range("F1", "F82").Cells; range.ColumnWidth = 50;
                        range = xlWorkSheet.get_Range("G1", "G82").Cells; range.ColumnWidth = 16;
                        range = xlWorkSheet.get_Range("H1", "H82").Cells; range.ColumnWidth = 16;
                        int index = 0;
                        string[] part = new string[4];
                        for (int i = 0; i < AllInfo.Count; i++)
                        {
                            part = AllInfo[i].Split(',');
                            index = i + 11;
                            range = range = xlWorkSheet.get_Range("A" + index.ToString()).Cells; range.FormulaR1C1 = part[1];
                            range = range = xlWorkSheet.get_Range("B" + index.ToString()).Cells; range.FormulaR1C1 = part[2];
                            range = range = xlWorkSheet.get_Range("C" + index.ToString()).Cells; range.FormulaR1C1 = part[3];
                            range = range = xlWorkSheet.get_Range("D" + index.ToString()).Cells; range.FormulaR1C1 = part[0];
                            if (part[4] == "л")
                            {
                                range = range = xlWorkSheet.get_Range("E" + index.ToString()).Cells; range.FormulaR1C1 = "2";
                            }
                            else
                            {
                                range = range = xlWorkSheet.get_Range("F" + index.ToString()).Cells; range.FormulaR1C1 = "2";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Возник сбой программы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        driver.Quit();
                        WebDriverIsOpen = false;
                    }
                }
                else MessageBox.Show("Вам нужно выбрать семестр");
            }
            else MessageBox.Show("Вам нужно подтвердить выбор преподавателя");
        }
    }
}
