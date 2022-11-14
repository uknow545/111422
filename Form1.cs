using OpenQA.Selenium.Chrome;
using System; 
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Reflection.Emit;
using System.Windows.Forms; 
namespace materialsautopilot
{
    public partial class Form1 : Form
    {
     int SleepTime = 500;
     string fotest = "";
     string fotest2 = "";
     string fotest3 = "";
     string fotest4 = "";
     string fotest5 = "";
     string fotest6 = "";
     string fotest7 = "";
     string fotest8 = "";
     string fotest9 = "";

     private int highestPercentageReached = 0;
     static string datasource = @"10.48.0.200";
     static string database = "DYDC";
     static string username = "sa";
     static string password = "Aa123456@";
     static string connString = @"Data Source=" + datasource + ";Initial Catalog="
                      + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
        public Form1()
        {
         InitializeComponent();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
         this.label2.Visible = true;
         this.label2.Text = this.fotest.ToString();
         this.label2.Visible = true;
         this.label3.Text = this.fotest2.ToString();
         this.label3.Visible = true;
         this.label4.Text = this.fotest3.ToString();
         this.label4.Visible = true;
         this.label5.Text = this.fotest4.ToString();
         this.label5.Visible = true;
         this.label6.Text = this.fotest5.ToString();
         this.label6.Visible = true;
         this.label7.Text = this.fotest6.ToString();
         this.label7.Visible = true;
         this.label8.Text = this.fotest7.ToString();
         this.label8.Visible = true;
         this.label9.Text = this.fotest8.ToString();
         this.label9.Visible = true;
         this.label9.Text = this.fotest8.ToString();
         this.label10.Visible = true;
         this.label10.Text = this.fotest9.ToString();
         progressBar1.PerformStep();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }

        int show2(int n, BackgroundWorker worker, DoWorkEventArgs e)
        {
         int result = 0;
         int percentComplete = 0;
         if (percentComplete > 90)
          { 
          }
         else
         {
          if (percentComplete >= highestPercentageReached)
          {
           highestPercentageReached = percentComplete;
           worker.ReportProgress(percentComplete);
             percentComplete += 10;
          }
           }
            return result;
          }

        private void button1_Click(object sender, EventArgs e)
        {
         InitializeBackgroundWorker();
        }
        private void InitializeBackgroundWorker()
        {
         this.backgroundWorker1 = new BackgroundWorker();
         this.backgroundWorker1.WorkerReportsProgress = true;
         this.backgroundWorker1.DoWork += new DoWorkEventHandler(this.backgroundWorker1_DoWork);
         this.backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
         this.backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
         this.backgroundWorker1.RunWorkerAsync(1);
        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            //this.Iron();
            //e.Result = this.show2((int)e.Argument, worker, e);
            //this.fotest = "鋼筋每公噸牌價下載完成";

            //this.Points();
            //this.fotest2 = "景氣指標下載完成";
            //e.Result = this.show2((int)e.Argument, worker, e);

            //this.Points2();
            //this.fotest3 = "領先指標下載完成";
            //e.Result = this.show2((int)e.Argument, worker, e);

            //this.Points3();
            //this.fotest4 = "同時指標下載完成";
            //e.Result = this.show2((int)e.Argument, worker, e);

            //this.Points4();
            //this.fotest5 = "落後指標下載完成";
            //e.Result = this.show2((int)e.Argument, worker, e);

            //this.Points5();
            //this.fotest6 = "租金指數下載完成";
            //e.Result = this.show2((int)e.Argument, worker, e);

            //this.Points6();
            //e.Result = this.show2((int)e.Argument, worker, e);
            //this.fotest7 = "製造業採購經理人指數(PMI)下載完成";

            //this.Points7();
            //e.Result = this.show2((int)e.Argument, worker, e);
            //this.fotest8 = "非製造業經理人指數(NMI)下載完成";

            this.materials();
            e.Result = this.show2((int)e.Argument, worker, e);
            this.fotest9 = "原物料下載完成";
        }

        private void Iron()
        {
            DateTime today = DateTime.Today;
            if (today.DayOfWeek.ToString() == "Wednesday")
            {
             var tab1 = new ChromeDriver();
             tab1.Navigate().GoToUrl("https://www.google.com.tw/search?q=公噸%2B牌價+豐興&tbs=qdr:w");
             System.Threading.Thread.Sleep(SleepTime);
             string pagesource1 = tab1.PageSource;
             tab1.Quit();
             Iron2(pagesource1);
            }
        }
        private void Iron2(string pageSource)
        {
            string pagehere = pageSource;
            int indextitle = pagehere.IndexOf("<br><h3 class=");
            int i = 0;
            while (pagehere.IndexOf("<br><h3 class=", indextitle + i) != -1)
            {
                indextitle = pagehere.IndexOf("<br><h3 class=", indextitle + i);
                string gettitle = pagehere.Substring(indextitle + 14, 500);
                int startindex = gettitle.IndexOf(">");
                int endindex = gettitle.IndexOf("</h3>");
                string gettitle2 = gettitle.Substring(startindex + 1, (endindex - (startindex) - 1));
                string linkO = pagehere.Substring(indextitle - 2300, 2300);
                int f1dex = linkO.LastIndexOf("<a href");
                string tmp1 = linkO.Substring(f1dex, linkO.Length - f1dex);
                int f2dex = tmp1.IndexOf("http");
                string tmp2 = tmp1.Substring(f2dex, tmp1.Length - f2dex);
                int f3dex = tmp2.IndexOf("\"");
                string tmp3 = tmp2.Substring(0, f3dex);

                if (tmp3.Contains("ctee"))
                {
                    var tab1 = new ChromeDriver();
                    tab1.Navigate().GoToUrl(tmp3);
                    System.Threading.Thread.Sleep(SleepTime);
                    string pagesource1 = tab1.PageSource;
                    
                    tab1.Quit();
                    Iron3(pagesource1);
                    break;
                }
                i++;
            }
        }
        private void Iron3(string Page)
        {
            string pageSource = Page;
            string skeyword = "";
            string skeyword2 = "";
            string skeyword3 = "";
            int start = 0;
            int end = 0;
            int start2 = 0;
            int end2 = 0;
            int start3 = 0;
            int end3 = 0;
            bool deter = false;
            start = pageSource.IndexOf("鋼筋每公噸牌價");
            end = pageSource.IndexOf("元", start);
            skeyword = pageSource.Substring(start, end - start);
            start2 = pageSource.IndexOf("datetime=\"");
            end2 = pageSource.IndexOf("T", start2);
            skeyword2 = pageSource.Substring(start2 + 10, end2 - (start2 + 10));

            start3 = pageSource.IndexOf("廢鋼每公噸收購牌價");
            end3 = pageSource.IndexOf("元", start3);
            skeyword3 = pageSource.Substring(start3, end3 - start3);

            string splitYear = skeyword2;
            string splitYear2 = splitYear.Substring(0, 4);
            int isplitYear2 = Int32.Parse(splitYear2);
            string splitMonth = skeyword2;
            string splitMonth2 = skeyword2.Substring(5, 2);
            string splitDay = skeyword2;
            string splitDay2 = skeyword2.Substring(8, 2);
            string test = skeyword;
            string test2 = string.Empty;
            for (int t1 = 0; t1 < test.Length; t1++)
            {
                if (char.IsDigit(test[t1]))
                {
                    test2 += test[t1]; // iron 20900
                }
            }
            int tmpskeyword = Int32.Parse(test2);

            string test3 = skeyword3;
            string test4 = string.Empty;

            for (int t1 = 0; t1 < test3.Length; t1++)
            {
                if (char.IsDigit(test3[t1]))
                {
                    test4 += test3[t1]; //badiron 10900
                }
            }
            decimal c = 0;
            int basepoint = 17100;
            double basecompare = 0;
            int tmpskeyword2 = Int32.Parse(test4);
            DateTime today = DateTime.Today;
            string rec = today.Date.ToString();
            int a = Int32.Parse(test2);
            int b = Int32.Parse(test4);

            basecompare = a - basepoint;
            basecompare = basecompare / basepoint;
            string d = basecompare.ToString("0.00");
          //  c = double.Parse(d) * 100;

          //   deter = RetreivedPrices(isplitYear2, splitMonth2, splitDay2, skeyword2, "", "new", a, b, 0, "", c, "每週", rec.Substring(0, 10));
           
            if (deter == true)
            {
            }
            else
            {
          //   InsertPrices(isplitYear2, splitMonth2, splitDay2, skeyword2, "", "new", a, b, 0, "", c, "每週", rec.Substring(0, 10));
            }
        }

        private void Points()
        {
            var tab1 = new ChromeDriver();
            tab1.Navigate().GoToUrl("https://index.ndc.gov.tw/n/zh_tw/lightscore#/");
            System.Threading.Thread.Sleep(SleepTime);
            string pagesource1 = tab1.PageSource;
            tab1.Quit();
            Signals(pagesource1, "https://index.ndc.gov.tw/n/zh_tw/lightscore#/");
        }

        /// <summary>
        /// 領先指標
        /// </summary>
        private void Points2()
        {
         var tab1 = new ChromeDriver();
         tab1.Navigate().GoToUrl("https://index.ndc.gov.tw/n/zh_tw/leading#/");
         System.Threading.Thread.Sleep(SleepTime);
         string pagesource1 = tab1.PageSource;
         tab1.Quit();
         LeadPoints(pagesource1, "https://index.ndc.gov.tw/n/zh_tw/leading#/");
        }

        /// <summary>
        /// 同時指標
        /// </summary>
        private void Points3()
        {
            var tab1 = new ChromeDriver();
            tab1.Navigate().GoToUrl("https://index.ndc.gov.tw/n/zh_tw/coincident#/");
            System.Threading.Thread.Sleep(SleepTime);
            string pagesource1 = tab1.PageSource;
            tab1.Quit();
            Coincidence(pagesource1, "https://index.ndc.gov.tw/n/zh_tw/coincident#/");
        }

        /// <summary>
        /// 落後指標
        /// </summary>
        private void Points4()
        {
            var tab1 = new ChromeDriver();
            tab1.Navigate().GoToUrl("https://index.ndc.gov.tw/n/zh_tw/lagged#/");
            System.Threading.Thread.Sleep(SleepTime);
            string pagesource1 = tab1.PageSource;
            tab1.Quit();
            Lagged(pagesource1, "https://index.ndc.gov.tw/n/zh_tw/lagged#/");
        } 

        /// <summary>
        /// 租金指數
        /// </summary>
        private void Points5()
        {
            var tab1 = new ChromeDriver();
            tab1.Navigate().GoToUrl("https://pip.moi.gov.tw/V3/Default.aspx");
            System.Threading.Thread.Sleep(SleepTime);
            string pagesource1 = tab1.PageSource;
            tab1.Quit();
            Rent(pagesource1, "https://pip.moi.gov.tw/V3/Default.aspx");
        }
        private void Points6()
        {
            var tab1 = new ChromeDriver();
            tab1.Navigate().GoToUrl("https://index.ndc.gov.tw/n/zh_tw/PMI#/");
            System.Threading.Thread.Sleep(SleepTime);
            string pagesource1 = tab1.PageSource;
            tab1.Quit();
            PMI(pagesource1, "https://index.ndc.gov.tw/n/zh_tw/PMI#/");
        }
        private void Points7()
        {
            var tab1 = new ChromeDriver();
            tab1.Navigate().GoToUrl("https://index.ndc.gov.tw/n/zh_tw/NMI#/");
            System.Threading.Thread.Sleep(SleepTime);
            string pagesource1 = tab1.PageSource;
            tab1.Quit();
            NMI(pagesource1, "https://index.ndc.gov.tw/n/zh_tw/NMI#/");
        }
        private void materials()
        {
         for (int i = 0; i < 5; i++)
         {
          string[] materialarray = new string[]
              {
              // "PVC 進口價",
              // "冷軋",
              // "熱軋鋼品中鋼盤價",
               "倫敦鋁現貨價",
               "倫敦鎳",
               "混凝土",
               "銅",
               "鋼板"
              };

                string[] tablename = new string[]
                {
                 // "[04.物料走勢_PVC)]",
                  //"[04.物料走勢_冷軋鋼品)]",
                  //"[04.物料走勢_熱軋鋼品)]",
                  "[04.物料走勢_倫敦鋁)]",
                  "[04.物料走勢_倫敦鎳)]",
                  "[04.物料走勢_混凝土(3000PSI)]",
                  "[04.物料走勢_銅)]",
                  "[04.物料走勢_鋼板)]"
                };

                string[] colname = new string[]
                {
                 // "[PVC進口價(美元/公噸)]",
                 // "[中鋼冷軋鋼品盤價(元/公噸)]",
                  //"[中鋼熱軋鋼品盤價(元/公噸)]",
                  "[倫敦鋁現貨價(美元/公噸)]",
                  "[倫敦鎳三個月期貨價(美元/公噸)]",
                  "[3000PSI混凝土台北價格(元/m3)]",
                  "[紐約銅期貨價(美分/磅)]",
                 
                  "[中鋼鋼板盤價(元/公噸)]"
                };

                var source0 = new ChromeDriver();
             source0.Navigate().GoToUrl("https://fubon-ebrokerdj.fbs.com.tw/z/ze/zeq/zeq.djhtm");
             System.Threading.Thread.Sleep(SleepTime);
             string sourcePage0 = source0.PageSource;
             source0.Quit();
             Materal2(sourcePage0, materialarray[i], tablename[i], colname[i]);
         }
        }
        private void Materal2(string Page, string Material, string tablename, string colname)
        {
            string pageSource = Page;
            string material = Material;
            string _tablename = tablename;
            string _colname = colname;
            int ikeyword = pageSource.IndexOf(material);
            if(ikeyword == -1)
            {
                string abc = "no need to update";
            }
            else 
            {
            string skeyword = "";
            string link = "";
            int start = 0;
            int end = 0;
            skeyword = pageSource.Substring(ikeyword - 40, 40);
            start = skeyword.IndexOf("href");
            end = skeyword.LastIndexOf("\">");
            link = skeyword.Substring(start + 6, end - (start + 6));
            Materal3(link, material, _tablename, _colname);
            }
        }
        private void Materal3(string Url, string Material, string tablename, string colname)
        {
            //https://fubon-ebrokerdj.fbs.com.tw/z/ze/zeq/zeq.djhtm
            string url = Url;
            string material = Material;
            string _tablename = tablename;
            string _colname = colname;
            string completedUrl = "https://fubon-ebrokerdj.fbs.com.tw" + url;
            var web = new ChromeDriver();
            web.Navigate().GoToUrl(completedUrl);
            System.Threading.Thread.Sleep(SleepTime);
            string pageSource = web.PageSource;
            web.Quit();
            Material4(url, pageSource, material, _tablename, _colname);
        }
        private void Material4(string url, string page, string material, string tablename, string colname)
        {
            string pageSource = page;
            string _material = material;
            string skeyword1 = "";
            string skeyword2 = "";
            string completedkeyword1 = "";
            string _tablename = tablename;
            string _colname = colname;
            int tip0 = 0;
            int tip1 = 0;
            int tipend = 0;
            int notehead = 0;
            int notehead2 = 0;
            int noteheadend = 0;
            string test1 ="" ;
            tip0 = pageSource.IndexOf("tip0");
            if (tip0 <0)
            {
              Materal3(url, _material, _tablename, _colname);
            }
            tip1 = pageSource.IndexOf("title", tip0);
            tipend = pageSource.IndexOf("class", tip1);
            skeyword1 = pageSource.Substring(tip1, tipend - tip1);
            completedkeyword1 = skeyword1.Substring(skeyword1.IndexOf("=") + 2, skeyword1.LastIndexOf("\"") - (skeyword1.IndexOf("=") + 2));
           
            if(completedkeyword1.Contains("美元") || completedkeyword1.Contains("台幣"))
            {
                completedkeyword1 = completedkeyword1.Remove(completedkeyword1.Length - 2, 2);
            }
            //2,275.00 混泥土
            //completedkeyword1 = test1;
            test1 = completedkeyword1;

            if(test1.Contains(","))
            {
                test1= test1.Replace(",", "");
            }
            // skeyword2 = skeyword2.Remove(skeyword2.Length - 1, 1);
            //int splitZeo = completedkeyword1.IndexOf(".");
            //string completedkeyword2 = completedkeyword1.Substring(0, splitZeo);

            notehead = pageSource.IndexOf("notehead");
            notehead2 = pageSource.IndexOf("\">", notehead);
            noteheadend = pageSource.IndexOf("<", notehead2);
            skeyword2 = pageSource.Substring(notehead2 + 2, noteheadend - (notehead2 + 2)); 

            //skeyword2 日期 ex: 2022/10/22
            int basepoint = 24311;        //steel 鋼板
            double basepoint2 = 263.17;   //copper 銅
            double basepoint3 = 1875;     //concrete 混凝土
            double basepoint4 = 12070;    //nickel 倫敦鎳
            double basepoint5 = 23067;    //cold-rolled 冷軋
            double basepoint6 = 20841;    //hot-rolled 熱軋鋼品中鋼盤價
            double basepoint7 = 1761;     //aluminium 倫敦鋁現貨價
            double basepoint8 = 876;      //PVC 
            double basecompare = 0; 

            string splitYear = skeyword2;
            string splitYear2 = splitYear.Substring(0, 4);
            int isplitYear2 = Int32.Parse(splitYear2);
            //splitYear2 = 年 ex: 2022
            string splitMonth = skeyword2;
            string splitMonth2 = skeyword2.Substring(5, 2);
            //splitMonth2 = 月 ex: 10
            string splitDay = skeyword2;
            string splitDay2 = skeyword2.Substring(8, 2);
            //splitDay2 = 日 ex: 22
            DateTime today = DateTime.Today;
            string rec = today.Date.ToString();
            //rec = 日期+時間 ex: 10/28/2022 12:00:00 AM
            //string test3 = completedkeyword2;
            //test3 = 取得值
            string test4 = string.Empty;
            //for (int t1 = 0; t1 < test3.Length; t1++)
            //{
            //    if (char.IsDigit(test3[t1]))
            //    {
            //        test4 += test3[t1];
            //    }
            //}
            //test4 = 每一字元是否是是數字
            //int tmpskeyword2 = Int32.Parse(test4); //get exact number of iron


            double tmpskeyword2 =double.Parse(test1, CultureInfo.InvariantCulture.NumberFormat);
            bool deter = false;
            double c;
            string a;
            if (_material.Contains("鋼板"))
            {
                basecompare = tmpskeyword2 - basepoint;
                basecompare = basecompare / basepoint;
            }
            if (_material.Contains("銅"))
            {
                basecompare = tmpskeyword2 - basepoint2;
                basecompare = basecompare / basepoint2;
            }
            if (_material.Contains("混凝土"))
            {
                basecompare = tmpskeyword2 - basepoint3;
                basecompare = basecompare / basepoint3;
            }
            if (_material.Contains("倫敦鎳"))
            {
                basecompare = tmpskeyword2 - basepoint4;
                basecompare = basecompare / basepoint4;
            }
            if (_material.Contains("冷軋"))
            {
                basecompare = tmpskeyword2 - basepoint5;
                basecompare = basecompare / basepoint5;
            }
            if (_material.Contains("熱軋"))
            {
                basecompare = tmpskeyword2 - basepoint6;
                basecompare = basecompare / basepoint6;
            }
            if (_material.Contains("倫敦鋁現貨價"))
            {
                basecompare = tmpskeyword2 - basepoint7;
                basecompare = basecompare / basepoint7;
            }
            if (_material.Contains("PVC 進口價"))
            {
                basecompare = tmpskeyword2 - basepoint8;
                basecompare = basecompare / basepoint8;
            }
            a = basecompare.ToString("0.000");
            c = double.Parse(a) * 100;

            deter = RetreivedPrices(
                                    isplitYear2,                //1
                                    splitMonth2,                //2
                                    splitDay2,                  //3
                                    skeyword2,                  //4
                                    test1,                      //5
                                    0,                         //6
                                    "",                      //7
                                    c,                          //8
                                    0,                          //9
                                    0,                          //10
                                    0,                          //11
                                    rec.Substring(0, 10),       //12
                                    "每週",                      //13
                                    tablename,                  //14
                                    colname);                   //15
            if (deter == true)
            {
            }
            else
            {
             InsertPrices(
                          isplitYear2,                //1
                          splitMonth2,                //2
                          splitDay2,                  //3
                          skeyword2,                  //4
                          test1,                      //5
                          0,                          //6
                          "",                         //7
                          c,                          //8
                          "",                          //9
                          0,                          //10
                          0,                          //11
                          rec.Substring(0, 10),       //12
                          "每週",                      //13
                          tablename,                  //14
                          colname);                   //15
            }

        }
        private void Signals(string Page, string Url)
        {
            string pageSource = Page;
            string url = Url;
            string skeyword = "";
            string skeyword2 = ""; //月的值
            string skeyword3 = ""; //年的值
            int start = 0;
            int end = 0;
            int nextpoint = 0;
            int nextpoint2 = 0;
            int i = 0;
            bool deter = false;
            while (nextpoint != -1 && nextpoint2 != -1)
            {
                start = pageSource.IndexOf("<tspan>", nextpoint);
                end = pageSource.IndexOf("</tspan>", nextpoint);
                nextpoint = pageSource.IndexOf("<tspan>", nextpoint + 1);

                if (nextpoint == -1 || nextpoint2 == -1)
                {
                }
                else
                {
                    skeyword = pageSource.Substring(nextpoint + 7, end - (start + 7));
                    i++;
                }
            }

            start = pageSource.IndexOf("month.t.m");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword2 = pageSource.Substring(nextpoint+2, nextpoint2- (nextpoint+2));

            start = pageSource.IndexOf("month.t.y");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword3 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));
 
            deter = GetPoints("景氣對策信號", skeyword3 + "/" + skeyword2, skeyword, "");
            if (deter == true)
            {
            }
            else
            {
                PutPoints("景氣對策信號", skeyword3 + "/" + skeyword2, skeyword, "");
            }
        }

        /// <summary>
        /// 領先指標
        /// </summary>
        /// <param name="Page"></param>
        /// <param name="Url"></param>
        /// <returns></returns>
        private  string LeadPoints(string Page, string Url)
        {
            string error = "";
            string pageSource = Page;
            string url = Url;
            string skeyword = "";
            string skeyword2 = "";
            string skeyword3 = "";
            string skeyword4 = "";
            bool deter = false;
            int start = 0;
            int end = 0;
            int start2 = 0;
            int end2 = 0;
            int nextpoint = 0;
            int nextpoint2 = 0;
            int i = 0;
            while (nextpoint != -1 && nextpoint2 != -1)
            {
                start = pageSource.IndexOf("number: 2\">", nextpoint);
                end = pageSource.IndexOf("</p>", start);
                nextpoint = pageSource.IndexOf("number: 2\">", nextpoint + 1);

                start2 = pageSource.IndexOf("percentage\">", nextpoint2);
                end2 = pageSource.IndexOf("</p>", start2);
                nextpoint2 = pageSource.IndexOf("percentage\">", nextpoint2 + 1);

                if (nextpoint == -1 || nextpoint2 == -1)
                {
                }
                else
                {
                 skeyword = pageSource.Substring(nextpoint + 11, end - (start + 11));
                 skeyword2 = pageSource.Substring(nextpoint2 + 12, end2 - (start2 + 12));
                 if (skeyword == "0.00")
                 {
                     error = "no";
                 }
                 else
                 {
                  i++;
                 }
                }
            }
 
            start = pageSource.IndexOf("month.t.m");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword3 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));
            start = pageSource.IndexOf("month.t.y");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword4 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));
            if (skeyword.Contains("0.0"))
            {
                Points2();
            }
            else 
            {
            deter = GetPoints("領先指標", skeyword4 + "/" + skeyword3, skeyword, "較上月變化" + skeyword2);
            if (deter == true)
            {

            }
            else
            {
                PutPoints("領先指標", skeyword4 + "/" + skeyword3, skeyword, "較上月變化" + skeyword2);
            }
            }
            return error;
        }

        /// <summary>
        /// 同時指標
        /// </summary>
        /// <param name="Page"></param>
        /// <param name="Url"></param>
        /// <returns></returns>
        private  string Coincidence(string Page, string Url)
        {
            bool deter = false;
            string error = "";
            string pageSource = Page;
            string url = Url;
            string skeyword = "";
            string skeyword2 = "";
            string skeyword3 = "";
            string skeyword4 = "";
            int start = 0;
            int end = 0;
            int start2 = 0;
            int end2 = 0;
            int nextpoint = 0;
            int nextpoint2 = 0;
            int i = 0;

            while (nextpoint != -1 && nextpoint2 != -1)
            {
                start = pageSource.IndexOf("number: 2\">", nextpoint);
                end = pageSource.IndexOf("</p>", start);
                nextpoint = pageSource.IndexOf("number: 2\">", nextpoint + 1);

                start2 = pageSource.IndexOf("percentage\">", nextpoint2);
                end2 = pageSource.IndexOf("</p>", start2);
                nextpoint2 = pageSource.IndexOf("percentage\">", nextpoint2 + 1);

                if (nextpoint == -1 || nextpoint2 == -1)
                {
                }
                else
                {
                    skeyword = pageSource.Substring(nextpoint + 11, end - (start + 11));
                    skeyword2 = pageSource.Substring(nextpoint2 + 12, end2 - (start2 + 12));

                    if (skeyword == "0.00")
                    {
                        error = "no";
                    }
                    else
                    {
                        i++;
                    }
                }
            }

            start = pageSource.IndexOf("month.t.m");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword3 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));
            start = pageSource.IndexOf("month.t.y");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword4 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));

            if (skeyword.Contains("0.0"))
            {
                Points3();
            }
            else
            { 
            deter = GetPoints("同時指標", skeyword4 + "/" + skeyword3, skeyword, "較上月變化" + skeyword2);
            if (deter == true)
            {
            }
            else
            {
                PutPoints("同時指標", skeyword4 + "/" + skeyword3, skeyword, "較上月變化" + skeyword2);
            }
            }
            return error;
        }

        /// <summary>
        /// 落後指標
        /// </summary>
        /// <param name="Page"></param>
        /// <param name="Url"></param>
        /// <returns></returns>
        private  string Lagged(string Page, string Url)
        {
            bool deter = false;
            string error = "";
            string pageSource = Page;
            string url = Url;
            string skeyword = "";
            string skeyword2 = "";
            string skeyword3 = "";
            string skeyword4 = "";
            int start = 0;
            int end = 0;
            int start2 = 0;
            int end2 = 0;
            int nextpoint = 0;
            int nextpoint2 = 0;
            int i = 0;
            while (nextpoint != -1 && nextpoint2 != -1)
            {
                start = pageSource.IndexOf("number: 2\">", nextpoint);
                end = pageSource.IndexOf("</p>", start);
                nextpoint = pageSource.IndexOf("number: 2\">", nextpoint + 1);
                start2 = pageSource.IndexOf("percentage\">", nextpoint2);
                end2 = pageSource.IndexOf("</p>", start2);
                nextpoint2 = pageSource.IndexOf("percentage\">", nextpoint2 + 1);
                if (nextpoint == -1 || nextpoint2 == -1)
                {
                }
                else
                {
                 skeyword = pageSource.Substring(nextpoint + 11, end - (start + 11));
                 skeyword2 = pageSource.Substring(nextpoint2 + 12, end2 - (start2 + 12));
                 if (skeyword == "0.00")
                 {
                  error = "no";
                 }
                 else
                 {
                  i++;
                 }
                }
            }
            start = pageSource.IndexOf("month.t.m");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword3 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));
            start = pageSource.IndexOf("month.t.y");
            nextpoint = pageSource.IndexOf("\">", start);
            nextpoint2 = pageSource.IndexOf("<", nextpoint);
            skeyword4 = pageSource.Substring(nextpoint + 2, nextpoint2 - (nextpoint + 2));
            if (skeyword.Contains("0.0"))
            {
                Points4();
            }
            else { 
            deter = GetPoints("落後指標", skeyword4 + "/" + skeyword3, skeyword, "較上月變化" + skeyword2);
            if (deter == true)
            {
            }
            else
            {
                PutPoints("落後指標", skeyword4 + "/" + skeyword3, skeyword, "較上月變化" + skeyword2);
            }
            }
            return error;
        }

        /// <summary>
        /// 租金指數
        /// </summary>
        /// <param name="Page"></param>
        /// <param name="Url"></param>
        /// <returns></returns>
        private string Rent(string Page, string Url)
        {
         //顯示主計總處租金指數(%)趨勢圖">主計總處租金指數(%) 
         string error = "";
         string pageSource = Page;
         string url = Url;
         string skeyword = "";  //租金指數
         string skeyword2 = ""; //日期
         int start = 0;
         int sNext = 0;
         int sNNext = 0;
         bool deter = false;
         start = pageSource.IndexOf("顯示主計總處租金指數(%)趨勢圖\">主計總處租金指數(%)");
         sNext = pageSource.IndexOf("<span>", start);
         sNNext = pageSource.IndexOf("</span>", sNext);
         skeyword = pageSource.Substring(sNext+6, sNNext - (sNext+6));
         sNNext = pageSource.IndexOf("</span></td></tr>", sNext);
         skeyword2 = pageSource.Substring(sNNext-6, 6);

            //111/10
            string s0 = skeyword2.Substring(0, 3);
            string s1 = skeyword2.Substring(1, 2);
            int inumber = 0;
            inumber = Int32.Parse(s1) + 11;
            s1 = inumber.ToString();
            skeyword2 = skeyword2.Replace('/', 'Y');
            skeyword2 = skeyword2.Replace(s0, s1);
 
            if (skeyword == "")
            {
                Points5();
            }
            if (skeyword.Length <= 3)
            {
                error = "no";
            }
            else
            {
                deter = GetPoints("租金指數(%)", skeyword2, skeyword, "");
                if (deter == true)
                {
                }
                else
                {
                 PutPoints("租金指數(%)", skeyword2, skeyword, "");
                }
            }

            return error;
        }
        private string PMI(string Page, string Url)
        {
          string error = "";
          string pageSource = Page;
          string url = Url;
          string skeyword = "";  //變化百分比
          string skeyword2 = ""; //PMI指數
          string skeyword3 = ""; //月
          string skeyword4 = ""; //年

          int start = 0;
          int start2 = 0;
          int start3 = 0;
          int end = 0;
          int sNext = 0;
          int sNNext = 0;
          bool deter = false;
          start = pageSource.IndexOf("percentage\"");
          end = pageSource.IndexOf("</span>", start);
          skeyword = pageSource.Substring(start, end - start);
          start2 = skeyword.IndexOf("\">");
          skeyword = skeyword.Substring(start2 + 2, skeyword.Length - (start2 + 2));
          start = pageSource.IndexOf("\"big.score+'");
          end = pageSource.IndexOf("</p>", start);

          skeyword2 = pageSource.Substring(start, end - start);
          start2 = skeyword2.IndexOf("\">");

          skeyword2 = skeyword2.Substring(start2 + 2, skeyword2.Length - (start2 + 2));
          skeyword2 = skeyword2.Remove(skeyword2.Length - 1, 1);

          start3 = pageSource.IndexOf("monthName");
          sNext = pageSource.IndexOf("\">", start3);
          sNNext = pageSource.IndexOf("<", sNext);
          skeyword3 = pageSource.Substring(sNext+2, sNNext - (sNext+2));

            start3 = pageSource.IndexOf("month.t.y");
            sNext = pageSource.IndexOf("</div>", start3);
            skeyword4 = pageSource.Substring(start3+11, sNext - (start3+11));

          if (skeyword2 == "")
          {
           Points6();
          }
          if (skeyword2.Length <= 3)
          {
           error = "no";
          }
          else
          {
           deter = GetPoints("製造業採購經理指數(PMI)", skeyword4+"/"+ skeyword3, skeyword2, "較上月變化" + skeyword + "百分點");
           if (deter == true)
           {
           }
           else
           {
            PutPoints("製造業採購經理指數(PMI)", skeyword4 + "/" + skeyword3, skeyword2, "較上月變化" + skeyword + "百分點");
           }
          }
 
            return error;
        }
        private string NMI(string Page, string Url)
        {
            bool deter = false;
            string error = "";
            string pageSource = Page;
            string url = Url;
            string skeyword = "";
            string skeyword2 = ""; //NMI值
            string skeyword3 = "";  //月
            string skeyword4 = ""; //年
            int start = 0;
            int end = 0;
            int start2 = 0;
            int sNext = 0;
            int sNNext = 0;
            start = pageSource.IndexOf("percentage\"");
            end = pageSource.IndexOf("</span>", start);
            skeyword = pageSource.Substring(start, end - start);
            start2 = skeyword.IndexOf("\">");
            skeyword = skeyword.Substring(start2 + 2, skeyword.Length - (start2 + 2));
            start = pageSource.IndexOf("\"big.score+'");
            end = pageSource.IndexOf("</p>", start);
            skeyword2 = pageSource.Substring(start, end - start);
            start2 = skeyword2.IndexOf("\">");
            skeyword2 = skeyword2.Substring(start2 + 2, skeyword2.Length - (start2 + 2));
            skeyword2 = skeyword2.Remove(skeyword2.Length - 1, 1);

            start2 = pageSource.IndexOf("monthName");
            sNext = pageSource.IndexOf("\">", start2);
            sNNext = pageSource.IndexOf("<", sNext);
            skeyword3 = pageSource.Substring(sNext + 2, sNNext - (sNext + 2));

            start2 = pageSource.IndexOf("month.t.y");
            sNext = pageSource.IndexOf("</div>", start2);
            skeyword4 = pageSource.Substring(start2 + 11, sNext - (start2 + 11));
 
            if (skeyword2 == "")
            {
                Points7();
            }
            else
            {
              deter = GetPoints("非製造業經理人指數(NMI)", skeyword4 + "/" + skeyword3, skeyword2, "較上月變化" + skeyword + "百分點");
                if (deter == true)
                {
                }
                else
                {
                    PutPoints("非製造業經理人指數(NMI)", skeyword4 + "/" + skeyword3, skeyword2, "較上月變化" + skeyword + "百分點");
                }
            }
            return error;
        }
 
        private static bool GetPoints(string title, string date, string rate, string e)
        {
            bool deter = false;
            SqlConnection myConn = new SqlConnection(connString);
            string str4 = "select count (*) from [DYSD_總體指標_2] where date =@date and rate =@rate ";
            SqlCommand myCommand = new SqlCommand(str4, myConn);
            try
            {
                myConn.Open();
                myCommand.Parameters.AddWithValue("@title", title);
                myCommand.Parameters.AddWithValue("@date", date);
                myCommand.Parameters.AddWithValue("@rate", rate);
                myCommand.Parameters.AddWithValue("@e", e);
                int DataExist = (int)myCommand.ExecuteScalar();

                if (DataExist > 0)
                {
                    deter = true;
                }
                else
                {
                    deter = false;
                }
            }
            catch (System.Exception ex)
            {
                string abc = (ex.ToString());
            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
            }
            return deter;

        }
        private static bool PutPoints(string title, string date, string rate, string e)
        {
            SqlConnection myConn = new SqlConnection(connString);
            string str4 = "INSERT INTO [DYSD_總體指標_2] (Title, Date, Rate, Memo) VALUES(@title, @date, @rate, @e)";
            SqlCommand myCommand = new SqlCommand(str4, myConn);
            try
            {
                myConn.Open(); 
                myCommand.Parameters.AddWithValue("@title", title);
                myCommand.Parameters.AddWithValue("@date", date);
                myCommand.Parameters.AddWithValue("@rate", rate);
                myCommand.Parameters.AddWithValue("@e", e);
                myCommand.ExecuteNonQuery();
            }
            catch (System.Exception ex)
            {
                string abc = (ex.ToString());
            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
            }
            return true;
        }

        private bool RetreivedPrices(
                                    int year,           //1
                                    string month,       //2
                                    string day,         //3
                                    string date,        //4
                                    string rate1,       //5
                                    int rate2,          //6
                                    string memo,        //7
                                    double baseprice,   //8 
                                    int rate3,       //9
                                    int rate4,          //10
                                    int rate5,          //11
                                    string update,      //12
                                    string upcircle,    //13
                                    string name,        //14
                                    string name2)       //15 
        {
            SqlConnection myConn = new SqlConnection(connString);
            // string str = " select count (*) from  [DYDC].[dbo].[listPrices] where name =@name and rate1 = @rate1";
            string str = " select count (*) 年 from  [DYDC].[dbo]." + name + " where 日期 =@date and " + name2 + " = @rate1";
            bool deter = false;
            SqlCommand myCommand = new SqlCommand(str, myConn);
            try
            {
                myConn.Open();
                myCommand.Parameters.AddWithValue("@year", year);
                myCommand.Parameters.AddWithValue("@month", month);
                myCommand.Parameters.AddWithValue("@day", day);
                myCommand.Parameters.AddWithValue("@date", date);
                myCommand.Parameters.AddWithValue("@rate1", rate1);
                myCommand.Parameters.AddWithValue("@rate2", rate2);
                myCommand.Parameters.AddWithValue("@memo", memo);
                myCommand.Parameters.AddWithValue("@baseprice", baseprice);
                myCommand.Parameters.AddWithValue("@rate3", rate3);
                myCommand.Parameters.AddWithValue("@rate4", rate4);
                myCommand.Parameters.AddWithValue("@rate5", rate5);
                myCommand.Parameters.AddWithValue("@update", update);
                myCommand.Parameters.AddWithValue("@upcircle", upcircle);
                //myCommand.Parameters.AddWithValue("@name", name);
                // myCommand.Parameters.AddWithValue("@name2", name2);
                int DataExist = (int)myCommand.ExecuteScalar();
                if (DataExist > 0)
                {
                    deter = true;
                }
                else
                {
                    deter = false;
                }
            }
            catch (System.Exception ex)
            {
                string err = (ex.ToString());
            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
            }
            return deter;
        }

        private void InsertPrices(
                                  int    year,        //1
                                  string month,       //2
                                  string day,         //3
                                  string date,        //4
                                  string rate1,       //5
                                  int rate2,          //6
                                  string memo,        //7
                                  double baseprice,   //8 
                                  string rate3,       //9
                                  int rate4,          //10
                                  int rate5,          //11
                                  string update,      //12
                                  string upcircle,    //13
                                  string name,        //14
                                  string name2)       //15 
        {
            SqlConnection myConn = new SqlConnection(connString);
            string str = "INSERT INTO [dbo]." + name +
            "([年]" +
            ",[月]" +
            ",[日]" +
            ",[日期]" +
            "," + name2 +
            ",[tmp]" +
            ",[重要記事]" +
            ",[2019.09為基準漲跌幅]" +
            ",[區間均價]" +
            ",[與均價差異]" +
            ",[差異百分比]" +
            ",[更新日期]" +
            ",[更新週期])" +
            "VALUES(" +
            "@year," +
            "@month," +
            "@day," +
            "@date," +
            "@rate1," +
            "@rate2," +
            "@memo," +
            "@baseprice," +
            "@rate3," +
            "@rate4," +
            "@rate5," +
            "@update," +
            "@upcircle)";
            SqlCommand myCommand = new SqlCommand(str, myConn);
            try
            {
                myConn.Open();
                myCommand.Parameters.AddWithValue("@year", year);
                myCommand.Parameters.AddWithValue("@month", month);
                myCommand.Parameters.AddWithValue("@day", day);
                myCommand.Parameters.AddWithValue("@date", date);
                myCommand.Parameters.AddWithValue("@rate1", rate1);
                myCommand.Parameters.AddWithValue("@rate2", rate2);
                myCommand.Parameters.AddWithValue("@memo", memo);
                myCommand.Parameters.AddWithValue("@baseprice", baseprice);
                myCommand.Parameters.AddWithValue("@rate3", rate3);
                myCommand.Parameters.AddWithValue("@rate4", rate4);
                myCommand.Parameters.AddWithValue("@rate5", rate5);
                myCommand.Parameters.AddWithValue("@update", update);
                myCommand.Parameters.AddWithValue("@upcircle", upcircle);
               // myCommand.Parameters.AddWithValue("@name", name);
               // myCommand.Parameters.AddWithValue("@name2", name2);
                myCommand.ExecuteNonQuery();
            }
            catch(System.Exception ex)
            {
                string err = (ex.ToString());
            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label2.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            label6.Visible = false;
            label7.Visible = false;
            label8.Visible = false;
            label9.Visible = false;
            label9.Visible = false;
            label10.Visible = false;
        }

        private void label1_Click(object sender, EventArgs e)
        {
         System.Diagnostics.Process.Start("mailto:chenfenghuang@dydc.com.tw");
        }
    }
}
