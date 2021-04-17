using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
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

namespace Upsocity6WpfNetcore
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        notify notify2 = new notify();
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = notify2;
            MainUpSocity6New();
        }
        ChromeDriver driver;
        ChromeOptions options = new ChromeOptions();
        ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
        List<Image> imageList;
        string username;
        string pass;
        int soLuongAnh1lan;
        int soLuonganh1Ngay;
        int soLanLoi;
        int demlanlogin = 0;
        string pathChromeExe;
        //.\\chromium\\win32-564778\\chrome-win32\\chrome.exe
        string ScripclickAll = "var items = document.querySelectorAll(\".undefined\");\n"
                    + "for (var i = 0; i < items.length; i++) {\n"
                    + "    \n"
                    + "        items[i].click();\n"
                    + "  \n"
                    + "}";

        int j = 0;
        Thread thread;

 
        public class Image
        {
            public string Foldername { get; set; }
            public string Imagename { get; set; }
            public string Title { get; set; }
            public string Des { get; set; }
            public string Tag { get; set; }
            public string Main { get; set; }
        }
        List<Image> readimage(string nameFile)
        {
            List<Image> imageList = new List<Image>();
            try
            {
               
                var package = new ExcelPackage(new FileInfo(nameFile));

                // lấy ra sheet đầu tiên để thao tác
                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file. lưu ý file excel bắt đầu từ số 1 không phải số 0
                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    try
                    {
                        // biến j biểu thị cho một column trong file
                        int j = 1;
                        Image image = new Image();
                        // lấy ra cột họ tên tương ứng giá trị tại vị trí [i, 1]. i lần đầu là 2
                        // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh
                        if (workSheet.Cells[i, j].Value == null)
                        {
                            break;
                        }
                        string Foldername = workSheet.Cells[i, j++].Value.ToString();
                        image.Foldername = Foldername;
                        string Imagename = workSheet.Cells[i, j++].Value.ToString();
                        image.Imagename = Imagename;
                        string Title = workSheet.Cells[i, j++].Value.ToString();
                        image.Title = Title;
                        string Des = workSheet.Cells[i, j++].Value.ToString();
                        image.Des = Des;

                        string Tag = workSheet.Cells[i, j++].Value.ToString();
                        image.Tag = Tag;
                        string main = workSheet.Cells[i, j++].Value.ToString();
                        image.Main = main;


                        imageList.Add(image);




                    }
                    catch
                    {

                    }
                }
            }
            catch
            {
                //System.Environment.Exit(0);
                //MessageBox.Show("Error read excel!");
                throw;
            }

            return imageList;
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Boolean stop = false;
            username = userName.Text;
            pass = passWord.Text;
            soLuongAnh1lan = Int32.Parse(soAnhUpMotLan.Text);
            soLuonganh1Ngay = Int32.Parse(soAnhUp.Text);
            soLanLoi = Int32.Parse(soLanLoiTxt.Text);
            pathChromeExe= PathChorme.Text;

            thread = new Thread(() =>
            {

                try
                {
                    imageList = readimage("listing.xlsx");
                }
                catch (Exception)
                {
                    stop = true;
                    notify2.DataValue = "Error read excel!";
                    //MessageBox.Show("Error read excel!");
                }

                if (stop == false)
                {
                    try
                    {
                        string FullPath = pathChromeExe;
                        options.BinaryLocation = FullPath;
                        options.AddArguments("user-data-dir=ChromeProfile");
                        options.AddArguments("--disable-notifications");
                        options.AddArguments("disable-extensions");
                        options.AddArguments("--no-sandbox");
                        options.AddArguments("start-maximized");
                        options.AddExcludedArgument("enable-automation");
                        options.AddAdditionalCapability("useAutomationExtension", false);
                        options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3452.0 Safari/537.36");
                        driverService.HideCommandPromptWindow = true;
                        driver = new ChromeDriver(driverService, options);
                        if (driver != null)
                        {
                            task(imageList, driver, username, pass, soLuongAnh1lan, soLuonganh1Ngay, soLanLoi);
                        }

                    }
                    catch (Exception)
                    {
                        notify2.DataValue = "Lỗi khởi tạo chrome";
                        //MessageBox.Show("Lỗi khởi tạo chrome");
                    }
                }





            });
            thread.IsBackground = true;
            thread.Start();
        }

        void task(List<Image> imageList, ChromeDriver driver, string username, string passs, int soLuongAnh1lan, int soLuonganh1Ngay, int soLanLoi)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            try
            {
                waitHandle.WaitOne();
                notify2.DataValue = "Bắt đầu + j";
                string exeFile = (new System.Uri(Assembly.GetEntryAssembly().CodeBase)).AbsolutePath;
                DirectoryInfo di = new DirectoryInfo(exeFile);
                Console.WriteLine(di.Parent.FullName);
                string fathparen = di.Parent.FullName;
                //Boolean oke = false;
                driver.Url = "https://society6.com/artist-studio";


                Thread.Sleep(randomNumber(5000, 2000));


                if (isElementCss("div.unauthorized_manage_1A5XP", driver))
                {
                    notify2.DataValue = "Login";
                    login(driver, username, passs);
                    notify2.DataValue = "Logined";
                    Thread.Sleep(randomNumber(5000, 3500));
                    driver.Url = "https://society6.com/artist-studio";
                    Thread.Sleep(randomNumber(5000, 2000));


                }
                driver.FindElement(By.CssSelector("button.uploadBtn_banner_2b9z8")).Click();

                Thread.Sleep(3 * 1000);

                for (int i = 0; i < imageList.Count(); i++)
                {
                    try
                    {
                        if (File.Exists(fathparen.Replace("%20", " ") + "\\" + imageList[j].Foldername + "\\" + imageList[j].Imagename))
                        {
                            wirte();
                        }
                        else
                        {

                            Console.WriteLine("lỗi file " + imageList[j].Imagename);
                            j++;
                            continue;
                        }
                    }
                    catch
                    {
                        j++;
                        continue;
                    }
                    notify2.DataValue = "Submit Ảnh  "+j;
                    waitHandle.WaitOne();
                    IWebElement elemtitle = driver.FindElement(By.CssSelector("input.titleInput_fileHandler_2QfDm"));
                    elemtitle.SendKeys(imageList[j].Title);
                    IWebElement elem = driver.FindElement(By.XPath("//input[@type='file']"));
                    elem.SendKeys(fathparen.Replace("%20", " ") + "\\" + imageList[j].Foldername + "\\" + imageList[j].Imagename);

                    Thread.Sleep(30000);
                    
                    if (isElementCss("button.continueBtn_fileHandler_3P74b", driver))
                    {
                        driver.FindElement(By.CssSelector("button.continueBtn_fileHandler_3P74b")).Click();

                    }
                    notify2.DataValue = "Click continew Ảnh  " + j;
                    Thread.Sleep(randomNumber(5000, 2000));
                    Actions action = new Actions(driver);
                    driver.FindElement(By.CssSelector("div.rightSideCheckbox_artistAgreement_1yKq1 input")).Click();
                    driver.FindElement(By.XPath("//input[@qa-id='matureContentFalse']")).Click();
                    Thread.Sleep(randomNumber(5000, 2000));
                    if (driver.FindElement(By.CssSelector("button.continueBtn_artistAgreement_1zGHn")).Enabled)
                    {
                        driver.FindElement(By.CssSelector("button.continueBtn_artistAgreement_1zGHn")).Click();

                        Thread.Sleep(randomNumber(15000, 10000));
                    }
                    Thread.Sleep(randomNumber(5000, 2000));
                    List<string> newtag = new List<string>();
                    string[] spearator = { "," };
                    string[] words = imageList[j].Tag.ToLower().Replace(" ", "").Split(spearator, StringSplitOptions.RemoveEmptyEntries);

                    for (int jkk = 0; jkk < words.Length; jkk++)
                    {


                        if (words[jkk] == null || String.IsNullOrEmpty(words[jkk].Trim()))
                        {
                            continue;
                        }

                        if (newtag.Contains(words[jkk].Trim()) == false)
                        {
                            newtag.Add(words[jkk].Trim());
                        }

                    }
                    int lengtag = 20;
                    if (newtag.Count() < 10)
                    {
                        lengtag = newtag.Count();
                    }
                    notify2.DataValue = "Nhập tag  " + j;
                    IWebElement tag = driver.FindElement(By.CssSelector("input#search-creatives"));


                    for (int k = 0; k < lengtag; k++)
                    {

                        foreach (char c in newtag[k].Trim())
                        {
                            tag.SendKeys(c.ToString());

                            Thread.Sleep(150);
                        }
                        tag.SendKeys(Keys.Enter);
                    }
                    Thread.Sleep(randomNumber(5000, 2000));
                    notify2.DataValue = "nhập des  " + j;
                    IWebElement elem55 = driver.FindElement(
                            By.XPath("//*[@id=\"creativesView\"]/div/div[2]/div[1]/div/div[5]/div/div/div[2]/div"));

                    action.MoveToElement(elem55);
                    action.Click().MoveByOffset(0, 181).Click().Build().Perform();
                    Thread.Sleep(randomNumber(5000, 2000));
                    IWebElement des = driver.FindElement(By.CssSelector("textarea.descriptionText_artworkDetails_2y6Zb"));

                    foreach (char c in imageList[j].Des)
                    {

                        des.SendKeys(c.ToString());
                        Thread.Sleep(150);
                    }



                    Thread.Sleep(randomNumber(5000, 2000));
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript(ScripclickAll);
                    Thread.Sleep(randomNumber(5000, 4000));
                    js.ExecuteScript(ScripclickAll);
                    //List<IWebElement> elemen = driver.FindElements(By.CssSelector("div.undefined")).ToList();
                    //foreach (IWebElement webElement in elemen)
                    //{
                    //    try
                    //    {
                    //        webElement.Click();
                    //        Thread.Sleep(1000);
                    //    }
                    //    catch
                    //    {
                    //        continue;
                    //    }
                    //}
                    notify2.DataValue = "Click publish  " + j;
                    Thread.Sleep(randomNumber(5000, 4000));
                    driver.FindElement(By.XPath("//*[@id=\"creativesView\"]/div/div[2]/div[3]/div/div[2]/div/input"))
                        .Click();
                    driver.FindElement(By.CssSelector("button.button_publishStatus_2vaJI")).Click();
                    Thread.Sleep(randomNumber(6000, 5000));
                    if (isElementCss("button.button_publishStatus_2vaJI", driver)
                            && driver.FindElement(By.CssSelector("button.button_publishStatus_2vaJI")).Enabled)
                    {
                        driver.FindElement(By.CssSelector("button.button_publishStatus_2vaJI")).Click();

                    }
                    Thread.Sleep(randomNumber(7000, 6000));
                    notify2.DataValue = "Wait publish  " + j;
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("span.statusPublished_artworkDetails_13LpC")));
                    Thread.Sleep(randomNumber(5000, 2000));
                    driver.Navigate().GoToUrl("https://society6.com/artist-studio");
                    
                    Thread.Sleep(randomNumber(7000, 6000));
                    notify2.DataValue = "Click upload tiep  " + j;
                    driver.FindElement(By.CssSelector("button.uploadBtn_banner_2b9z8")).Click();
                    Thread.Sleep(randomNumber(5000, 2000));
                    if (j != 0 && j % soLuongAnh1lan == 0)
                    {
                        notify2.DataValue = "Đủ 50 ảnh nghỉ  " + j;
                        Thread.Sleep(1000000);
                    }
                    if (j >= soLuonganh1Ngay)
                    {
                        notify2.DataValue = "đủ số lượng ảnh ngày nghỉ  " + j;
                        Thread.Sleep(86400);

                    }
                    j++;

                }


            }
            catch (Exception)
            {
                driver.Navigate().Refresh();
                Thread.Sleep(2000);
                j++;
                notify2.DataValue = "Lỗi thực hiện lại task  " + j;
                task(imageList, driver, username, pass, soLuongAnh1lan, soLuonganh1Ngay, soLanLoi);
            }



        }

        void login(ChromeDriver driver, String username, String passs)
        {
            try
            {

                //

                demlanlogin++;
                driver.Url = "https://society6.com/login?done=/";
                Thread.Sleep(5000);

                int timeoutto = 5 * 1000;
                int timeout = 0;
                Boolean oke = false;

                //WebDriverWait wait = new WebDriverWait(driver, 10);
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                while (!oke && timeout < timeoutto)
                {
                    try
                    {

                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("input#email")));
                        IWebElement usernameele = driver.FindElement(By.CssSelector("input#email"));
                        usernameele.SendKeys(Keys.Control + "a");
                        usernameele.SendKeys(Keys.Delete);


                        foreach (char c in username)
                        {

                            usernameele.SendKeys(c.ToString());
                            Thread.Sleep(150);
                        }
                        // usernameele.SendKeys(username);

                        IWebElement passele = driver.FindElement(By.CssSelector("input#password"));
                        passele.SendKeys(Keys.Control + "a");
                        passele.SendKeys(Keys.Delete);
                        // passele.SendKeys(passs);


                        foreach (char c in passs)
                        {

                            passele.SendKeys(c.ToString());
                            Thread.Sleep(150);
                        }

                        oke = true;

                    }
                    catch
                    {
                        Thread.Sleep(100);
                        timeout = timeout + 100;
                    }
                }

                oke = false;

                while (!oke && timeout < timeoutto)
                {
                    try
                    {
                        driver.FindElement(By.CssSelector("button#submitButton")).Click();
                        oke = true;
                    }
                    catch
                    {
                        Thread.Sleep(100);
                        timeout = timeout + 100;
                    }
                }
                //

            }
            catch
            {
                // TODO: handle exception
            }

        }
        public async void wirte()
        {
            using (StreamWriter sw = File.AppendText("log.txt"))
            {
                sw.WriteLine(imageList[j].Imagename);

            }
        }


        private readonly Random _random = new Random();
        public int randomNumber(int max, int min)
        {
            return _random.Next(min, max);
        }
        protected static Boolean isElementXpath(String tagcss, ChromeDriver driver)
        {
            try
            {
                driver.FindElement(By.XPath(tagcss));
                return true;
            }
            catch
            {
                return false;
            }
        }

        protected static Boolean isElementCss(String tagcss, ChromeDriver driver)
        {
            try
            {
                driver.FindElement(By.CssSelector(tagcss));
                return true;
            }
            catch
            {
                return false;
            }
        }


        Boolean isElement(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            pathChromeExe = PathChorme.Text;
            thread = new Thread(() =>
            {

                try
                {
                    notify2.DataValue = "www.howkteam.com";

                    string FullPath = pathChromeExe;
                    options.BinaryLocation = FullPath;
                    options.AddArguments("user-data-dir=ChromeProfile");
                    options.AddArguments("--disable-notifications");
                    options.AddArguments("start-maximized");
                    options.AddExcludedArgument("enable-automation");
                    options.AddAdditionalCapability("useAutomationExtension", false);
                    options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3452.0 Safari/537.36");
                    driverService.HideCommandPromptWindow = true;
                    driver = new ChromeDriver(driverService, options);
                    notify2.DataValue = "www.howkteam.com 22";
                    driver.Url = "https://www.google.com/";



                }
                catch (Exception)
                {
                    MessageBox.Show("Lỗi khởi tạo chrome");
                    throw;

                }




            });
            thread.IsBackground = true;
            thread.Start();
        }
        public class notify : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            protected virtual void OnPropertyChanged(string name)
            {
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
            string dataValue;
            public string DataValue
            {
                get { return dataValue; }
                set { dataValue = value; OnPropertyChanged("DataValue"); }
            }
        }
        public void MainUpSocity6New()
        {


            string[] lines = File.ReadAllLines("info.txt");
            string username = lines[0];
            string pass = lines[1];
            string profile = lines[2];
            string userdata = lines[3];
            string chromeexe = lines[4];
            PathChorme.Text = chromeexe;
            userName.Text=username;
            passWord.Text = pass;
        }
        private static EventWaitHandle waitHandle = new ManualResetEvent(initialState: true);

        // Main thread
        private void OnPauseClick(object sender, RoutedEventArgs e)
        {
            waitHandle.Reset();
        }

        private void OnResumeClick(object sender, RoutedEventArgs e)
        {
            waitHandle.Set();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Thread thread = new Thread(() =>
            {

                try
                {
                    driver.Quit();
                    driver.Close();


                }
                catch (Exception)
                {
                 

                }




            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Thread thread = new Thread(() =>
            {
                while (true)
                {
                    waitHandle.WaitOne();
                    Console.WriteLine("Hey");
                    Thread.Sleep(2000);
                }

            });
            thread.IsBackground = true;
            thread.Start();

        }
    }
}
