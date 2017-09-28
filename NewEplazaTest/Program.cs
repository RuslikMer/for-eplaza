using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using System.Drawing.Imaging;
using System.IO;
using OpenQA.Selenium;
using System.Drawing;
using System.Net;
using System.Net.Mail;
using OpenQA.Selenium.Support.UI;


namespace NewEplazaTest
{
    class Program
    {
        public class Order
        {
            public int OrderId { get; set; }
            public string UrlAdress { get; set; }
            public string ProdName { get; set; }
            public double ProductValue { get; set; }
            public double DeliveryValue { get; set; }
            public double OrderValue { get; set; }
            public double AdminOrderValue { get; set; }
            public double Sap { get; set; }
            public double Uniteller { get; set; }
            public double nProductValue { get; set; }
            public double nDeliveryValue { get; set; }
            public double nOrderValue { get; set; }
            public double nAdminOrderValue { get; set; }
            public double nSap { get; set; }
            public double nUniteller { get; set; }
            public RemoteWebDriver driver { set; get; }
            Object WrapText { get; set; }


            static void Main(string[] args)
            {

                using (var driver = new ChromeDriver())
                {
                    //void WaitShowElement(By XPath)
                    //{
                    //    WebDriverWait iWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    //    iWait.Until(ExpectedConditions.ElementIsVisible(XPath));
                    //}

                    var timeout = new TimeSpan(00, 00, 07);
                    string[] UrlAdress = new string[] { "https://eplaza.panasonic.ru/products/digital_av/av_accessories/head_phone/RP-TCM105E/", "https://eplaza.panasonic.ru/products/digital_av/av_accessories/head_phone/RP-HT161E-K/" };

                    var productsPrise = new List<int>();

                    var count = -1;
                    for(int t = 0; t< UrlAdress.Length; t++)
                    {
                        count++;
                        driver.Navigate().GoToUrl(UrlAdress[t]);
                        var ProductPrise = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div/div[2]/div[1]/div[2]/div[2]/div[1]/span[1]").Text;
                        int PrPrise = Convert.ToInt32(ProductPrise);
                        productsPrise.Add(PrPrise);
                        var buy = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div/div/div/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/div/span/span/a[1]")));
                        buy.Click();
                        driver.Manage().Window.Maximize();
                    }

                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/personal/order/make/");

                    //рандом емайл
                    string s = "";
                    Random rand = new Random();
                    for (int v = 0; v < 6; v++)
                    {
                        s += Convert.ToChar(rand.Next(65, 90));
                    }

                    var city = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[2]/div[2]/div[2]/div/a/span")));
                    city.Click();
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
                    city = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[3]/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/div[1]/ul/li[1]/a")));
                    city.Click();
                    Task.Delay(1500).Wait();
                    var adress = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[4]/div[2]/input");
                    adress.SendKeys("Тестовая");
                    var index = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[5]/div[2]/input");
                    index.SendKeys("123456");
                    var mail = driver.FindElementByName("NEW_EMAIL");
                    mail.SendKeys(s + "@mail.ru");
                    var password = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[3]/div[2]/input");
                    password.SendKeys("Cent73");
                    password = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[4]/div[2]/input");
                    password.SendKeys("Cent73");
                    var name = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[5]/div[2]/input");
                    name.SendKeys("Тест");
                    var lastname = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[6]/div[2]/input");
                    lastname.SendKeys("Тестов");

                    //генератор номера
                    string namb = "";
                    Random randj = new Random();
                    for (int v = 0; v < 7; v++)
                    {
                        namb += Convert.ToChar(rand.Next(48, 57));
                    }
                    var number = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[7]/div[2]/input");
                    number.SendKeys("993" + namb);
                    Directory.CreateDirectory("Public\\Test");
                    Directory.CreateDirectory("\\Users\\Stefan-PC\\Documents\\Test");
                    Directory.CreateDirectory("C:\\Users\\Stefan-PC\\Documents\\Test");
                    var pol = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[3]/div[8]/div[2]/div/div/select/option[2]");
                    pol.Click();

                    //  скриншот + генератор названия
                    //string screenname = "";
                    //Random randscreenname = new Random();
                    //for (int v = 0; v < 4; v++)
                    //{
                    //    screenname += Convert.ToChar(rand.Next(65, 90));
                    //}

                    //var screenShot = driver.GetScreenshot();
                    //screenShot.SaveAsFile("C:\\Users\\Stefan-PC\\Documents\\Test\\" + screenname + ".jpg", ScreenshotImageFormat.Png);
                    //driver.ToString();

                    //начальная стоимость  из оформления
                    var PriseOrder = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[5]/table/tbody/tr[1]/td[2]/span[1]").Text;
                    string[] prOrd = PriseOrder.Split(new Char[] { ' ' });
                    foreach (string m in prOrd)
                    {
                        if (m.Trim() != "") ;
                    }
                    int PrOr = Convert.ToInt32(prOrd[0] + prOrd[1]);

                    //итоговая стоимость из оформления
                    var PriseOrderTotal = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[5]/table/tbody/tr[3]/td[2]/span[1]").Text;
                    string[] prise = PriseOrderTotal.Split(new Char[] { ' ' });
                    foreach (string m in prise)
                    {
                        if (m.Trim() != "");
                    }
                    int PrOrT = Convert.ToInt32(prise[0]+prise[1]);
                    pol.Submit();

                    //подтверждение совершеннолетия
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
                    var PopAp = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"inline_is_adult\"]/div/div[2]/div/div/div/label")));
                    PopAp.Click();
                    Task.Delay(1000).Wait();

                    //переход в лк
                    driver.Navigate().Forward();
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/personal/orders/");

                    //номер заказа
                    var OrderNumber = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[4]/div/div/div/div[2]/div[1]/div[1]/a").Text;
                    int OrdNum = Convert.ToInt32(OrderNumber);

                    //стоимость доставки из лк
                    var delivery = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[4]/div/div/div/div[2]/div[2]/div[4]/div[2]/div[2]/div/span[1]").Text;
                    string[] split = delivery.Split(new Char[] { ' ', '+', 'Р' });
                    foreach (string m in split)
                    {
                        if (m.Trim() != "");
                    }
                    int Del = Convert.ToInt32(split[1]);

                    //авторизация админки
                    driver.Navigate().Forward();
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/bitrix/admin/sale_order.php?lang=ru");
                    mail = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.Name("USER_LOGIN")));
                    mail.Clear();
                    mail.SendKeys("Bot");
                    password = driver.FindElementByName("USER_PASSWORD");
                    password.SendKeys("123456");
                    password.Submit();

                    //поиск заказа
                    var search = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.Name("filter_user_login")));
                    search.SendKeys(s + "@mail.ru");
                    search = driver.FindElementById("tbl_sale_order_filterset_filter");
                    search.Click();

                    //детальная страница заказа
                    search = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"tbl_sale_order\"]/tbody/tr/td[3]/table/tbody/tr/td[2]/b/a")));
                    search.Click();
                    var AdminOrder = driver.FindElementByXPath("//*[@id=\"edit1_edit_table\"]/tbody/tr[62]/td/table/tbody/tr/td[2]/div/table/tbody/tr[5]/td[2]/div").Text;
                    string[] Adm = AdminOrder.Split(new Char[] { ' ' });
                    foreach (string m in Adm)
                    {
                        if (m.Trim() != "");
                    }
                    int AdmOrd = Convert.ToInt32(Adm[0] + Adm[1]);
                    
                    //отмена заказа
                    var action = driver.FindElementByXPath("//*[@id=\"btn_show_cancel\"]/td[2]/a/span");
                    action.Click();
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
                    action = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[6]/table/tbody/tr[2]/td[2]/div[3]/span[1]/span[2]")));
                    action.Click();

                    //url заказа
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().First());
                    var label = driver.SwitchTo().Window(driver.WindowHandles.ToList().First()).Url; 

                    //XML
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/manager/services/showAllPricesXmlFormat.php");
                    var bodyxml = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/form/input[1]")));
                    bodyxml.SendKeys(OrderNumber);
                    bodyxml.Submit();

                    //юнителлер
                    string XmlUniteller = driver.FindElementByXPath("/html/body/order[1]/item[2]/total_summ").Text;
                    int Unit = Convert.ToInt32(XmlUniteller);

                    //названия товаров
                    var ProductsNames = driver.FindElementsByTagName("name");
                    var ProductName = new List<string>();
                    for(int o = 0; o< UrlAdress.Length; o++)
                    {
                        var PrName = ProductsNames[o].Text;
                        ProductName.Add(PrName);
                    }
                    
                    //сап
                    string XmlSap = driver.FindElementByXPath("/html/body/order[4]/item[1]/total_summ").Text;
                    int Sap = Convert.ToInt32(XmlSap);
                                        
                    int a = OrdNum;
                    string a1 = String.Join(" / ", ProductName);
                    int b = PrOr;
                    int c = Del;
                    int d = PrOrT;
                    int d1 = AdmOrd;
                    int e = Sap;
                    int f = Unit;
                    int g = productsPrise.Sum();
                    int h = 240;
                    int i = productsPrise.Sum() + h;
                    int i1 = i;
                    int j = i;
                    int k = i;
                    string k1 = label;

                    var EplazaOrders = new List<Order>
                    {
                        new Order
                        {
                            OrderId = a,
                            ProdName = a1,
                            ProductValue = b,
                            DeliveryValue = c,
                            OrderValue = d,
                            AdminOrderValue = d1,
                            Sap = e,
                            Uniteller = f,

                            nProductValue = g,
                            nDeliveryValue = h,
                            nOrderValue = i,
                            nAdminOrderValue = i1,
                            nSap = j,
                            nUniteller = k,
                            UrlAdress = k1
                        }
                    };
                        DisplayInExcel(EplazaOrders);
                }
                       
                 void DisplayInExcel(IEnumerable<Order> orders)
                 {
                    var excelApp = new Excel.Application();
                    Object missing = Type.Missing;
                    excelApp.Visible = true;
                    excelApp.Workbooks.Add(missing);
                    Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                    int[] Rep = new int[] { 4, 6, 8, 10, 12, 14, 2};
                    int[] nRep = new int[] { 3, 5, 7, 9, 11, 13, 15, 1 };

                    //обЪединение ячеек
                    for (int i = 0; i < 6; i++)
                    {
                        Excel.Range oRange1;
                        oRange1 = workSheet.Range[workSheet.Cells[1, Rep[i]], workSheet.Cells[1, nRep[i]]];
                        oRange1.Merge(Type.Missing);
                    }

                    string[] Arr = new string[] { "D", "F", "H", "J", "L","N" };
                    string[] Arr2 = new string[] { "C", "E", "G", "I", "K", "M", "O" ,"A", "B" };
                    string[] Arr3 = new string[] { "Cтоимость товара(ов)", "Стоимость доставки", "Стоимость заказа","Заказ в админке", "САП", "Uniteller","Ссылка на заказ" ,"Номер заказа", "Наименование товара" };
                    string[] Arr4 = new string[] { "должно быть", "факт" };
                    for (int i = 0; i < 9; i++)
                    {
                        workSheet.Cells[1, Arr2[i]] = Arr3[i];
                    }
                    for(int i = 0; i < 6; i++)
                    {
                        workSheet.Cells[2, Arr2[i]] = Arr4[0];
                    };
                    for (int i = 0; i < 6; i++)
                    {
                        workSheet.Cells[2, Arr[i]] = Arr4[1];
                    };
                    var row = 2;
                    foreach (var ord in orders)
                    {
                        row++;
                        workSheet.Cells[row, "A"] = ord.OrderId;
                        workSheet.Cells[row, "B"] = ord.ProdName;
                        workSheet.Cells[row, "D"] = ord.ProductValue;
                        workSheet.Cells[row, "F"] = ord.DeliveryValue;
                        workSheet.Cells[row, "H"] = ord.OrderValue;
                        workSheet.Cells[row, "J"] = ord.AdminOrderValue;
                        workSheet.Cells[row, "L"] = ord.Sap;
                        workSheet.Cells[row, "N"] = ord.Uniteller;
                        workSheet.Cells[row, "C"] = ord.nProductValue;
                        workSheet.Cells[row, "E"] = ord.nDeliveryValue;
                        workSheet.Cells[row, "G"] = ord.nOrderValue;
                        workSheet.Cells[row, "I"] = ord.nAdminOrderValue;
                        workSheet.Cells[row, "K"] = ord.nSap;
                        workSheet.Cells[row, "M"] = ord.nUniteller;
                        workSheet.Cells[row, "O"] = ord.UrlAdress;
                        workSheet.Cells[row, "P"] = " ";
                    }

                    //цвет текста          
                    for (int i = 0; i < 6; i++)
                    {
                        if (workSheet.Cells[3, Arr2[i]].FormulaLocal == workSheet.Cells[3, Arr[i]].FormulaLocal)
                        {
                            Excel.Range rng2 = workSheet.get_Range(Arr[i] + "3");
                            rng2.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Green);
                        }
                        else
                        {
                            Excel.Range rng2 = workSheet.get_Range(Arr[i] + "3");
                            rng2.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    }

                    //редактирование ячеек 
                    workSheet.Range["A1", "O2"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    for (int i = 0; i < 7; i++)
                    {
                        (workSheet.Cells[3, Rep[i]] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        (workSheet.Cells[3, nRep[i]] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    }
                    
                    //границы ячеек 
                    for (int i = 1; i < 4; i++)
                    {
                        Excel.Range rt = workSheet.get_Range("A" + i, "O" + i);
                        rt.Borders.ColorIndex = 0;
                        rt.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        rt.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }

                    workSheet.Cells[3, "B"].WrapText = true;

                    //сохранение отчета
                    excelApp.DisplayAlerts = false;
                    workSheet.SaveAs(string.Format(@"{0}\Test.xlsx", Environment.CurrentDirectory));
                    excelApp.Quit();

                    //отправка отчета
                    MailAddress fromMailAddress = new MailAddress("boteplaza@gmail.com", "Test");
                    MailAddress toAddress = new MailAddress("sag@m-st.ru", "Uncle Bob");
                    using (MailMessage mailMessage = new MailMessage(fromMailAddress, toAddress))
                    using (SmtpClient smtpClient = new SmtpClient())
                    {
                        mailMessage.Subject = "Отчет по автотестированию Еплазы";
                        mailMessage.Body = "Откройте документ";
                        //прикрепляем вложение
                        Attachment attData = new Attachment("C:/Users/new/Documents/Visual Studio 2017/Projects/NewEplazaTest/NewEplazaTest/bin/Debug/Test.xlsx");
                        mailMessage.Attachments.Add(attData);

                        smtpClient.Host = "smtp.gmail.com";
                        smtpClient.Port = 587;
                        smtpClient.EnableSsl = true;
                        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtpClient.UseDefaultCredentials = false;
                        smtpClient.Credentials = new NetworkCredential(fromMailAddress.Address, "123456eplaza");
                        smtpClient.Send(mailMessage);
                    }
                }
            }
        }
    }
}