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


namespace NewEplazaTest
{
    class Program
    {
        public class Order
        {
            public static int PriseOrder;
            public static int PriseOrder2;
            public static int delivery;
            public static int PriseOrderTotals;
            public int OrderId { get; set; }
            public double ProductValue { get; set; }
            public double DeliveryValue { get; set; }
            public double OrderValue { get; set; }
            public double Sap { get; set; }
            public double Uniteller { get; set; }
            public double nProductValue { get; set; }
            public double nDeliveryValue { get; set; }
            public double nOrderValue { get; set; }
            public double nSap { get; set; }
            public double nUniteller { get; set; }
            public RemoteWebDriver driver { set; get; }
            public static string a { set; get; }

            static void Main(string[] args)
            {
                using (var driver = new ChromeDriver())
                {
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/products/digital_av/av_accessories/head_phone/RP-TCM105E/");
                    driver.Manage().Window.Maximize();
                    Task.Delay(9000).Wait();
                    var to = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/div/span/span/a[1]");
                    to.Click();
                    Task.Delay(2000).Wait();
                    to = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/div[5]/a");
                    to.Click();

                    //рандом емайл
                    string s = "";
                    Random rand = new Random();
                    for (int v = 0; v < 6; v++)
                    {
                        s += Convert.ToChar(rand.Next(65, 90));
                    }

                    var city = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[2]/div[2]/div[2]/div/a/span");
                    city.Click();
                    Task.Delay(2000).Wait();
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
                    var city2 = driver.FindElementByXPath("/html/body/div[3]/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/div[1]/ul/li[1]/a");
                    city2.Click();
                    Task.Delay(2000).Wait();
                    var adress = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[4]/div[2]/input");
                    adress.SendKeys("Тестовая");
                    var index = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[5]/div[2]/input");
                    index.SendKeys("123456");
                    to = driver.FindElementByName("NEW_EMAIL");
                    to.SendKeys(s + "@mail.ru");
                    var password = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[3]/div[2]/input");
                    password.SendKeys("Cent73");
                    var password2 = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[3]/div[4]/div[2]/input");
                    password2.SendKeys("Cent73");
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
                    Task.Delay(2000);


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
                    int PrOr = Convert.ToInt32(PriseOrder);
                    

                    //итоговая стоимость из оформления
                    var PriseOrderTotals = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[5]/table/tbody/tr[3]/td[2]/span[1]").Text;
                    string[] prise = PriseOrderTotals.Split(new Char[] { ' ' });
                    foreach (string m in prise)
                    {
                        if (m.Trim() != "");
                            //Console.WriteLine(m);
                    }
                    int PrOrT = Convert.ToInt32(prise[0]+prise[1]);
                    Console.WriteLine(PrOrT);
                    pol.Submit();

                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/personal/orders/");
                    Task.Delay(4000);
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
                    var PopAp = driver.FindElementByXPath("//*[@id=\"inline_is_adult\"]/div/div[2]/div/div/div/label");
                    PopAp.Click();
                    Task.Delay(3500);
                    //номер заказа
                    var OrderNumber = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[4]/div/div/div/div[2]/div[1]/div[1]/a").Text;
                    int OrdNum = Convert.ToInt32(OrderNumber);
                    //стоимость доставки из лк
                    Task.Delay(1500);
                    var delivery = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[4]/div/div/div/div[2]/div[2]/div[4]/div[2]/div[2]/div/span[1]").Text;

                    string[] split = delivery.Split(new Char[] { ' ', '+', 'Р' });
                    foreach (string m in split)
                    {
                        if (m.Trim() != "") ;
                            //Console.WriteLine(m);
                    }
                    int Del = Convert.ToInt32(split[1]);



                    //авторизация админки
                    Task.Delay(3000);
                    driver.Navigate().Forward();
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/bitrix/admin/sale_order.php?lang=ru");
                    Task.Delay(1000).Wait();
                    var mail = driver.FindElementByName("USER_LOGIN");
                    mail.Clear();
                    mail.SendKeys("Bot");
                    mail = driver.FindElementByName("USER_PASSWORD");
                    mail.SendKeys("123456");
                    mail.Submit();
                    Task.Delay(9000).Wait();

                    //поиск заказа
                    var nmail = driver.FindElementByName("filter_user_login");
                    nmail.SendKeys(s + "@mail.ru");
                    mail = driver.FindElementById("tbl_sale_order_filterset_filter");
                    mail.Click();
                    Task.Delay(5000).Wait();

                    //детальная страница заказа
                    mail = driver.FindElementByXPath("//*[@id=\"tbl_sale_order\"]/tbody/tr/td[3]/table/tbody/tr/td[2]/b/a");
                    mail.Click();
                    var AdminOrder = driver.FindElementByXPath("//*[@id=\"edit1_edit_table\"]/tbody/tr[62]/td/table/tbody/tr/td[2]/div/table/tbody/tr[5]/td[2]/div").Text;
                    Console.WriteLine(AdminOrder);

                    //отмена заказа
                    mail = driver.FindElementByXPath("//*[@id=\"btn_show_cancel\"]/td[2]/a/span");
                    mail.Click();
                    Task.Delay(1500).Wait();
                    driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
                    mail = driver.FindElementByXPath("/html/body/div[6]/table/tbody/tr[2]/td[2]/div[3]/span[1]/span[2]");
                    mail.Click();

                    //XML
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/manager/services/showAllPricesXmlFormat.php");
                    Task.Delay(10000).Wait();
                    var bodyxml = driver.FindElementByXPath("/html/body/form/input[1]");
                    bodyxml.SendKeys(OrderNumber);
                    bodyxml.Submit();

                    //юнителлер
                    string XmlUniteller = driver.FindElementByXPath("/html/body/order[1]/item[2]/total_summ").Text;
                    Console.WriteLine(XmlUniteller);
                    int Unit = Convert.ToInt32(XmlUniteller);

                    //сап
                    string XmlSap = driver.FindElementByXPath("/html/body/order[4]/item[1]/total_summ").Text;
                    Console.WriteLine(XmlSap);
                    int Sap = Convert.ToInt32(XmlSap);
                    //string words = "This is a list of words, with: a bit of punctuation" +
                    //   "\tand a tab character.";

                    //string[] split = XMLuniteller.Split(new Char[] { ' ', 'и', 'ц', '=', '>','q','s','v','t','(',')','[',']','' });

                    //foreach (string m in split)
                    //{

                    //    if (m.Trim() != "")
                    //        Console.WriteLine(m);
                    //}

                    //Console.WriteLine(XMLuniteller);
                    //Task.Delay(2000).Wait();
                    //var p = driver.FindElementByXPath("/html/body").Text;
                    //File.WriteAllText(@"C:\Users\Stefan-PC\Documents\Test\page.txt", p);




                    //string[] Words = new string[] { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k" };
                    //int[] Value = new int[] { OrderNumber, "", "", "", "", "", "", "", "", "", "" };
                    int a = OrdNum;
                    int b = PrOr;
                    int c = Del;
                    int d = PrOrT;
                    int e = Sap;
                    int f = Unit;
                    int g = 850;
                    int h = 240;
                    int i = 1090;
                    int j = 1090;
                    int k = 1090;

                    var EplazaOrders = new List<Order>
                    {
                        new Order
                        {
                                  OrderId = a,
                                  ProductValue = b,
                                  DeliveryValue = c,
                                  OrderValue = d,
                                  Sap = e,
                                  Uniteller = f,

                                  nProductValue = g,
                                  nDeliveryValue = h,
                                  nOrderValue = i,
                                  nSap = j,
                                  nUniteller = k
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

                    int[] Rep = new int[] { 2, 4, 6, 8, 10 };
                    int[] nRep = new int[] { 3, 5, 7, 9, 11, 1 };

                    for (int i = 0; i < 5; i++)
                    {
                        Excel.Range oRange1;
                        oRange1 = workSheet.Range[workSheet.Cells[1, Rep[i]], workSheet.Cells[1, nRep[i]]];
                        oRange1.Merge(Type.Missing);
                    }

                    string[] Arr = new string[] { "C", "E", "G", "I", "K" };
                    string[] Arr2 = new string[] { "B", "D", "F", "H", "J", "A" };
                    string[] Arr3 = new string[] { "Cтоимость товара", "Стоимость доставки", "Стоимость заказа", "САП", "Uniteller", "Номер заказа" };
                    for (int i = 0; i < 6; i++)
                    {
                        workSheet.Cells[1, Arr2[i]] = Arr3[i];
                    }

                    var row = 1;
                    foreach (var ord in orders)
                    {
                        row++;
                        workSheet.Cells[row, "A"] = ord.OrderId;
                        workSheet.Cells[row, "C"] = ord.ProductValue;
                        workSheet.Cells[row, "E"] = ord.DeliveryValue;
                        workSheet.Cells[row, "G"] = ord.OrderValue;
                        workSheet.Cells[row, "I"] = ord.Sap;
                        workSheet.Cells[row, "K"] = ord.Uniteller;
                        workSheet.Cells[row, "B"] = ord.nProductValue;
                        workSheet.Cells[row, "D"] = ord.nDeliveryValue;
                        workSheet.Cells[row, "F"] = ord.nOrderValue;
                        workSheet.Cells[row, "H"] = ord.nSap;
                        workSheet.Cells[row, "J"] = ord.nUniteller;

                    }

                    // workSheet.Columns[1].AutoFit();
                    // workSheet.Columns[2].AutoFit();

                    //цвет текста          
                    for (int i = 0; i < 5; i++)
                    {
                        if (workSheet.Cells[2, Arr2[i]].FormulaLocal == workSheet.Cells[2, Arr[i]].FormulaLocal)
                        {

                            Excel.Range rng2 = workSheet.get_Range(Arr[i] + "2");
                            rng2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                        }
                        else
                        {
                            Excel.Range rng2 = workSheet.get_Range(Arr[i] + "2");
                            rng2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    }
                    //редактирование ячеек 2 порядка
                    for (int i = 0; i < 6; i++)
                    {
                        (workSheet.Cells[2, nRep[i]] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        //(workSheet.Cells[2, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                    }
                    //редактирование ячеек 1 порядка
                    for (int i = 0; i < 5; i++)
                    {
                        (workSheet.Cells[2, Rep[i]] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    }


                    //границы ячеек
                    workSheet.Range["A1", "K1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    for (int i = 1; i < 3; i++)
                    {
                        Excel.Range rt = workSheet.get_Range("A" + i, "K" + i);
                        rt.Borders.ColorIndex = 0;
                        rt.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        rt.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }

                    //сохранение отчета
                    excelApp.DisplayAlerts = false;
                    workSheet.SaveAs(string.Format(@"{0}\Test.xlsx", Environment.CurrentDirectory));
                    //Console.ReadKey();
                    excelApp.Quit();

                    //отправка отчета
                    MailAddress fromMailAddress = new MailAddress("boteplaza@gmail.com", "Test");
                    MailAddress toAddress = new MailAddress("sag@m-st.ru", "Uncle Bob");
                    using (MailMessage mailMessage = new MailMessage(fromMailAddress, toAddress))
                    using (SmtpClient smtpClient = new SmtpClient())
                    {

                        mailMessage.Subject = "Отчет по тесту";
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
