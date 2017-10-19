using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Chrome;
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
    
    public class Program
    {
        public class Order
        {
            public string OrderId { get; set; }
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
            public RemoteWebDriver Driver { set; get; }
            Object WrapText { get; set; }


            static void Main(string[] args)
            {
                using (var driver = new ChromeDriver())
                {
                    TimeSpan timeout = new TimeSpan(00, 00, 05);
                    var find = new Finder(driver);

                    string[][] Urls =
                    {
                        new string[] { "https://eplaza.panasonic.ru/products/digital_av/av_accessories/head_phone/RP-TCM105E/", "https://eplaza.panasonic.ru/products/digital_av/av_accessories/head_phone/RP-HT161E-K/" },
                        new string[] { /*"https://eplaza.panasonic.ru/products/composite_sets/composite_sets/composite_sets/ES-LT2N-S820%20+%20WES9015Y1361/", "https://eplaza.panasonic.ru/products/composite_sets/composite_sets/composite_sets/ES-LV6N-S820%20+%20WES9034Y1361/"*/ }
                    };
                    var productsPrise = new List<int>();

                    var count = -1;
                    foreach (var q in Urls)
                    {
                        if (q == Urls[1])
                        {
                            for (int t = 0; t < Urls[1].Length; t++)
                            {
                                count++;
                                driver.Navigate().GoToUrl(Urls[1][t]);
                                var ProductPrise = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[4]/div/div[1]/div[1]/span/span[1]").Text;
                                string[] prPri = ProductPrise.Split(new Char[] { ' ', 'Р' });
                                foreach (string m in prPri)
                                {
                                    if (m.Trim() != "");
                                }
                                int PrPrise = Convert.ToInt32(prPri[4] + prPri[5]);
                                productsPrise.Add(PrPrise);
                                var buy = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div/div/div/div[4]/div/div[1]/div[1]/div[2]/div[1]/a")));
                                buy.Click();
                                driver.Manage().Window.Maximize();
                            }
                        }
                        else
                        {
                            for (int t = 0; t < Urls[0].Length; t++)
                            {
                                count++;
                                driver.Navigate().GoToUrl(Urls[0][t]);
                                var ProductPrise = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div/div[2]/div[1]/div[2]/div[2]/div[1]/span[1]").Text;
                                int PrPrise = Convert.ToInt32(ProductPrise);
                                productsPrise.Add(PrPrise);
                                var buy = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div/div/div/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr/td[3]/div/div/span/span/a[1]")));
                                buy.Click();
                                driver.Manage().Window.Maximize();
                            }
                        }
                    }
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/personal/order/make/");

                    //NewUser user = new NewUser(driver);
                    //user.NewUsers();
                    User user = new User(driver);
                    user.Auth();
                    user.Action();

                    //начальная стоимость  из оформления
                    var PriseOrder = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[5]/table/tbody/tr[1]/td[2]/span[1]").Text;
                    string[] prOrd = PriseOrder.Split(new Char[] { ' ' });
                    foreach (string m in prOrd)
                    {
                        if (m.Trim() != "");
                    }
                    int PrOr = Convert.ToInt32(prOrd[0] + prOrd[1]);

                    //стоимость доставки из оформления
                    var delivery = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[5]/table/tbody/tr[2]/td[2]/span[1]").Text;
                    int Del = Convert.ToInt32(delivery);

                    //итоговая стоимость из оформления
                    var PriseOrderTotal = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[5]/table/tbody/tr[3]/td[2]/span[1]").Text;
                    string[] prise = PriseOrderTotal.Split(new Char[] { ' ' });
                    foreach (string m in prise)
                    {
                        if (m.Trim() != "");
                    }
                    int PrOrT = Convert.ToInt32(prise[0] + prise[1]);
                    user.Submit();

                    //подтверждение совершеннолетия
                    //_18_ check = new _18_(driver);
                    //check.Check();

                    //переход в лк
                    driver.Navigate().Forward();
                    driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/personal/orders/");

                    //номер заказа
                    string OrdNum = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[4]/div/div/div/div[2]/div[1]/div[1]/a").Text;
                    
                    //все действия в админке
                    Admin adm = new Admin(driver);
                    adm.Revocation();

                    //XML
                    XML xml = new XML(driver, OrdNum);
                    xml.Info();

                    int h = 260; //cтоимость доставки
                    int i = productsPrise.Sum() + h;

                    var EplazaOrders = new List<Order>
                    {
                        new Order
                        {
                            OrderId = OrdNum,
                            ProdName = String.Join(" / ", xml.ProductName),
                            ProductValue = PrOr,
                            DeliveryValue = Del,
                            OrderValue = PrOrT,
                            AdminOrderValue = adm.AdmOrd,
                            Sap = xml.Sap,
                            Uniteller = xml.Unit,

                            nProductValue = productsPrise.Sum(),
                            nDeliveryValue = h,
                            nOrderValue = i,
                            nAdminOrderValue = i,
                            nSap = i,
                            nUniteller = i,
                            UrlAdress = adm.label
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

                    int[] Rep = new int[] { 4, 6, 8, 10, 12, 14, 2 };
                    int[] nRep = new int[] { 3, 5, 7, 9, 11, 13, 15, 1 };

                    //обЪединение ячеек
                    for (int i = 0; i < 6; i++)
                    {
                        Excel.Range oRange1;
                        oRange1 = workSheet.Range[workSheet.Cells[1, Rep[i]], workSheet.Cells[1, nRep[i]]];
                        oRange1.Merge(Type.Missing);
                    }

                    string[] Arr = new string[] { "D", "F", "H", "J", "L", "N" };
                    string[] Arr2 = new string[] { "C", "E", "G", "I", "K", "M", "O", "A", "B" };
                    string[] Arr3 = new string[] { "Cтоимость товара(ов)", "Стоимость доставки", "Стоимость заказа", "Заказ в админке", "САП", "Uniteller", "Ссылка на заказ", "Номер заказа", "Наименование товара" };
                    string[] Arr4 = new string[] { "должно быть", " факт " };
                    for (int i = 0; i < 9; i++)
                    {
                        workSheet.Cells[1, Arr2[i]] = Arr3[i];
                    }
                    for (int i = 0; i < 6; i++)
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
                            rng2.Font.Color = ColorTranslator.ToOle(Color.Green);
                        }
                        else
                        {
                            Excel.Range rng2 = workSheet.get_Range(Arr[i] + "3");
                            rng2.Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }

                    //редактирование ячеек 
                    workSheet.Range["A1", "O2"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    workSheet.Range["A3", "O3"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //границы ячеек 
                    for (int i = 1; i < 4; i++)
                    {
                        Excel.Range rt = workSheet.get_Range("A" + i, "O" + i);
                        rt.Borders.ColorIndex = 0;
                        rt.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        rt.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }

                    //перенос текста
                    workSheet.Cells[3, "B"].WrapText = true;

                    //сохранение отчета
                    excelApp.DisplayAlerts = false;
                    workSheet.SaveAs(string.Format(@"{0}\Test.xlsx", Environment.CurrentDirectory));
                    excelApp.Quit();

                    //отправка отчета
                    SendMail mail = new SendMail();
                    mail.Mail();
                }

            }
        }
    }
}