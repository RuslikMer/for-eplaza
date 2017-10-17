using System;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;

namespace NewEplazaTest
{
    public class XML
    {
        public List<string> ProductName { set; get; }
        public ChromeDriver driver { set; get; }
        public int Sap { set; get; }
        public int Unit { set; get; }
        public XML(ChromeDriver driver, string OrdNum)
        {
            ProductName = new List<string>();
            this.driver = driver;
            this.ordNum = OrdNum;
        }
        
        private string ordNum;



        public void Info()
        {
            NewUser user = new NewUser(driver);
            TimeSpan timeout3 = new TimeSpan(00, 00, 05);

            this.driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/manager/services/showAllPricesXmlFormat.php");
            var bodyxml = (new WebDriverWait(driver, timeout3)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/form/input[1]")));
            bodyxml.SendKeys(this.ordNum);
            bodyxml.Submit();

            //юнителлер
            string XmlUniteller = driver.FindElementByXPath("/html/body/order[1]/item[2]/total_summ").Text;
            Unit = Convert.ToInt32(XmlUniteller);

            //названия товаров
            var ProductsNames = driver.FindElementsByTagName("name").ToList();
            for (int o = 0; o < ProductsNames.Count / 4; o++)
            {
                var PrName = ProductsNames[o].Text;
                ProductName.Add(PrName);
            }

            //сап
            string XmlSap = driver.FindElementByXPath("/html/body/order[4]/item[1]/total_summ").Text;
            Sap = Convert.ToInt32(XmlSap);
        }
    }
}