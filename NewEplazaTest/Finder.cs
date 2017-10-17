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
    public class Finder
    {
        public ChromeDriver driver { set; get; }
        TimeSpan timeout = new TimeSpan(00, 00, 05);

        public Finder(ChromeDriver driver)
        {
            this.driver = driver;
        }

        public IWebElement Xpath(string x)
        {
            return driver.FindElementByName(x);
        }

        public IWebElement Name(string x)
        {
            return driver.FindElementByName(x);
        }

        public IWebElement Id(string x)
        {
            return driver.FindElementById(x);
        }

        public WebDriverWait Time()
        {
            return new WebDriverWait(this.driver, timeout);
        }
    }
}
