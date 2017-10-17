using System;
using System.Linq;
using System.Threading.Tasks;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;

namespace NewEplazaTest
{
    public class _18_
    {
        public ChromeDriver driver { set; get; }
        public _18_(ChromeDriver driver)
        {
            this.driver = driver;
        }

        public void Check()
        {
            var timeout = new TimeSpan(00, 00, 07);

            //подтверждение совершеннолетия
            driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
            var PopAp = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"inline_is_adult\"]/div/div[2]/div/div/div/label")));
            PopAp.Click();
            Task.Delay(1000).Wait();
        }
    }
}
