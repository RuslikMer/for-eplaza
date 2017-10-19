using System;
using System.Linq;
using System.Threading.Tasks;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace NewEplazaTest
{
    public class User
    {
        TimeSpan timeout = new TimeSpan(00, 00, 05);
        public ChromeDriver driver { set; get; }
        public string s { set; get; }
        public IWebElement sub { set; get; }

        public User(ChromeDriver driver)
        {
            this.driver = driver;
        }

        public void Auth()
        {
            var auth = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/header/nav/div/div/ul/li[2]/a")));
            auth.Click();
            driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
            var login = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.Name("USER_LOGIN")));
            login.SendKeys("botadmin@p33.org");
            var pass = driver.FindElementByName("USER_PASSWORD");
            pass.SendKeys("123456");
            auth = driver.FindElementById("js_auth_button");
            auth.Click();
            Task.Delay(1500).Wait();
        }

        public void Action()
        {
            var city = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"process_order\"]/div[1]/div/div[1]/div[2]/div[2]/div/a/span")));
            city.Click();
            driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
            city = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.LinkText("Москва")));
            city.Click();
            Task.Delay(4000).Wait();
            var adress = (new WebDriverWait(driver, timeout)).Until(ExpectedConditions.ElementIsVisible(By.Name("ORDER_PROP_5")));
            adress.SendKeys("Тестовая");
            var index = driver.FindElementByName("ORDER_PROP_4");
            index.SendKeys("123456");
        }
        public void Submit()
        {
            sub.Submit();
        }
    }
}
