using System;
using System.Linq;
using System.Threading.Tasks;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;


namespace NewEplazaTest
{
    public class NewUser
    {
        public ChromeDriver driver { set; get; }
        public IWebElement sex { set; get; }
        public string s { set; get; }

        public NewUser(ChromeDriver driver)
        {
            this.driver = driver;
        }

        public void NewUsers()
        {
            TimeSpan timeout1 = new TimeSpan(00, 00, 05);
            var find = new Finder(driver);

            //генератор почты
            Random rand = new Random();
            for (int v = 0; v < 6; v++)
            {
                s += Convert.ToChar(rand.Next(65, 90));
            }

            //заполнение данных
            var city = find.Time().Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[2]/div[2]/div/a/span")));
            //var city = find.Xpath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[2]/div[2]/div/a/span");
            city.Click();
            driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
            city = (new WebDriverWait(driver, timeout1)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[3]/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/div[1]/ul/li[1]/a")));
            city.Click();
            Task.Delay(1500).Wait();
            var adress = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[4]/div[2]/input");
            adress.SendKeys("Тестовая");
            var index = driver.FindElementByXPath("//*[@id=\"process_order\"]/div[1]/div/div[2]/div[5]/div[2]/input");
            index.SendKeys("123456");
            var mail = driver.FindElementByName("NEW_EMAIL");
            mail.SendKeys(s + "@bik.ru");
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
            sex = driver.FindElementByXPath("/html/body/div[1]/div/div/div/div[3]/div/div[2]/form/div[1]/div/div[3]/div[8]/div[2]/div/div/select/option[2]");
            sex.Click();
        }

        public void Submit()
        {
            sex.Submit();
        }
    }
}