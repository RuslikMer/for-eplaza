using System;
using System.Linq;
using System.Threading.Tasks;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;

namespace NewEplazaTest
{
    public class Admin
    {
        public int AdmOrd { set; get; }
        public string label { set; get; }
       
        public ChromeDriver driver { set; get; }
        public Admin(ChromeDriver driver)
        {
            this.driver = driver;
        }

        public void Revocation()
        {
            NewUser user = new NewUser(driver);
            TimeSpan timeout2 = new TimeSpan(00, 00, 05);

            //авторизация админки
            driver.Navigate().Forward();
            driver.Navigate().GoToUrl("https://eplaza.panasonic.ru/bitrix/admin/sale_order.php?lang=ru");
            var mail = (new WebDriverWait(driver, timeout2)).Until(ExpectedConditions.ElementIsVisible(By.Name("USER_LOGIN")));
            mail.Clear();
            mail.SendKeys("Bot");
            var password = driver.FindElementByName("USER_PASSWORD");
            password.SendKeys("123456");
            password.Submit();

            //поиск заказа
            var search = (new WebDriverWait(driver, timeout2)).Until(ExpectedConditions.ElementIsVisible(By.Name("filter_user_login")));
            search.SendKeys(user.s + "@mail.ru");
            search = driver.FindElementById("tbl_sale_order_filterset_filter");
            search.Click();

            //детальная страница заказа
            search = (new WebDriverWait(driver, timeout2)).Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"tbl_sale_order\"]/tbody/tr/td[3]/table/tbody/tr/td[2]/b/a")));
            search.Click();
            var AdminOrder = driver.FindElementByXPath("//*[@id=\"edit1_edit_table\"]/tbody/tr[62]/td/table/tbody/tr/td[2]/div/table/tbody/tr[5]/td[2]/div").Text;
            string[] Adm = AdminOrder.Split(new Char[] { ' ' });
            foreach (string m in Adm)
            {
                if (m.Trim() != "") ;
            }
            AdmOrd = Convert.ToInt32(Adm[0] + Adm[1]);

            //отмена заказа
            var action = driver.FindElementByXPath("//*[@id=\"btn_show_cancel\"]/td[2]/a/span");
            action.Click();
            driver.SwitchTo().Window(driver.WindowHandles.ToList().Last());
            action = (new WebDriverWait(driver, timeout2)).Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[6]/table/tbody/tr[2]/td[2]/div[3]/span[1]/span[2]")));
            action.Click();

            //url заказа
            driver.SwitchTo().Window(driver.WindowHandles.ToList().First());
            label = driver.SwitchTo().Window(driver.WindowHandles.ToList().First()).Url;
        }
    }
}