using System;
using OpenQA.Selenium.Remote;
using System.IO;
using OpenQA.Selenium;


namespace NewEplazaTest
{
    public class Screenshot
    {
        public RemoteWebDriver driver { set; get; }
        public void Screen()
        {
            //  скриншот + генератор названия
            string screenname = "";
            Random randsn = new Random();
            for (int v = 0; v < 4; v++)
            {
                screenname += Convert.ToChar(randsn.Next(65, 90));
            }

            var screenShot = driver.GetScreenshot();
            Directory.CreateDirectory("C:\\Users\\Test");
            screenShot.SaveAsFile("C:\\Users\\Test\\" + screenname + ".jpg", ScreenshotImageFormat.Png);
            driver.ToString();
        }
    }
}