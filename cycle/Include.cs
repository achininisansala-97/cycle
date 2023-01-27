using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cycle
{
    class Include
    {
        public IWebDriver driver;


        public void Cromelink()
        {
            
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://ims4l-uat.i-wonder.co.uk/product-select");
            this.Sleep();
            IWebElement CookiesPopup = driver.FindElement(By.Id("accept-and-continue"));
            CookiesPopup.Click();
        }


        public void Sleep()
        {
            Thread.Sleep(1000);
        }

        public void Xpaths(string path)
        {
            IWebElement Product = driver.FindElement(By.XPath(path));
            Product.Click();
        }


    }


}
