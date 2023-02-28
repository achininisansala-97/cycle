using IronXL;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace cycle
{
    class Include
    {
        public IWebDriver driver;



        public void TakingHTML2CanvasFullPageScreenshot()
        {
            
      
            try
            {
                System.Threading.Thread.Sleep(4000);
                Screenshot TakeScreenshot = ((ITakesScreenshot)driver).GetScreenshot();
                string imagePath = "./../../../selenium-screenshot.png";
                TakeScreenshot.SaveAsFile(imagePath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
            }
            
        }



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
            Thread.Sleep(2000);
        }

        public void Xpaths(string path)
        {
            IWebElement Product = driver.FindElement(By.XPath(path));
            Product.Click();

        }


        public void Xpaths2(string path2,string key, WorkSheet worksheet)
        {
            IWebElement Product2 = driver.FindElement(By.XPath(path2));
            Product2.SendKeys(key);
        }



        public void workSheetCall(string url, string worksheet, string range)
            {
                WorkBook workbook = WorkBook.Load(url);
                WorkSheet rangeInWorksheet = workbook.GetWorkSheet(worksheet);
                _ = rangeInWorksheet[range];
            }
        


    }


}
