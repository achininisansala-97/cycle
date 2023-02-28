using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cycle
{
    class Check_login
    {

        public IWebDriver driver;

        public void login()
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


    }
}
