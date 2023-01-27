using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Xml.Linq;


namespace cycle
{


    public class UnitTest1
    {
        public IWebDriver driver;
        Include include = new Include();


        [Fact]
        public void Test1()
        {

            this.cycle();

        }

        public void cycle(){

            include.Cromelink();
            include.Xpaths("/html/body/app-root/div/div[1]/iw-product-select/div/div/div[8]/div/div[2]");
            include.Sleep();



        }
    }
}