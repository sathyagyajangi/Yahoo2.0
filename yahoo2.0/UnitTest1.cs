using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace yahoo2._0
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            IWebDriver driver = new ChromeDriver();

            driver.Url = "https://beta-cricket-yahoo.sportz.io/";

            driver.Manage().Window.Maximize();

            ReadOnlyCollection<IWebElement> elements = driver.FindElements(By.CssSelector(".section-wrap"));

            foreach(IWebElement url in elements)
            {
                if(url.Displayed)
                {

                    Console.WriteLine(url.GetAttribute("href"));



                }

            }

            Console.WriteLine();

        }
    }
}
