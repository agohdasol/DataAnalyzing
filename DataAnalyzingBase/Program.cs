using System;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace DataAnalyzingBase
{
    /*
    주소
    https://kto.visitkorea.or.kr/kor/notice/data/statis/tstat/profit/notice/inout/popup.kto
    분류-목적별/국적별
    //*[@id="gubun_2"]
    연도-2010 2020
    //*[@id="yyyy"]
    월-1-12
    //*[@id="mm"]
    다운로드버튼
    //*[@id="popContents"]/div[2]/a[2]
    */

    class Program
    {
        static void Main(string[] args)
        {
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = false;

            string url = "https://kto.visitkorea.or.kr/kor/notice/data/statis/tstat/profit/notice/inout/popup.kto";

            var options = new ChromeOptions();
            using(var driver=new ChromeDriver(driverService, options))
            {
                driver.Navigate().GoToUrl(url);
                var dropboxCategory = driver.FindElementByXPath("//*[@id=\"gubun_2\"]");
                var dropboxYear = driver.FindElementByXPath("//*[@id=\"yyyy\"]");
                var dropboxMonth = driver.FindElementByXPath("//*[@id=\"mm\"]");
                var buttonSearch = driver.FindElementByXPath("//*[@id=\"popContents\"]/form[1]/fieldset/div/div[2]/a");
                var buttonDownload = driver.FindElementByXPath("//*[@id=\"popContents\"]/div[2]/a[2]");

                dropboxCategory.SendKeys("목적별/국적별");
                for(int year = 2010; year < 2021; year++)
                {
                    dropboxYear.SendKeys(year.ToString());
                    for(int month = 1; month < 13; month++)
                    {
                        dropboxMonth.SendKeys(month.ToString());
                        buttonSearch.Click();
                        Thread.Sleep(500);
                        buttonDownload.Click();
                        Thread.Sleep(500);

                    }
                }
                

            }

        }
    }
}
