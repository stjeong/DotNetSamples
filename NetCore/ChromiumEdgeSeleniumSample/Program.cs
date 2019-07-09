using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string platform = (IntPtr.Size == 8) ? "x64" : "x86";
        string path = Path.Combine(Path.GetDirectoryName(typeof(Program).Assembly.Location), "lib", platform);

        ChromeOptions co = new ChromeOptions();
        co.BinaryLocation = @"C:\Program Files (x86)\Microsoft\Edge Dev\Application\msedge.exe";

        ChromeDriverService cds = ChromeDriverService.CreateDefaultService(path, "msedgedriver.exe");

        using (IWebDriver driver = new ChromeDriver(cds, co))
        {
            driver.Url = "https://www.sysnet.pe.kr";

            {
                IWebElement div = driver.FindElement(By.ClassName("leftCommentArea"));
                foreach (var item in div.FindElements(By.XPath(".//a")))
                {
                    Console.WriteLine(item.Text);
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
        }
    }
}
