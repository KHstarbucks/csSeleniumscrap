using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace scrapping
{
    class Program
    {
        static void Main()
        {
            DateTime today = DateTime.Now.Date;

            string startDate = today.AddDays(-5).ToString("yyyy/MM/dd").Replace("-", "/");
            string endDate = today.ToString("yyyy/MM/dd").Replace("-", "/");
            string keyword = "RPA";
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.g2b.go.kr/index.jsp");

            //검색 조건 집어넣기
            SetSearchCriteria(driver, keyword, startDate, endDate);

            //엑셀 설정
            //excelPath는 장표파일의 위치(바탕화면), 바탕화면의 위치는 사용자 이름에 따라 약간씩 다르므로 수정요함
            string excelPath = @"C:\Users\tjftm\OneDrive\Desktop\장표.xlsx";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelPath, ReadOnly: false);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            int lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row + 1;

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40));
            driver.SwitchTo().Frame("sub");
            driver.SwitchTo().Frame("main");

            ExtractDataAndSaveToExcel(driver, worksheet, lastRow);
            lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row + 1;

            //더보기 버튼이 있을 경우 클릭 후 작업 수행
            try 
            {
                while (IsElementPresent(driver, By.ClassName("default")))
                {
                    driver.FindElement(By.ClassName("default")).Click();
                    ExtractDataAndSaveToExcel(driver, worksheet, lastRow);
                    lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row + 1;
                }
            //더보기 버튼이 없을 경우 작업 종료
            }catch(Exception e)
            {
                Console.WriteLine("no more page",e);
            }
            finally
            {
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                driver.Quit();
            }

            

            

            
        }
        static bool IsElementPresent(IWebDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        static void SetSearchCriteria(IWebDriver driver, string searchKeyword, string startDate, string endDate)
        {
            IWebElement bidNmInput = driver.FindElement(By.Id("bidNm"));
            bidNmInput.SendKeys(searchKeyword);

            IWebElement fromBidDtInput = driver.FindElement(By.Id("fromBidDt"));
            fromBidDtInput.Clear();
            fromBidDtInput.SendKeys(startDate);

            IWebElement toBidDtInput = driver.FindElement(By.Id("toBidDt"));
            toBidDtInput.Clear();
            toBidDtInput.SendKeys(endDate);

            IWebElement searchButton = driver.FindElement(By.CssSelector("a.btn_dark"));
            searchButton.Click();
        }

        static void ExtractDataAndSaveToExcel(IWebDriver driver, Worksheet worksheet, int lastRow)
        {
            IWebElement table = driver.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table"));
            IWebElement tbody = table.FindElement(By.TagName("tbody"));
            var rows = tbody.FindElements(By.TagName("tr"));

            int rowIndex = lastRow;

            foreach (var row in rows)
            {
                var tls = row.FindElements(By.XPath(".//td[@class='tl']/div"));
                var tc = row.FindElement(By.ClassName("tc"));
                int col = 1;

                foreach (var tl in tls)
                {
                    if (col == 2)
                    {
                        worksheet.Cells[rowIndex, col].Value = "RPA";
                        col++;
                        worksheet.Cells[rowIndex, col].Value = tl.Text;
                    }
                    else if (col == 5)
                    {
                        break;
                    }
                    else
                    {
                        worksheet.Cells[rowIndex, col].Value = tl.Text;
                    }
                    col++;
                }

                worksheet.Cells[rowIndex, col].Value = tc.Text;
                worksheet.Cells[rowIndex, col + 1].Value = DateTime.Now;
                rowIndex++;
            }
        }
    }
}
