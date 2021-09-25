using Aspose.Cells;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;

namespace Rediff_Mini_Project
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Well-come-Rediff-project");

			//Lauch Chrome
			IWebDriver driver = new ChromeDriver("C:\\SeleniumProjects\\chromedriver_win32");

			driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3000);

			//Maximize the browser
			driver.Manage().Window.Maximize();

			//to launch the website
			driver.Url = "https://portfolio.rediff.com/portfolio-login";
			
			//send mail
			driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/input")).SendKeys("erbhushan25198@rediffmail.com");
			
			//send password
			driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/div[3]/input[1]")).SendKeys("Nagaon@123");


			//To click on sign in1
			driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/div[3]/input[2]")).Click();
			
			//To Enter the registered Email
			//driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/input")).SendKeys(emails);
			//To Enter the Password
			//driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/div[3]/input[1]")).SendKeys(password);
			// print the login to console 
			//Console.WriteLine(emails);
			//Console.WriteLine(password);

			//To click on checkbox Remember me
			//IWebElement check_box = driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/div[3]/div[2]/input"));
			//To Click on SignIn Button
			//driver.FindElement(By.Name("loginsubmit")).Click();

				//Click to add portfolio
                driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/b/div[2]/div[1]/div[1]/div/input")).Click();
		
	

				//To ReadExcel File
				string path = @"C:\SeleniumProjects\Rediff_Mini_Project.Xlsx";

            // Instantiate a Workbook object that represents Excel file.
            Workbook wb = new Workbook(path);
		
				// Access "Sheet1" from the workbook.
				Worksheet sheet = wb.Worksheets[0];


				#region First Portfolio Value

				// Access the "A1" cell in the sheet.
				// To pass the Stock name
				Cell cell11 = sheet.Cells.GetCell(1, 0);
				String value11 = cell11.Value.ToString();

				IWebElement stock1 = driver.FindElement(By.Id("addstockname_hidden"));
				stock1.SendKeys(value11);
				if (stock1.Text.Contains(value11))
				{
					driver.FindElement(By.Id("addstockname_hidden")).Click();
				}



				//To pass the date value
				Cell cell12 = sheet.Cells.GetCell(1, 1);
				String value12 = cell12.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[2]/div[1]/input")).SendKeys(value12);

				//To pass the Quantity
				Cell cell13 = sheet.Cells.GetCell(1, 2);
				var value13 = cell13.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[3]/input")).SendKeys(value13);

				//To pass the Total Amount
				Cell cell14 = sheet.Cells.GetCell(1, 3);
				var value14 = cell14.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[4]/input")).SendKeys(value14);

				// To click the Exchange Type



				//To click on Add Stock Button
				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[6]/div/input")).Click();
				Console.WriteLine("PORTFOLIO HAS BEEN CREATED");
				Console.WriteLine(value11);
				Console.WriteLine(value12);
				Console.WriteLine(value13);
				Console.WriteLine(value14);
				#endregion
				
			#region Second Portfolio value

				//// Access the "A1" cell in the sheet.
				//// To pass the Stock name
				Cell cell21 = sheet.Cells.GetCell(2, 0);
				String value21 = cell21.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[1]/input[1]")).SendKeys(value21);

				////To pass the date value
				Cell cell22 = sheet.Cells.GetCell(2, 1);
				String value22 = cell22.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[2]/div[1]/input")).SendKeys(value22);

				////To pass the Quantity
				Cell cell23 = sheet.Cells.GetCell(2, 2);
				var value23 = cell23.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[3]/input")).SendKeys(value23);

				////To pass the Total Amount
				Cell cell24 = sheet.Cells.GetCell(2, 3);
				var value24 = cell24.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[4]/input")).SendKeys(value24);

				//// To click the Exchange Type



				////To click on Add Stock Button
				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[6]/div/input")).Click();
				Console.WriteLine("PORTFOLIO HAS BEEN CREATED");
				Console.WriteLine(value21);
				Console.WriteLine(value22);
				Console.WriteLine(value23);
				Console.WriteLine(value24);
				#endregion
				#region Third Portfolio value

				//// Access the "A1" cell in the sheet.
				//// To pass the Stock name
				Cell cell31 = sheet.Cells.GetCell(3, 0);
				String value31 = cell31.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[1]/input[1]")).SendKeys(value31);

				////To pass the date value
				Cell cell32 = sheet.Cells.GetCell(3, 1);
				String value32 = cell32.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[2]/div[1]/input")).SendKeys(value32);

				////To pass the Quantity
				Cell cell33 = sheet.Cells.GetCell(3, 2);
				var value33 = cell13.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[3]/input")).SendKeys(value33);

				////To pass the Total Amount
				Cell cell34 = sheet.Cells.GetCell(3, 3);
				var value34 = cell14.Value.ToString();

				driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[4]/input")).SendKeys(value34);

				//// To click the Exchange Type



				////To click on Add Stock Button
			      driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[6]/div/input")).Click();
				Console.WriteLine("PORTFOLIO HAS BEEN CREATED");
				Console.WriteLine(value31);
				Console.WriteLine(value32);
				Console.WriteLine(value33);
				Console.WriteLine(value34);
				#endregion
				//Close the Browser
				driver.Close();
				//driver.Quit();*/
		}









	}
}

