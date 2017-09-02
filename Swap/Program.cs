using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using Swap.model;
namespace Swap
{
    class Program
    {
        public static void ErrorLogging(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "certs@lightsourcehr.com", "Error log", ex.Message, strPath);
        }
        public static void ErrorLogging1(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "ba@lightsourcehr.com", "Error log", ex.Message, strPath, "data.txt");
        }
        static void Main(string[] args)
        {
            List<string> mail = new List<string>();
            TimeSpan waitTime = new TimeSpan(0, 1, 0);
            Console.WriteLine("Bot SWAP starting...");
            List<Record> records = new List<Record>();
            List<Record> withLocation = new List<Record>();
            List<Record> withoutlocation = new List<Record>();
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string[] reporturl = { "https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AEE1+Ancillary+Risk+Fees", "https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AEE2+Ancillary+Risk+Fees", "https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\WOS1+Ancillary+Risk+Fees", "https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\WOS2+Ancillary+Risk+Fees" };
            string[] reportName = { "AEE1", "AEE2", "WOS1", "WOS2" };
            int report;
            Wait:
            report = 0;
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("System on wait");
            Console.ForegroundColor = ConsoleColor.White;
            while (true)
            {
                String wait = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                if (wait == "2350")
                {
                    report = 0;
                    break;
                }
                System.Threading.Thread.Sleep(waitTime);
            }
            mail.Add("bot#1 (a.k.a SWAP) sucessfully started at " + DateTime.Now);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("System starting up...");
            Console.ForegroundColor = ConsoleColor.White;
            stepUp:
            records.Clear();
            withLocation.Clear();
            withoutlocation.Clear();
            Console.ForegroundColor = ConsoleColor.Gray;
            if (report == 0)
            {
                Console.WriteLine("AEE1 report process beings :");
            }
            else if (report == 1)
            {
                Console.WriteLine("AEE2 report process beings :");
            }
            else if (report == 2)
            {
                Console.WriteLine("WOS1 report process beings :");
            }
            else if (report == 3)
            {
                Console.WriteLine("WOS2 report process beings :");
            }
            Console.ForegroundColor = ConsoleColor.White;

            IWebDriver gc = new ChromeDriver();

            //logging in to client space
            try
            {

                gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
                gc.FindElement(By.Name("LoginID")).SendKeys("lightbot");
                gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
                gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Client space login failed");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //navigating to desire report
            try
            {
                gc.Navigate().GoToUrl(reporturl[report]);
                System.Threading.Thread.Sleep(4000);
            }
            catch
            {

            }
            //dump code section

            //gc.FindElement(By.Id("ndbfc0")).Click();
            //gc.FindElements(By.TagName("option"))[3].Click();
            //gc.FindElement(By.Id("updateBtnP")).Click();
            System.Threading.Thread.Sleep(5000);

            //....dump code section ends here

           
            //extracting data from the report
            try
            {
                Console.WriteLine("Scrapping data :");
                //getting reports from alteranating rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text);
                    records.Add(r);
                }
                //getting records from normal rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text);
                    records.Add(r);
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Data Extraction Completed !");
                Console.ForegroundColor = ConsoleColor.White;
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Data scrapping failed !");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //..extraction of data ends here

            //data filteration

            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine("Data Filteration started..!");
            Console.ForegroundColor = ConsoleColor.White;
            foreach (Record r in records)
            {
                if(r.location=="" ||r.location==null)
                {
                    withoutlocation.Add(r);
                }
                else
                {
                    withLocation.Add(r);
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Data Filteration Completed...!");
            Console.ForegroundColor = ConsoleColor.White;

            //data filteration ends here

            //Summary

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\t**** Data summaray *****");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Total Number of records found : " + records.Count);
            Console.WriteLine("Number of Importable Records: " + withLocation.Count);
            Console.WriteLine("Number of Records for cases :" + withoutlocation.Count);
            Console.WriteLine("\n\n");
            mail.Add("*****Data summaray for : "+ reportName[report]+" *****");
            mail.Add("total number of records found " +" : " + records.Count);
            mail.Add("Number of Importable Records: " + withLocation.Count);
            mail.Add("Number of Records for cases :" + withoutlocation.Count);
            //summary ends here

            //case creation

            if(withoutlocation.Count>0)
            {
                foreach(Record r in withoutlocation)
                {
                    Exception e = new Exception("Client Data without location! \n Client ID:" + r.clientID + " Case no :" + r.caseNo);
                    ErrorLogging(e);
                    try
                    {
                        gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/peo/client");
                        gc.FindElement(By.Id("dropdownMenu1")).Click();
                        gc.FindElement(By.Id("All")).Click();
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ClientNumber")).SendKeys(r.clientID.ToString()); // put the client number here.
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.ClassName("formSearchBtn")).Click();
                        System.Threading.Thread.Sleep(1000);
                        gc.FindElements(By.ClassName("cs-underline"))[1].Click();
                        gc.FindElement(By.XPath("//*[@id='customHeaderHtml']/div[2]/div[6]/div/div[1]/table/tbody/tr/td[1]/span")).Click();
                        gc.FindElement(By.XPath("//*[@id='lstDataform_btnAdd']")).Click();
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys("R");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys(Keys.Tab);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys("M");
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys(Keys.Down);
                        System.Threading.Thread.Sleep(300);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys(Keys.Enter);
                        gc.FindElement(By.XPath("//*[@id='CallerName']")).SendKeys("Lightbot");
                        gc.FindElement(By.XPath("//*[@id='EmailAddress']")).SendKeys("lightbot@lightsourcehr.com");
                        DateTime dateTime = DateTime.Today;
                        gc.FindElement(By.XPath("//*[@id='DueDate']")).SendKeys(dateTime.ToString("MM/dd/yyyy"));

                        gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys("#" + r.caseNo.ToString());

                        gc.FindElement(By.XPath("//*[@id='btnSave']")).Click();
                    }
                    catch
                    {
                        Exception ex = new Exception("Case create for Case #:" + r.caseNo + " , client #" + r.clientID + " failed");
                        mail.Add("Case create for Case #:" + r.caseNo + " , client #" + r.clientID + " failed");
                        ErrorLogging(ex);

                    }
                }
            }

            //case Creation Ends here

            //Importing data into prism
            if (withLocation.Count > 0)
            {
                //creating importable text file
                string delimiter = "\t";
                try
                {
                    using (var writer = new System.IO.StreamWriter(basePath + "\\data.txt"))
                    {
                        foreach (Record i in withLocation)
                        {
                            writer.WriteLine(i.billDate + delimiter + i.billEventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + delimiter + i.location + delimiter + delimiter + i.comment);
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Import file creation failed.");
                    mail.Add("Import file creation failed.");
                    goto End;
                }
                //....File Creation Ends here

                //logging into prism
                top:
                gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
                System.Threading.Thread.Sleep(1000);
                //logging in prism
                try
                {
                    gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                    gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
                    gc.FindElement(By.XPath("//*[@id='button8v1']")).Click();
                    System.Threading.Thread.Sleep(1000);

                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Prism login failed!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Exception e = new Exception("Prism Login failed !");
                    ErrorLogging(ex);
                    ErrorLogging(e);
                    mail.Add("Prism Login failed !");
                }

                try
                {
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).Click();
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("C");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("c");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("l");

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("i");

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("e");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("n");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("t");

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(" bill");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);
                    System.Threading.Thread.Sleep(1000);

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Enter);
                    System.Threading.Thread.Sleep(10000);

                }
                catch (Exception ex)
                {
                    mail.Add("Client Bill Pending report opening failed !");
                    Exception e = new Exception("Client Bill Pending report opening failed !");
                    ErrorLogging(ex);
                    ErrorLogging(e);
                    goto End;
                }
                try
                {
                    System.Threading.Thread.Sleep(1000);
                    gc.FindElement(By.XPath("//*[@id='button31v2']")).Click();
                    System.Threading.Thread.Sleep(1000);
                }
                catch
                {

                }
                    String current = gc.CurrentWindowHandle;
                foreach (String winHandle in gc.WindowHandles)
                {
                    gc.SwitchTo().Window(winHandle);
                }
                //sometimes the upload window doesn't open
                if (gc.CurrentWindowHandle != current)
                {
                    try
                    {
                        gc.FindElement(By.XPath("//*[@id='fname']")).SendKeys(basePath + "\\data.txt"); //put the full path of file here

                        gc.FindElement(By.XPath("//*[@id='submit1']")).Click();

                        System.Threading.Thread.Sleep(1000);
                        gc.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();
                        System.Threading.Thread.Sleep(20000);
                        gc.SwitchTo().Window(current);
                    }
                    catch
                    {
                        try
                        {
                            gc.Close();
                        }
                        catch
                        {

                        }
                        try
                        {
                            gc.SwitchTo().Window(current);
                            gc.Close();
                        }
                        catch
                        {

                        }
                        goto End;
                    }
                    try
                    {
                        Exception s = new Exception(gc.FindElement(By.XPath("//*[@id='body_span29v2']")).Text);
                       // ErrorLogging1(s);

                    }
                    catch
                    {

                    }

                    try
                    {
                        gc.FindElement(By.XPath("//*[@id='button33v2']")).Click();
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.XPath("//*[@id='button32v2']")).Click();
                        System.Threading.Thread.Sleep(2000);
                        gc.SwitchTo().Alert().Accept();
                        gc.SwitchTo().Window(current);
                        gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        gc.SwitchTo().Window(current);
                        gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();

                    }
                    tryagain:
                    try
                    {

                        gc.Close();
                        gc.Dispose();
                        Console.WriteLine("Process Complete..!");

                    }
                    catch (Exception ex)
                    {
                        Exception e = new Exception("Chrome closing failed failed !");
                        System.Threading.Thread.Sleep(10000);
                        goto tryagain;
                    }

                }
                else
                {
                    tryagain:
                    try
                    {
                        gc.Close();
                        gc.Dispose();
                        goto top;
                    }
                    catch (Exception ex)
                    {
                        Exception e = new Exception("Chrome closing failed failed !");
                        System.Threading.Thread.Sleep(10000);
                        goto tryagain;
                    }
                }

            }
            //Importing data into prism ends here
            End:
            try
            {
                gc.Close();
                gc.Dispose();

            }
            catch
            {

            }
            report++;
            if (report < 4)
            {
                goto stepUp;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkGreen ;
                Console.WriteLine("Process Completed");
                Console.ForegroundColor = ConsoleColor.White;
                Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "andreforde@usa.net", "Daily Report: "+DateTime.Now, String.Join("\n",mail));
            }
            goto Wait;
        }//main ends here 
    }
}
