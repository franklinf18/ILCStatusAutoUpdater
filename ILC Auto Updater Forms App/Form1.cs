using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Globalization;
using OpenQA.Selenium.Support.Extensions;
using System.Drawing.Imaging;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Xml;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using System.Management;

namespace ILC_Auto_Updater_Forms_App
{
    public partial class Form1 : Form
    {
        public IWebDriver driver { get; private set; }
        public bool addPasswordInfoSucc { get; set; }
        public string webConfigPath { get; set; }
        public string[] managerIntranetIds { get; set; }
        public string totalEmployees { get; set; }
        public int noOfRowsDelinquencyTable { get; set; }
        public int numOfColumnsDelinquencyTable { get; set; }
        public string delinquencyCountText { get; set; }
        public string totalEmployeesDefaulters { get; set; }
        public string completedPercentage { get; set; }
        public string weekNo { get; set; }
        public string completionStatus { get; set; }
        public string timeUpdated { get; set; }
        public int completedCountITSec { get; set; }
        public int iterationCount { get; set; }
        public string emailSender { get; set; }
        public string managerintranetId { get; set; }
        public string scsLink { get; set; }

        public string connectionString { get; set; }
        public string tableName { get; set; }
        public string tableNamePalas { get; set; }
        public int managerNumber { get; set; }
        public string managerName { get; set; }
        public string excelFilePath { get; set; }
        public string errorFilePath { get; set; }
        public string excelFileName { get; set; }

        public System.Data.DataSet dataSet = new DataSet("ILCDelinquency");
        public DataColumn column;
        public DataRow row;
        public DataSet ds;

        //Excel Objects
        Microsoft.Office.Interop.Excel.Application excelApp;
        Microsoft.Office.Interop.Excel.Workbook workbookILCDelinquency;
        Microsoft.Office.Interop.Excel.Worksheet worksheetILCDelinquency;


        public Form1()
        {
            InitializeComponent();
            //TestCredentialCopy();
            //TestLexicon();
            SetProperties();
            //MySQLConnector("SELECT * FROM template_arunanshu");

            UpdateCredentialsFromWebConfig();
            ProcessILCStatusUpdate();
            //MySQLConnector("SELECT * FROM template");
        }

        public bool AutoUpdateILCStatus(string intranetId, string password)
        {
            try
            {
                driver = new InternetExplorerDriver();
                driver.Manage().Window.Maximize();
                driver.Manage().Window.Position = new System.Drawing.Point(-2000, 0);
                managerintranetId = intranetId;

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60)); //you can play with the time integer  to wait for longer.

                driver.Navigate().GoToUrl("https://w3.ibm.com/services/bicentral/protect/primaV2/prima.wss");
                //driver.TakeScreenshot().SaveAsFile("screenshot.jpg", ImageFormat.Jpeg);

                IAlert alert = driver.SwitchTo().Alert();

                alert.SetAuthenticationCredentials(intranetId, password);
                alert.Accept();


                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.
                //New code to remove if block of SLM. Will check for all span tags with innertext as My Employees and click on the matching span.

                try
                {

                    //Check if BI Central has logged in to home page
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[4]/div[3]/div/div/div[1]/div[1]/div[3]/div/div[2]/div[2]/div[1]/span[3]/span[2]"))); //Check "Prima home" span

                    bool isElementPresent = IsElementPresent(By.XPath("/html/body/div[4]/div[2]/div/div/div/h2"));

                    //if H2 Element is present for Non acceptance of BI Central Terms and Conditions, then skip the user
                    if (isElementPresent)
                    {
                        driver.Close();
                        driver.Dispose();
                        LogErrorText(intranetId + " needs to accept Terms and Conditions. Automation could not proceed.");
                        return false;
                    }

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[4]/div[3]/div/div/div[1]/div[1]/div[3]/div/div[2]/div[5]/div[1]")));

                    IList<IWebElement> spanElements = driver.FindElements(By.TagName("span"));
                    for (int i = 1; i < spanElements.Count; i++)
                    {

                        if (spanElements[i].GetAttribute("innerHTML").Equals("My Employees"))
                        {
                            spanElements[i].Click();
                            ////driver.FindElement(By.Id("button-1175-btnEl")).Click();
                            //action.DoubleClick(spanElements[i]).Build().Perform();
                            break;
                        }

                    }
                }
                catch (Exception ex)
                {
                    Thread.Sleep(4000);
                    if (ex.Message.Contains("Modal dialog present"))
                    {
                        bool isElementPresent = IsElementPresent(By.XPath("/html/body/div[4]/div[2]/div/div/h1"));
                        if (isElementPresent)
                        {
                            bool isEmailSent = ValidateEmailSending(managerintranetId);
                            if (!isEmailSent)
                            {
                                SendEmailINotes(managerintranetId);
                                LogErrorText("Email Sent to " + managerintranetId);
                                return false;
                            }
                        }

                    }
                    throw;
                }


                Console.WriteLine("Employee Div clicked!!!");


                // Get Total No of Employees ====================================================================================================
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[1]/td")));
                string empNos = driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[1]/td")).Text;
                //System.out.println(empNos);
                int indexOf = empNos.LastIndexOf("of");
                int indexOfChar = empNos.LastIndexOf("<");

                //totalEmployees = empNos.Substring(indexOf - 3, indexOf - 1);
                totalEmployees = empNos.Substring(indexOf + 3, indexOfChar - 2);

                indexOfChar = totalEmployees.LastIndexOf("<");

                totalEmployees = totalEmployees.Substring(0, indexOfChar - 3);
                //totalEmployees = empNos.Substring(indexOf + 4,);

                // Get Employee Table rows and column number =====================================================================================
                IWebElement tableEmpData = driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[2]/td/table"));
                //Get number of rows in table 
                int numOfRows = tableEmpData.FindElements(By.TagName("tr")).Count;

                //Get number of columns In table.
                int numOfColumns = driver.FindElements(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).Count;

                //========================================================================================================

                // Click on Current Delinquency link ============================================================================================

                driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div/div/div[1]/ul/li[3]/a")).SendKeys(OpenQA.Selenium.Keys.Enter);    //Since Click doesnt work always

                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[4]/div[2]/div/div/div/strong")));


                try
                {
                    Thread.Sleep(2000);
                }
                catch (Exception ex)
                {
                    // TODO Auto-generated catch block
                    LogError(ex);
                    //throw ex;
                }

                bool isDisplayed = IsElementPresent(By.XPath("/html/body/div[4]/div[2]/div/div/div/h2"));


                if (isDisplayed)
                {
                    delinquencyCountText = "0";
                    totalEmployeesDefaulters = "0";
                    driver.Close();
                    driver.Dispose();
                    return true;
                }


                else
                {
                    IWebElement tableDelinquency = driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[2]/td/table"));
                    noOfRowsDelinquencyTable = tableDelinquency.FindElements(By.TagName("tr")).Count;
                    numOfColumnsDelinquencyTable = driver.FindElements(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).Count;
                    IWebElement divElement = driver.FindElement(By.Id("content-main"));
                    // add code to check if direct copy paste to excel works. ===================
                    string html = divElement.GetAttribute("innerHTML");
                    DataSet dsNew = new DataSet();
                    dsNew = ConvertHTMLTablesToDataSet(html, managerintranetId);
                    System.Data.DataTable newDt = dsNew.Tables[0];
                    dataSet.Tables.Add(newDt.Copy());

                    //SaveTextToExcel(divElement.Text,managerintranetId);
                    //SaveDataSetToExcel(dsNew);
                    // ==========================================================================
                    Console.WriteLine(noOfRowsDelinquencyTable);
                    Console.WriteLine(numOfColumnsDelinquencyTable);

                    //New code to write data to datatable ====================================================================================
                    //System.Data.DataTable table = MakeDataTable(managerintranetId);

                    //IList<IWebElement> trCollection = tableDelinquency.FindElements(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr"));

                    //foreach (IWebElement trElement in trCollection)
                    //{
                    //    DataRow row = table.NewRow();
                    //    IList<IWebElement> tdCollection = trElement.FindElements(By.XPath("td"));
                    //    for (int i = 0; i < tdCollection.Count; i++)
                    //    {
                    //        row[i] = tdCollection[i].Text;
                    //    }
                    //    table.Rows.Add(row);
                    //    //foreach (IWebElement item in tdCollection)
                    //    //{

                    //    //}
                    //}

                    //dataSet.Tables.Add(table);
                    // =======================================================================================================================
                    delinquencyCountText = driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div/div/table/tbody/tr[1]/td")).Text;

                    int indexOfNew = delinquencyCountText.LastIndexOf("of");
                    int indexOfCharNew = delinquencyCountText.LastIndexOf("<");

                    //totalEmployees = empNos.Substring(indexOf - 3, indexOf - 1);
                    totalEmployeesDefaulters = delinquencyCountText.Substring(indexOfNew + 3, indexOfCharNew - 2);

                    indexOfCharNew = totalEmployeesDefaulters.LastIndexOf("<");

                    totalEmployeesDefaulters = totalEmployeesDefaulters.Substring(0, indexOfCharNew - 3);


                ExitIteration:
                    driver.Close();
                    driver.Dispose();
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                driver.Close();
                driver.Dispose();
                return false;
                //throw;
            }

        }

        public void SendEmailINotes(string managerId)
        {
            try
            {
                //This code block sends email to manager whose Intranet password does not match with SCS provided password.
                //CloseAllBrowserWindows();
                driver = new InternetExplorerDriver();

                driver.Manage().Window.Maximize();
                string password = "";
                int weekNo = GetIso8601WeekOfYear(DateTime.Now);

                string baseURL = "https://mail.notes.na.collabserv.com/";
                driver.Navigate().GoToUrl(baseURL + "/livemail/iNotes/Mail/?OpenDocument&noredir=1");

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("Intranet_ID")));

                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver); // define Actions object




                if (System.Configuration.ConfigurationManager.AppSettings[emailSender] != null)
                {
                    //MessageBox.Show("Entered Loop");
                    string encryptedPassword = Convert.ToString(ConfigurationManager.AppSettings[emailSender]);
                    if (encryptedPassword != "" || encryptedPassword != null)
                    {
                        //MessageBox.Show("Entered 2nd Loop "+ encryptedPassword);

                    }
                    //Decryption
                    password = AesCryp.Decrypt(encryptedPassword);

                    //MessageBox.Show("Entered 2nd Loop " + password);


                    //driver.FindElement(By.Id("Intranet_ID")).Clear();
                    driver.FindElement(By.XPath("/html/body/div/div[2]/div/div[2]/div[1]/div/div/div/form/div/fieldset/p[1]/span/input")).Clear();
                    driver.FindElement(By.Id("Intranet_ID")).SendKeys(emailSender);
                    driver.FindElement(By.Id("password")).Clear();
                    driver.FindElement(By.Id("password")).SendKeys(password);
                    driver.FindElement(By.Name("ibm-submit")).Click();

                    // ERROR: Caught exception [ERROR: Unsupported command [selectFrame | s_MainFrame | ]]
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("s_MainFrame")));

                    driver.SwitchTo().Frame("s_MainFrame");

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("e-actions-mailview-inbox-new-text")));

                    driver.FindElement(By.Id("e-actions-mailview-inbox-new-text")).Click();
                    //driver.FindElement(By.Id("e-actions-mailview-inbox-new-text")).Click();

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("e-$new-0-sendto")));

                    driver.FindElement(By.Id("e-$new-0-sendto")).Clear();
                    driver.FindElement(By.Id("e-$new-0-sendto")).SendKeys(managerId + "," + emailSender); //need to test
                    driver.FindElement(By.Id("e-$new-0-subject")).Clear();
                    driver.FindElement(By.Id("e-$new-0-subject")).SendKeys("Incorrect password on Secure Credential Storage tool");
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys("Hi, \n\nYour password on Secure Credential Storage does not match with your Intranet Password. Please log in to SCS tool and update your Intranet password.\n\nLink to the tool: " + scsLink + "\n").Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();

                    IList<IWebElement> spanElements = driver.FindElements(By.TagName("span"));
                    for (int i = 1; i < spanElements.Count; i++)
                    {
                        //System.out.println("*********************************************");
                        //System.out.println(divElements.get(i).getText());
                        //Console.WriteLine(divElements[i].GetAttribute("title"));
                        if (spanElements[i].GetAttribute("innerHTML").Equals("Send"))
                        {
                            //spanElements[i].Click();
                            //driver.FindElement(By.Id("button-1175-btnEl")).Click();
                            action.DoubleClick(spanElements[i]).Build().Perform();
                            break;
                        }

                    }
                    Thread.Sleep(10000);


                    UpdateEmailSendStatus(managerId);

                }
                //Thread.Sleep(2000);

                try
                {   //Need to put this in try block as closing the browser results in a modal dialog that causes an exception
                    driver.Close();
                    driver.Dispose();
                }
                catch (Exception ex)
                {
                    driver.Close();
                    driver.Dispose();
                    //throw;
                }


            }
            catch (Exception ex)
            {
                this.LogError(ex);
                //throw;
                if (iterationCount <= 3)
                {
                    iterationCount++;
                    RestartEmailSender();

                }
                else
                {
                    driver.Close();
                    driver.Dispose();
                }
            }


        }

        public bool ValidateEmailSending(string managerId)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);

                string key = "emailsent" + managerId;
                if (ConfigurationManager.AppSettings[key] != null)
                {

                    string value = ConfigurationManager.AppSettings[key];
                    if (value == "false")
                    {
                        return false;
                    }
                    else if (value == "true")
                    {
                        return true;
                    }

                    return true;

                }

                else
                {
                    // Now do your magic..
                    //Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

                    //string encryptedPassword = EncryptString(ToSecureString(password));

                    config.AppSettings.Settings.Add(key, "false");
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    Properties.Settings.Default.Reload();
                    return false;

                }

            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }
        public bool UpdateEmailSendStatus(string managerId)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);

                string key = "emailsent" + managerId;
                if (ConfigurationManager.AppSettings[key] != null)
                {

                    string value = ConfigurationManager.AppSettings[key];
                    if (value == "false")
                    {
                        config.AppSettings.Settings.Remove(key);
                        config.AppSettings.Settings.Add(key, "true");
                        config.Save(ConfigurationSaveMode.Modified);
                        ConfigurationManager.RefreshSection("appSettings");
                        Properties.Settings.Default.Reload();

                        return true;
                    }
                    else if (value == "true")
                    {
                        return true;
                    }

                    return false;

                }

                else
                {
                    // Now do your magic..
                    //Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

                    //string encryptedPassword = EncryptString(ToSecureString(password));

                    config.AppSettings.Settings.Add(key, "true");
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    Properties.Settings.Default.Reload();
                    return true;

                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        public bool ResetEmailSendStatus(string managerId)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);

                string key = "emailsent" + managerId;
                if (ConfigurationManager.AppSettings[key] != null)
                {

                    string value = ConfigurationManager.AppSettings[key];
                    if (value == "true")
                    {
                        config.AppSettings.Settings.Remove(key);
                        config.AppSettings.Settings.Add(key, "false");
                        config.Save(ConfigurationSaveMode.Modified);
                        ConfigurationManager.RefreshSection("appSettings");
                        Properties.Settings.Default.Reload();

                        return true;
                    }
                    else if (value == "true")
                    {
                        return false;
                    }

                    return false;

                }

                return false;
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }


        public bool AddPasswordInfo(string intranetID, string password)
        {
            try
            {

                if (ConfigurationManager.AppSettings[intranetID] != null)
                {

                    Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
                    config.AppSettings.Settings.Remove(intranetID);
                    string encryptedPassword = AesCryp.Encrypt(password); // Change this line with AES.Encrypt method
                    config.AppSettings.Settings.Add(intranetID, encryptedPassword);
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    return true;



                }

                else
                {
                    // Now do your magic..
                    Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);

                    string encryptedPassword = AesCryp.Encrypt(password);

                    config.AppSettings.Settings.Add(intranetID, encryptedPassword);
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    Properties.Settings.Default.Reload();
                    return true;

                }
                //return false;
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                return false;
                //throw;
            }


        }

        public bool CopyCredentialsFromWebConfig(string path, string username)
        {
            try
            {
                if (File.Exists(path))
                {
                    ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                    map.ExeConfigFilename = path;

                    Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                    //string connStr_Con = config.ConnectionStrings.ConnectionStrings["AppDb"].ToString();
                    //config.AppSettings[]
                    //Console.WriteLine(connStr_Con);
                    //Console.WriteLine();

                    string encryptedPassword = config.AppSettings.Settings[username].Value;
                    Console.WriteLine(encryptedPassword);
                    Console.WriteLine();

                    if (encryptedPassword.Length > 70)
                    {
                        string passwordNew = DecryptString(encryptedPassword);
                    }

                    //string password = DecryptString(encryptedPassword);
                    string password = AesCryp.Decrypt(encryptedPassword);  //need to change with old password decryption method

                    addPasswordInfoSucc = AddPasswordInfo(username, password);

                    if (addPasswordInfoSucc)
                    {
                        return true;
                    }

                    else
                    {
                        return false;
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogError(ex);

                throw;
            }

        }

        public void TestCredentialCopy()
        {
            string path = @"c:\users\ibm_admin\documents\visual studio 2015\Projects\ILC Auto Updater Web App\ILC Auto Updater Web App\Web.config";
            CopyCredentialsFromWebConfig(path, "");
        }

        public void UpdateCredentialsFromWebConfig()
        {
            try
            {
                emailSender = System.Configuration.ConfigurationManager.AppSettings["emailsender"];

                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = webConfigPath;

                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);

                foreach (string intranetId in managerIntranetIds)
                {
                    if (config.AppSettings.Settings[intranetId] != null)
                    {
                        bool isCopySucc = CopyCredentialsFromWebConfig(webConfigPath, intranetId);

                        // Can add code to handle failure of copy. But later.
                    }
                    else
                    {
                        continue;
                    }
                }

                if (config.AppSettings.Settings[emailSender] != null)
                {
                    bool isCopySucc = CopyCredentialsFromWebConfig(webConfigPath, emailSender);

                    // Can add code to handle failure of copy. But later.
                }

            }
            catch (Exception ex)
            {
                LogError(ex);
                if (ex.Message.Contains("The input data is not a complete block"))
                {

                }
                throw;
            }


        }

        public void SetProperties()
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
                string managerIntranetIdsTemp = config.AppSettings.Settings["managerIntranetIDList"].Value;
                connectionString = config.AppSettings.Settings["connectionString"].Value;
                tableName = config.AppSettings.Settings["tableName"].Value;
                tableNamePalas = config.AppSettings.Settings["tableNamePalas"].Value;
                //Set managerIntranetIds string array
                managerIntranetIds = managerIntranetIdsTemp.Split(',');

                //Set webConfigPath string 
                webConfigPath = @config.AppSettings.Settings["webConfigPath"].Value;
                //New properties to set
                emailSender = System.Configuration.ConfigurationManager.AppSettings["emailsender"];
                scsLink = System.Configuration.ConfigurationManager.AppSettings["scsLink"];
                emailSender = config.AppSettings.Settings["emailSender"].Value;
                excelFilePath = config.AppSettings.Settings["excelFilePath"].Value;
                errorFilePath = config.AppSettings.Settings["errorFilePath"].Value;
                excelFileName = config.AppSettings.Settings["excelFileName"].Value;
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }


        }

        public string GetDecryptedPassword(string intranetId)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
                //config.AppSettings[]
                if (config.AppSettings.Settings[intranetId] != null)
                {
                    string encryptedPassword = Convert.ToString(config.AppSettings.Settings[intranetId].Value);
                    if (encryptedPassword != "" || encryptedPassword != null)
                    {
                        //MessageBox.Show("Entered 2nd Loop "+ encryptedPassword);

                    }
                    //Decryption
                    string password = AesCryp.Decrypt(encryptedPassword);
                    return password;
                }
                else
                {
                    return "Manager Information Not Found";
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }


        }

        public void ProcessILCStatusUpdate()
        {

            try
            {
                if (managerIntranetIds != null)
                {

                    foreach (string managerId in managerIntranetIds)
                    {

                        switch (managerId)
                        {
                            case "sedasari@in.ibm.com":
                                managerName = "Servesh Babu";
                                break;
                            case "navpitta@in.ibm.com":
                                managerName = "Naveen Pitta";
                                break;
                            case "vankired@in.ibm.com":
                                managerName = "Vijay Ankireddy";
                                break;
                            case "spaibonu@in.ibm.com":
                                managerName = "Siva Paibonu";
                                break;
                            case "vish.siram@in.ibm.com":
                                managerName = "Vishwanath Ranga Siram";
                                break;
                            case "ravinnak@in.ibm.com":
                                managerName = "Ravi Vinnakota";
                                break;
                            case "deepak.jakhotiya@in.ibm.com":
                                managerName = "Deepak Jakhotiya";
                                break;
                            case "jaypasal@in.ibm.com":
                                managerName = "Jayaram Pasala";
                                break;
                            case "pallove.pinnaka@in.ibm.com":
                                managerName = "Pallove Pinnaka";
                                break;
                            case "ipikaray@in.ibm.com":
                                managerName = "Lipika Ray";
                                break;
                            case "Abel.DSilva@in.ibm.com":
                                managerName = "Abel DSilva";
                                break;
                            case "lvemulap@in.ibm.com":
                                managerName = "Lakshmi Vemulapalli";
                                break;
                            case "sjanardh@in.ibm.com":
                                managerName = "Sushanth Dev Janardhan";
                                break;
                            case "gkasibha@in.ibm.com":
                                managerName = "Gayatri Kasibhatta";
                                break;
                            case "shisuran@in.ibm.com":
                                managerName = "Shilpa Surana";
                                break;
                            case "mahkotta@in.ibm.com":
                                managerName = "Maheshwar Kotta";
                                break;
                            case "skaligot@in.ibm.com":
                                managerName = "Srinivas Kaligotla";
                                break;

                            default:
                                break;

                        }

                        
                        string password2 = GetDecryptedPassword(managerId);

                        if (password2 != "Manager Information Not Found")
                        {
                            bool isSuccessful = AutoUpdateILCStatus(managerId, password2);
                            if (isSuccessful)
                            {
                                ComputeUpdateValues(managerName);
                                ResetEmailSendStatus(managerId);
                            }

                        }

                        else
                        {
                            LogErrorText("Manager Information Not Found for " + managerId);
                        }
                    }

                    //Data Set Part. Save Data Set to excel sheet.
                    CloseAllExcelWindows();
                    LaunchExcel();
                    excelApp.Quit();
                    Thread.Sleep(5000);
                    //CloseAllExcelWindows();
                    DeleteExcelSheets();
                    SaveDataSetToExcel(dataSet);



                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }


        }


        #region PasswordSecurity ===========================================================

        static byte[] entropy = System.Text.Encoding.Unicode.GetBytes("Salt Is Not A Password");

        public static string EncryptString(System.Security.SecureString input)
        {
            byte[] encryptedData = System.Security.Cryptography.ProtectedData.Protect(
                System.Text.Encoding.Unicode.GetBytes(ToInsecureString(input)),
                entropy,
                System.Security.Cryptography.DataProtectionScope.CurrentUser);

            return Convert.ToBase64String(encryptedData);
        }

        public static string DecryptString(string encryptedData)
        {
            try
            {
                byte[] decryptedData = System.Security.Cryptography.ProtectedData.Unprotect(
                    Convert.FromBase64String(encryptedData),
                    entropy,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);
                //return ToSecureString(System.Text.Encoding.Unicode.GetString(decryptedData));
                return System.Text.Encoding.Unicode.GetString(decryptedData);
            }
            catch
            {
                //return new SecureString();
                return "";
            }
        }

        public static SecureString ToSecureString(string input)
        {
            SecureString secure = new SecureString();
            foreach (char c in input)
            {
                secure.AppendChar(c);
            }
            secure.MakeReadOnly();
            return secure;
        }

        public static string ToInsecureString(SecureString input)
        {
            string returnValue = string.Empty;
            IntPtr ptr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input);
            try
            {
                returnValue = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr);
            }
            return returnValue;
        }

        #endregion PasswordSecurity

        private bool IsElementPresent(By by)
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

        public DataSet MySQLConnector(string query)
        {
            try
            {
                //connectionString = "Server=9.120.216.237;Database=ilc;Uid=frank;Pwd=welcome2ibm";
                MySqlConnection connection = new MySqlConnection(connectionString);
                connection.Open();

                MySqlCommand cmd = connection.CreateCommand();
                cmd.CommandText = query;
                MySqlDataAdapter adap = new MySqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                adap.Fill(ds);
                if (ds.Tables.Count > 0)
                {
                    dataGridView1.DataSource = ds.Tables[0].DefaultView;

                }
                LogErrorText("Update successful for query \n \n " + query);

                return ds;

            }
            catch (Exception ex)
            {
                LogError(ex);
                return null;
                //throw;
            }

        }

        public DataSet MySQLConnectorITSec(string query)
        {
            connectionString = "Server=9.120.216.237;Database=itsecurity;Uid=frank;Pwd=welcome2ibm";
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();

            try
            {
                MySqlCommand cmd = connection.CreateCommand();
                cmd.CommandText = query;
                MySqlDataAdapter adap = new MySqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                adap.Fill(ds);
                if (ds.Tables.Count > 0)
                {
                    dataGridView1.DataSource = ds.Tables[0].DefaultView;

                }
                LogErrorText("Update successful for query \n \n " + query);

                return ds;

            }
            catch (Exception ex)
            {
                LogError(ex);
                return null;
                //throw;
            }

        }

        public void ExtractDataSetValues(DataSet ds)
        {
            foreach (System.Data.DataTable table in ds.Tables)
            {
                foreach (DataRow row in table.Rows)
                {


                    foreach (DataColumn column in table.Columns)
                    {
                        var item = row[column];
                    }
                }
            }
        }

        public void ComputeUpdateValues(string managerName)
        {
            try
            {
                int totalEmployeesCount = Convert.ToInt16(totalEmployees);
                int totalDefaultersCount = Convert.ToInt16(totalEmployeesDefaulters);
                int totalCompletedCount = totalEmployeesCount - totalDefaultersCount;
                completedPercentage = GetPercentage(totalCompletedCount, totalEmployeesCount, 2);
                completedPercentage = completedPercentage + " %";
                weekNo = " Week " + GetIso8601WeekOfYear(DateTime.Now);

                if (totalEmployeesCount != totalCompletedCount)
                {
                    completionStatus = "Inprogress";
                }
                else
                {
                    completionStatus = "Completed";
                }

                timeUpdated = Convert.ToString(DateTime.Now);

                UpdateILCStatus(totalEmployeesCount, totalDefaultersCount, totalCompletedCount);

                //ResetCounters(); //Uncomment this line when ITSec code is removed/ commented.
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        public void ComputeUpdateValuesForITSec(string managerName)
        {
            try
            {
                int totalEmployeesCount = Convert.ToInt16(totalEmployees);
                int totalDefaultersCount = totalEmployeesCount - completedCountITSec;
                completedPercentage = GetPercentage(completedCountITSec, totalEmployeesCount, 2);
                completedPercentage = completedPercentage + " %";
                weekNo = " Week " + GetIso8601WeekOfYear(DateTime.Now);

                if (totalEmployeesCount != completedCountITSec)
                {
                    completionStatus = "Inprogress";
                }
                else
                {
                    completionStatus = "Completed";
                }

                timeUpdated = Convert.ToString(DateTime.Now);

                //UpdateILCStatus(totalEmployeesCount, totalDefaultersCount, totalCompletedCount);
                UpdateITSecStatus(totalEmployeesCount, totalDefaultersCount, completedCountITSec);
                ResetCounters();
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        public void ResetCounters()
        {
            try
            {
                totalEmployees = "";
                totalEmployeesDefaulters = "";
                completedPercentage = "";
                completedCountITSec = 0;
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        public string GetPercentage(Int32 value, Int32 total, Int32 places)
        {
            Decimal percent = 0;
            String retval = string.Empty;
            String strplaces = new String('0', places);

            if (value == 0 || total == 0)
            {
                percent = 0;
            }

            else
            {
                percent = Decimal.Divide(value, total) * 100;

                if (places > 0)
                {
                    strplaces = "." + strplaces;
                }
            }

            retval = percent.ToString("#" + strplaces);

            return retval;
        }

        public void UpdateILCStatus(int totalEmployeesCount, int totalDefaultersCount, int totalCompletedCount)
        {
            try
            {
                //managerName = "Sunder Ramboopalan"; //Comment this line
                if (managerName == "Palas Panja")
                {
                    string updateQuery = "UPDATE " + tableNamePalas + " SET Reportees='" + totalEmployeesCount + "', Completed='" + totalCompletedCount + "', Comp='" + completedPercentage + "', Pending='" + totalDefaultersCount + "', Week='" + weekNo + "',TimeUpdated ='" + timeUpdated + "',status='" + completionStatus + "' WHERE ManagerName='" + managerName + "'";
                    DataSet ds = MySQLConnector(updateQuery);
                }
                else
                {
                    string updateQuery = "UPDATE " + tableName + " SET Reportees='" + totalEmployeesCount + "', Completed='" + totalCompletedCount + "', Comp='" + completedPercentage + "', Pending='" + totalDefaultersCount + "', Week='" + weekNo + "',TimeUpdated ='" + timeUpdated + "',status='" + completionStatus + "' WHERE ManagerName='" + managerName + "'";
                    DataSet ds = MySQLConnector(updateQuery);
                }

                //string updateQueryNew = "UPDATE agiletable SET Reportees='" + totalEmployeesCount + "'  WHERE ManagerName = '" + managerName + "'";
                //DataSet dsNew = MySQLConnectorITSec(updateQueryNew);
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }


        }

        public void UpdateITSecStatus(int totalEmployeesCount, int totalDefaultersCount, int totalCompletedCount)
        {
            
            string updateQueryNew = "UPDATE agiletable SET Reportees = '" + totalEmployeesCount + "', Completed = '" + totalCompletedCount + "', Completion = '" + completedPercentage + "', Pending = '" + totalDefaultersCount + "' WHERE ManagerName = '" + managerName + "'";
            DataSet dsNew = MySQLConnectorITSec(updateQueryNew);

        }

        public int GetIso8601WeekOfYear(DateTime time)
        {
            try
            {
                // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
                // be the same week# as whatever Thursday, Friday or Saturday are,
                // and we always get those right
                DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
                if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
                {
                    time = time.AddDays(3);
                }

                // Return the week of our adjusted day
                return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                return 0;
                //throw;
            }

        }

        private void LogError(Exception ex)

        {

            string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            message += string.Format("Message: {0}", ex.Message);

            message += Environment.NewLine;

            message += string.Format("StackTrace: {0}", ex.StackTrace);

            message += Environment.NewLine;

            message += string.Format("Source: {0}", ex.Source);

            message += Environment.NewLine;

            message += string.Format("TargetSite: {0}", ex.TargetSite.ToString());

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            string path = errorFilePath;

            FileInfo fi = new FileInfo(path);
            if (!fi.Exists)
            {
                File.Create(path).Dispose();    //with Dispose() it will give error that file is being used by other process.
            }

            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }

        }

        private void LogErrorText(string messageText)

        {

            string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            message += string.Format("Message: {0}", messageText);

            message += Environment.NewLine;

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            string path = errorFilePath;

            FileInfo fi = new FileInfo(path);
            if (!fi.Exists)
            {
                File.Create(path).Dispose();    //with Dispose() it will give error that file is being used by other process.
            }

            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        public bool TestLexicon(string managerId, string password)
        {
            try
            {
                driver = new InternetExplorerDriver();
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));

                string baseURL = "https://lexicon.in.ibm.com/";
                driver.Navigate().GoToUrl(baseURL + "/Synergy/Reports/Common/Login/login.jsp");
                driver.FindElement(By.Id("login")).Clear();
                driver.FindElement(By.Id("login")).SendKeys(managerId);
                driver.FindElement(By.Id("password")).Clear();
                driver.FindElement(By.Id("password")).SendKeys(password);
                driver.FindElement(By.Id("logon_logon")).Click();
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("action:agree")));

                driver.FindElement(By.Name("action:agree")).Click();


                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ext-comp-1011-body")));

                IList<IWebElement> divElementsForITSec = driver.FindElements(By.TagName("font"));
                for (int i = 1; i < divElementsForITSec.Count; i++)
                {
                    //System.out.println("*********************************************");
                    //System.out.println(divElements.get(i).getText());
                    //Console.WriteLine(divElements[i].GetAttribute("title"));
                    if (divElementsForITSec[i].GetAttribute("innerHTML").Equals("IT Security Dairy(Q4)"))
                    {
                        divElementsForITSec[i].Click();
                        //driver.FindElement(By.Id("button-1175-btnEl")).Click();

                        break;
                    }

                }

                //wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div[3]/div/table/tbody/tr[6]/td[1]/div")));
                //driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div[3]/div/table/tbody/tr[6]/td[1]/div")).Click();

                //Select Line Manager option from the prompt

                IList<IWebElement> divElements = driver.FindElements(By.TagName("div"));
                for (int i = 1; i < divElements.Count; i++)
                {
                    //System.out.println("*********************************************");
                    //System.out.println(divElements.get(i).getText());
                    //Console.WriteLine(divElements[i].GetAttribute("title"));
                    if (divElements[i].GetAttribute("innerHTML").Equals("Line Manager"))
                    {
                        divElements[i].Click();
                        driver.FindElement(By.Id("button-1175-btnEl")).Click();

                        break;
                    }

                }

                Thread.Sleep(3000);
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div[1]/div[1]/div[2]/div[2]/div[2]/div/div[1]/div[3]/div/table")));

                IWebElement tableData = driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[2]/div[2]/div[2]/div/div[1]/div[3]/div/table"));

                // Below code introduced to change totalEmployee count as per Lexicon
                IList<IWebElement> trElementsInTable = tableData.FindElements(By.XPath("/html/body/div[1]/div[1]/div[2]/div[2]/div[2]/div/div[1]/div[3]/div/table/tbody/tr"));
                int totalEmp = trElementsInTable.Count - 2;
                totalEmployees = totalEmp.ToString();
                //=====================================================================
                IList<IWebElement> divElementsInTable = tableData.FindElements(By.TagName("div"));

                for (int i = 1; i < divElementsInTable.Count; i++)
                {
                    //System.out.println("*********************************************");
                    //System.out.println(divElements.get(i).getText());
                    //Console.WriteLine(divElements[i].GetAttribute("title"));
                    if (divElementsInTable[i].GetAttribute("innerHTML").Equals("Completed") || divElementsInTable[i].GetAttribute("innerHTML").Equals("Exception"))
                    {
                        completedCountITSec++;
                        //break;
                    }

                }

                completedCountITSec = completedCountITSec - 1;
                driver.Close();
                driver.Dispose();
                return true;

            }
            catch (Exception ex)
            {
                LogError(ex);
                driver.Close();
                driver.Dispose();
                return false;
                //throw;
            }


        }

        private void CloseAllBrowserWindows()
        {
            try
            {
                Process[] AllProcesses = Process.GetProcesses();
                foreach (var process in AllProcesses)
                {
                    if (process.MainWindowTitle != "")
                    {
                        string s = process.ProcessName.ToLower();
                        if (s == "iexplore")
                            process.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        private void CloseAllExcelWindows()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string username = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];

            try
            {
                Process[] AllProcesses = Process.GetProcesses();
                foreach (var process in AllProcesses)
                {
                    if (process.MainWindowTitle != "")
                    {

                        string s = process.ProcessName.ToLower();
                        if (s == "excel")
                        {
                            string userName2 = GetProcessOwner("EXCEL.EXE");
                            if (username == userName2)
                            {
                                process.Kill();
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        public void RestartEmailSender()
        {
            try
            {
                driver.Close();
                driver.Dispose();
                SendEmailINotes(managerintranetId);
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }
            // Add code to close existing driver and resend email

        }

        public System.Data.DataTable MakeDataTable(string managerId)
        {
            // Create a new DataTable.
            System.Data.DataTable table = new System.Data.DataTable(managerId);
            //// Declare variables for DataColumn and DataRow objects.
            //DataColumn column;
            //DataRow row;

            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.    
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Country";
            column.ReadOnly = false;
            column.Unique = false;
            // Add the Column to the DataColumnCollection.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Company";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Serial";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Last name";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "First name";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Dpt id";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Empl stat";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Lbr repid";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Qty weeks";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "First missing week";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            // Create subsequent column.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Latest missing week";
            column.AutoIncrement = false;
            column.ReadOnly = false;
            column.Unique = false;
            // Add the column to the table.
            table.Columns.Add(column);

            return table;
        }

        public void LaunchExcel()
        {
            try
            {
                //code to launch Excel file for the shift. Need to open two sheets
                FileInfo fi = new FileInfo(excelFilePath);
                bool wasFoundRunning = false;
                Microsoft.Office.Interop.Excel.Application tApp = null;
                int noOfWorkbooks = 0;

                if (fi.Exists)
                {
                    try
                    {
                        tApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        wasFoundRunning = true;
                    }
                    catch
                    {
                        // Excel is not running.
                        wasFoundRunning = false;
                        //excelApp = new Microsoft.Office.Interop.Excel.Application();
                    }
                    finally
                    {
                        if (true == wasFoundRunning)
                        {
                            excelApp = tApp;
                            //Check if SWACalloutTracker is open
                            foreach (Workbook book in excelApp.Workbooks)
                            {
                                noOfWorkbooks++;
                                if (book.Name == excelFileName)
                                {

                                    workbookILCDelinquency = excelApp.Workbooks[book.Name];
                                    worksheetILCDelinquency = (Microsoft.Office.Interop.Excel.Worksheet)workbookILCDelinquency.Sheets[1];
                                    //worksheetILCDelinquency.Name = "New Sheet";
                                    break;
                                }
                            }

                            if (workbookILCDelinquency == null || noOfWorkbooks == 0)
                            {
                                string myPath = excelFilePath;
                                excelApp.Workbooks.Open(myPath);
                                workbookILCDelinquency = excelApp.Workbooks[excelFileName];
                                worksheetILCDelinquency = (Microsoft.Office.Interop.Excel.Worksheet)workbookILCDelinquency.ActiveSheet;
                                excelApp.Visible = true;
                                SaveExcel(workbookILCDelinquency.Name);
                            }

                        }
                        else
                        {
                            excelApp = new Microsoft.Office.Interop.Excel.Application();
                            //workbooks = excelApp.Workbooks;
                            //workbook = workbooks.Add(Type.Missing);
                            //worksheet.Name = "Shift_Tickets_Routed";
                            //excelApp.Visible = true;
                            //labelError.Text = "Opening previously saved tracker";


                            string myPath = excelFilePath;
                            excelApp.Workbooks.Open(myPath);
                            workbookILCDelinquency = excelApp.Workbooks[excelFileName];
                            worksheetILCDelinquency = (Microsoft.Office.Interop.Excel.Worksheet)workbookILCDelinquency.ActiveSheet;
                            excelApp.Visible = true;
                            SaveExcel(workbookILCDelinquency.Name);
                        }

                    }

                    //string myPath = (@"C:\ILCDelinquency.xlsx");
                    //excelApp.Workbooks.Open(myPath);
                    //workbook = excelApp.Workbooks[@"C:\ILCDelinquency.xlsx"];

                    //worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    //excelApp.Visible = true;

                    //Microsoft.Office.Interop.Excel.Range last = worksheetILCDelinquency.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    //int lastRow = last.Row;

                    //if (lastRow > 1)
                    //{

                    //    MessageBox.Show("You have data saved in this tracker. Please ensure this data is backed up. It will be deleted once you login to Dash through the form", "Alert!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //}

                }

                else
                {
                    //labelError.Text = "Creating new tracker";

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.DisplayAlerts = false;
                    workbookILCDelinquency = excelApp.Workbooks.Add(Type.Missing);
                    worksheetILCDelinquency = (Microsoft.Office.Interop.Excel.Worksheet)workbookILCDelinquency.ActiveSheet;
                    //worksheetILCDelinquency.Name = "Shift_Tickets_Routed";
                    excelApp.Visible = true;
                    excelApp.DisplayAlerts = false;
                    workbookILCDelinquency.SaveAs(excelFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    excelApp.DisplayAlerts = true;

                    SaveExcel(workbookILCDelinquency.Name);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }


        }

        public void SaveExcel(string workbookName)
        {
            try
            {
                //code to save Excel File
                // handle file not found exception. Ensure only one copy is open with correct name.
                //labelError.Text = "Saving Excel Tracker";

                if (workbookName == excelFileName)
                {
                    if (IsOpen(excelFileName, excelApp))
                    {
                        excelApp.DisplayAlerts = false;
                        workbookILCDelinquency.SaveAs(excelFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        excelApp.DisplayAlerts = true;
                        //labelError.Text = "Data saved";

                    }

                    else
                    {
                        MessageBox.Show("File is already closed or cannot be found. Unable to save the file. Attempting to open and save the file", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.LaunchExcel();
                        SaveExcel(workbookILCDelinquency.Name);
                    }

                }

                //else if (true)
                //{
                //    if (IsOpen("SWAOpsQueueMonitoringTracker.xlsx", excelApp))
                //    {
                //        excelApp.DisplayAlerts = false;
                //        workbookOpsQueueMonitoring.SaveAs(@"C:\SWAOpsQueueMonitoringTracker.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //        excelApp.DisplayAlerts = true;
                //        //labelError.Text = "Data saved";
                //        ChangeMyText(labelError, "Data saved");
                //    }

                //    else
                //    {
                //        MessageBox.Show("File is already closed or cannot be found. Unable to save the file. Attempting to open and save the file", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        LaunchExcelOpsQueue();
                //        SaveExcel(workbookOpsQueueMonitoring.Name);
                //    }

                //}



            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("disconnected"))
                {
                    this.LogError(ex);
                    MessageBox.Show("File is already closed or cannot be found. Unable to save the file", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }


        }

        static bool IsOpen(string book, Microsoft.Office.Interop.Excel.Application excelApp)
        {
            foreach (_Workbook wb in excelApp.Workbooks)
            {
                if (wb.Name.Contains(book))
                {
                    return true;
                }
            }
            return false;
        }

        public void SaveDataSetToExcel(DataSet dataSet)
        {
            //worksheetILCDelinquency.Cells["A1"].LoadFromDataSet(ds, true);
            try
            {
                using (ExcelPackage pck = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ////Rename first sheet to avoid duplication error
                    //ExcelWorksheet wkSheet = (ExcelWorksheet)pck.Workbook.Worksheets[1];
                    //wkSheet.Name = "NewSheet";
                    //pck.Save(); //line added as exception was occuring on Worksheets.Add method because rename of the sheet was not reflecting.

                    if (dataSet.Tables.Count > 0)
                    {
                        foreach (System.Data.DataTable table in dataSet.Tables)
                        {
                            ExcelWorksheet ws = pck.Workbook.Worksheets.Add(table.TableName);
                            ws.Cells["A1"].LoadFromDataTable(table, true);
                            ws.Cells[ws.Dimension.Address].AutoFitColumns();

                            //XmlDocument tabXml = ws.Tables[0].TableXml;
                            //XmlNode tableNode = tabXml.ChildNodes[1];
                            //tableNode.Attributes["ref"].Value = string.Format("A1:U{0}", table.Rows.Count + 1);
                            //XmlNode autoFilterNode = tableNode.ChildNodes[0];
                            //autoFilterNode.Attributes["ref"].Value = string.Format("A1:U{0}", table.Rows.Count + 1);



                            string address = ws.Dimension.Address;
                            var modelTable = ws.Cells[address];

                            // Assign borders
                            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            var modelTable2 = ws.Cells["A1:K1"];
                            modelTable2.Style.Font.Bold = true;

                            //New LOC's to delete extra column "Total"
                            ws.Cells["L1"].Value = "";

                            pck.Save();
                        }

                        ExcelWorksheet wkSheet = (ExcelWorksheet)pck.Workbook.Worksheets[1];
                        pck.Workbook.Worksheets.Delete(wkSheet);
                        pck.Save();
                    }

                    else
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("No Data");

                        ws.Cells["A1"].Value = "No Delinquency";
                        ws.Cells["A1"].Style.Font.Bold = true;
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();

                        pck.Workbook.Worksheets.Delete((ExcelWorksheet)pck.Workbook.Worksheets[1]);

                        pck.Save();

                    }

                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }



        }

        public void SaveTextToExcel(string text, string managerId)
        {
            //worksheetILCDelinquency.Cells["A1"].LoadFromDataSet(ds, true);
            try
            {
                using (ExcelPackage pck = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    //Rename first sheet to avoid duplication error
                    ExcelWorksheet wkSheet = (ExcelWorksheet)pck.Workbook.Worksheets[1];
                    wkSheet.Name = "NewSheet";


                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add(managerId);
                    //ws.Cells["A1"].LoadFromDataTable(table, true);
                    //ws.Cells["A1"].Value = text;
                    ws.Cells["A1"].LoadFromText(text);
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    //XmlDocument tabXml = ws.Tables[0].TableXml;
                    //XmlNode tableNode = tabXml.ChildNodes[1];
                    //tableNode.Attributes["ref"].Value = string.Format("A1:U{0}", table.Rows.Count + 1);
                    //XmlNode autoFilterNode = tableNode.ChildNodes[0];
                    //autoFilterNode.Attributes["ref"].Value = string.Format("A1:U{0}", table.Rows.Count + 1);



                    string address = ws.Dimension.Address;
                    var modelTable = ws.Cells[address];

                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    var modelTable2 = ws.Cells["A1:K1"];
                    modelTable2.Style.Font.Bold = true;

                    pck.Save();


                    //ExcelWorksheet wkSheet = (ExcelWorksheet)pck.Workbook.Worksheets[1];
                    pck.Workbook.Worksheets.Delete(wkSheet);
                    pck.Save();


                    //else
                    //{
                    //    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("No Data");

                    //    ws.Cells["A1"].Value = "No Delinquency";
                    //    ws.Cells["A1"].Style.Font.Bold = true;
                    //    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    //    pck.Workbook.Worksheets.Delete((ExcelWorksheet)pck.Workbook.Worksheets[1]);

                    //    pck.Save();

                    //}

                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }



        }

        public void DeleteExcelSheets()
        {
            try
            {
                bool isFileUsed = IsFileLocked(new FileInfo(excelFilePath));
                CloseAllExcelWindows();
                using (ExcelPackage pck = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    for (int i = pck.Workbook.Worksheets.Count; i > 1; i--)
                    {
                        ExcelWorksheet wkSheet = (ExcelWorksheet)pck.Workbook.Worksheets[i];
                        pck.Workbook.Worksheets.Delete(wkSheet);
                        //wkSheet.Delete();
                        //if (wkSheet.Name == "NameOfSheetToDelete")
                        //{
                        //    wkSheet.Delete();
                        //}
                    }
                    pck.Save();
                }

                Thread.Sleep(3000);

                using (ExcelPackage pck = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet wkSheet = (ExcelWorksheet)pck.Workbook.Worksheets[1];
                    wkSheet.Name = "NewSheet";
                    pck.Save(); //line added as exception was occuring on Worksheets.Add method because rename of the sheet was not reflecting.

                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                if (ex.Message.Contains("The process cannot access the file"))
                {
                    CloseAllExcelWindows();
                    Thread.Sleep(2000);
                }
                //throw;
            }

                        
        }

        public bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        public string GetProcessOwner(string processName)
        {
            string query = "Select * from Win32_Process Where Name = \"" + processName + "\"";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection processList = searcher.Get();

            foreach (ManagementObject obj in processList)
            {
                string[] argList = new string[] { string.Empty, string.Empty };
                int returnVal = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList));
                if (returnVal == 0)
                {
                    // return DOMAIN\user
                    string owner = argList[1] + "\\" + argList[0];
                    return owner;
                }
            }

            return "NO OWNER";
        }

        private DataSet ConvertHTMLTablesToDataSet(string HTML, string managerId)
        {
            try
            {
                // Declarations   
                DataSet ds = new DataSet();
                System.Data.DataTable dt = null;
                DataRow dr = null;
                string TableExpression = "<table[^>]*>(.*?)</table>";
                string HeaderExpression = "<th[^>]*>(.*?)</th>";
                string RowExpression = "<tr[^>]*>(.*?)</tr>";
                string ColumnExpression = "<td[^>]*>(.*?)</td>";
                bool HeadersExist = false;
                int iCurrentColumn = 0;
                int iCurrentRow = 0;

                // Get a match for all the tables in the HTML   
                MatchCollection Tables = Regex.Matches(HTML, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                // Loop through each table element   
                foreach (Match Table in Tables)
                {

                    // Reset the current row counter and the header flag   
                    iCurrentRow = 0;
                    HeadersExist = false;

                    // Add a new table to the DataSet   
                    dt = new System.Data.DataTable();
                    dt.TableName = managerId;
                    // Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)   
                    if (Table.Value.Contains("<th"))
                    {
                        // Set the HeadersExist flag   
                        HeadersExist = true;

                        // Get a match for all the rows in the table   
                        MatchCollection Headers = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        // Loop through each header element   
                        foreach (Match Header in Headers)
                        {
                            if (!dt.Columns.Contains(Header.Groups[1].ToString()))
                            {
                                string columnData = RemoveAnchorTags(Header.Groups[1].ToString());
                                columnData = Regex.Replace(columnData, "<.*?>", String.Empty);
                                //dt.Columns.Add(Header.Groups[1].ToString().Replace("&nbsp;", ""));
                                dt.Columns.Add(columnData.Replace("&nbsp;", ""));

                            }
                        }
                    }
                    else
                    {
                        for (int iColumns = 1; iColumns <= Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).ToString(), RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).ToString(), ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).Count; iColumns++)
                        {
                            dt.Columns.Add("Column " + iColumns);
                        }
                    }

                    // Get a match for all the rows in the table   
                    MatchCollection Rows = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    // Loop through each row element   
                    foreach (Match Row in Rows)
                    {

                        // Only loop through the row if it isn't a header row   
                        if (!(iCurrentRow == 0 & HeadersExist == true))
                        {

                            // Create a new row and reset the current column counter   
                            dr = dt.NewRow();
                            iCurrentColumn = 0;

                            // Get a match for all the columns in the row   
                            MatchCollection Columns = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                            // Loop through each column element   
                            foreach (Match Column in Columns)
                            {
                                if (Columns.Count - 1 != iCurrentColumn)
                                // Add the value to the DataRow   
                                {
                                    string rowData = RemoveAnchorTags(Convert.ToString(Column.Groups[1]));
                                    rowData = Regex.Replace(rowData, "<.*?>", String.Empty);
                                    //dr[iCurrentColumn] = Convert.ToString(Column.Groups[1]).Replace("&nbsp;", "");
                                    dr[iCurrentColumn] = rowData.Replace("&nbsp;", "");

                                }

                                // Increase the current column    
                                iCurrentColumn += 1;
                            }

                            // Add the DataRow to the DataTable   
                            dt.Rows.Add(dr);

                        }

                        // Increase the current row counter   
                        iCurrentRow += 1;
                    }

                    // Add the DataTable to the DataSet   
                    ds.Tables.Add(dt);

                }

                return ds;
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }


        }

        public string RemoveAnchorTags(string text)
        {
            string re = @"<a [^>]+>(.*?)<\/a>";
            return (Regex.Replace(text, re, "$1"));
        }
    }
}
